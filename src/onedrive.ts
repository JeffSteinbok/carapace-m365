import { existsSync } from "node:fs";
import { mkdir, readFile, stat, writeFile } from "node:fs/promises";
import https from "node:https";
import { basename, dirname, resolve } from "node:path";
import {
  getGraphAccessToken,
  type OutlookConfig,
} from "./handlers.js";
import { isM365FeatureEnabled } from "@carapace/m365-graph-auth";

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
export const ONEDRIVE_SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024;
export const ONEDRIVE_DOWNLOAD_LIMIT = 100 * 1024 * 1024;
const FILES_READ = ["Files.Read"] as const;
const FILES_WRITE = ["Files.ReadWrite"] as const;

function filesReadScopes(config: OutlookConfig): readonly ["Files.Read"] | readonly ["Files.ReadWrite"] {
  return isM365FeatureEnabled(config.features, "onedrive-write")
    ? FILES_WRITE
    : FILES_READ;
}

interface GraphResponse {
  status: number;
  headers: Record<string, string | string[] | undefined>;
  body: Buffer;
}

interface DriveAddress {
  item_id?: string;
  path?: string;
}

interface ParentAddress {
  parent_id?: string;
  parent_path?: string;
}

interface DriveItem {
  id?: string;
  name?: string;
  size?: number;
  webUrl?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  file?: { mimeType?: string };
  folder?: { childCount?: number };
  parentReference?: { id?: string; path?: string };
  deleted?: Record<string, unknown>;
  [key: string]: unknown;
}

function request(
  method: string,
  url: string,
  token: string,
  body?: Buffer | string,
  contentType?: string,
  redirects = 0,
  maxResponseBytes?: number,
): Promise<GraphResponse> {
  return new Promise((resolveRequest, reject) => {
    let rejected = false;
    const headers: Record<string, string> = { Accept: "application/json" };
    if (token) headers.Authorization = `Bearer ${token}`;
    if (body !== undefined) {
      headers["Content-Type"] = contentType ?? "application/json";
      headers["Content-Length"] = String(Buffer.byteLength(body));
    }
    const req = https.request(url, { method, headers, timeout: 30_000 }, (res) => {
      const chunks: Buffer[] = [];
      let responseBytes = 0;
      res.on("data", (chunk: Buffer | string) => {
        if (rejected) return;
        const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
        responseBytes += buffer.length;
        if (maxResponseBytes !== undefined && responseBytes > maxResponseBytes) {
          rejected = true;
          res.destroy();
          req.destroy();
          reject(new Error(`Microsoft Graph response exceeded ${maxResponseBytes} bytes`));
          return;
        }
        chunks.push(buffer);
      });
      res.on("end", () => {
        if (rejected) return;
        const status = res.statusCode ?? 0;
        const location = res.headers?.location;
        if (status >= 300 && status < 400 && location && redirects < 5) {
          const next = new URL(location, url);
          const sameOrigin = next.origin === new URL(url).origin;
          void request(
            "GET",
            next.toString(),
            sameOrigin ? token : "",
            undefined,
            undefined,
            redirects + 1,
            maxResponseBytes,
          )
            .then(resolveRequest, reject);
          return;
        }
        resolveRequest({
          status,
          headers: res.headers ?? {},
          body: Buffer.concat(chunks),
        });
      });
      res.on("error", (error) => {
        if (!rejected) reject(error);
      });
    });
    req.on("error", (error) => {
      if (!rejected) reject(error);
    });
    req.on("timeout", () => {
      req.destroy();
      reject(new Error("Microsoft Graph request timed out"));
    });
    if (body !== undefined) req.write(body);
    req.end();
  });
}

function graphError(response: GraphResponse): Error {
  let message = "";
  try {
    const parsed = JSON.parse(response.body.toString("utf8")) as {
      error?: { code?: string; message?: string };
    };
    const code = parsed.error?.code ? `${parsed.error.code}: ` : "";
    message = `${code}${parsed.error?.message ?? ""}`;
  } catch {
    message = response.body.toString("utf8").slice(0, 300);
  }
  return new Error(`Microsoft Graph error ${response.status}${message ? `: ${message}` : ""}`);
}

async function graphJson<T>(
  method: string,
  path: string,
  token: string,
  body?: unknown,
): Promise<T> {
  const serialized = body === undefined ? undefined : JSON.stringify(body);
  const response = await request(method, `${GRAPH_BASE}${path}`, token, serialized);
  if (response.status < 200 || response.status >= 300) throw graphError(response);
  if (response.body.length === 0) return {} as T;
  try {
    return JSON.parse(response.body.toString("utf8")) as T;
  } catch {
    throw new Error(`Microsoft Graph returned invalid JSON (HTTP ${response.status})`);
  }
}

function normalizeDrivePath(value: string): string {
  const normalized = value.trim().replace(/\\/g, "/").replace(/^\/+|\/+$/g, "");
  if (!normalized) return "";
  const parts = normalized.split("/");
  if (parts.some((part) => !part || part === "." || part === ".." || part.includes("\0"))) {
    throw new Error("OneDrive paths must contain safe, non-empty path segments");
  }
  return parts.join("/");
}

function encodeDrivePath(value: string): string {
  return normalizeDrivePath(value).split("/").map(encodeURIComponent).join("/");
}

function normalizeItemName(value: string): string {
  const name = normalizeDrivePath(value);
  if (!name || name.includes("/")) {
    throw new Error("name must be a single safe file or folder name");
  }
  return name;
}

function validateAddress(address: DriveAddress, allowRoot = true): void {
  if (address.item_id?.trim() && address.path?.trim()) {
    throw new Error("Specify either item_id or path, not both");
  }
  if (address.path?.trim()) normalizeDrivePath(address.path);
  if (!allowRoot && !address.item_id?.trim() && !address.path?.trim()) {
    throw new Error("item_id or path is required");
  }
}

function validateParent(parent: ParentAddress): void {
  if (parent.parent_id?.trim() && parent.parent_path?.trim()) {
    throw new Error("Specify either parent_id or parent_path, not both");
  }
  if (parent.parent_path?.trim()) normalizeDrivePath(parent.parent_path);
}

function itemPath(address: DriveAddress): string {
  validateAddress(address);
  if (address.item_id?.trim()) {
    return `/me/drive/items/${encodeURIComponent(address.item_id.trim())}`;
  }
  const path = normalizeDrivePath(address.path ?? "");
  return path ? `/me/drive/root:/${encodeDrivePath(path)}` : "/me/drive/root";
}

function childrenPath(address: DriveAddress): string {
  validateAddress(address);
  if (address.item_id?.trim()) {
    return `/me/drive/items/${encodeURIComponent(address.item_id.trim())}/children`;
  }
  const path = normalizeDrivePath(address.path ?? "");
  return path ? `/me/drive/root:/${encodeDrivePath(path)}:/children` : "/me/drive/root/children";
}

function childItemPath(parent: ParentAddress, name: string): string {
  validateParent(parent);
  const safeName = normalizeItemName(name);
  const encodedName = encodeURIComponent(safeName);
  if (parent.parent_id?.trim()) {
    return `/me/drive/items/${encodeURIComponent(parent.parent_id.trim())}:/${encodedName}`;
  }
  const parentPath = normalizeDrivePath(parent.parent_path ?? "");
  const fullPath = parentPath ? `${encodeDrivePath(parentPath)}/${encodedName}` : encodedName;
  return `/me/drive/root:/${fullPath}`;
}

function formatItem(item: DriveItem): Record<string, unknown> {
  return {
    id: item.id ?? "",
    name: item.name ?? "",
    type: item.folder ? "folder" : "file",
    size: item.size ?? 0,
    mime_type: item.file?.mimeType ?? null,
    child_count: item.folder?.childCount ?? null,
    created: item.createdDateTime ?? "",
    last_modified: item.lastModifiedDateTime ?? "",
    web_url: item.webUrl ?? "",
    parent_id: item.parentReference?.id ?? "",
    parent_path: item.parentReference?.path ?? "",
  };
}

async function resolveItem(token: string, address: DriveAddress): Promise<DriveItem> {
  validateAddress(address, false);
  return graphJson<DriveItem>("GET", itemPath(address), token);
}

export async function listOneDrive(
  config: OutlookConfig,
  params: DriveAddress & { limit?: number },
): Promise<unknown> {
  validateAddress(params);
  const limit = params.limit ?? 50;
  if (!Number.isInteger(limit) || limit < 1 || limit > 200) {
    return { error: "limit must be an integer between 1 and 200" };
  }
  const token = await getGraphAccessToken(config, filesReadScopes(config));
  const separator = childrenPath(params).includes("?") ? "&" : "?";
  const data = await graphJson<{ value?: DriveItem[] }>(
    "GET",
    `${childrenPath(params)}${separator}$top=${limit}&$select=id,name,size,webUrl,createdDateTime,lastModifiedDateTime,file,folder,parentReference`,
    token,
  );
  const items = (data.value ?? []).map(formatItem);
  return { count: items.length, items };
}

export async function searchOneDrive(
  config: OutlookConfig,
  params: { query: string; limit?: number },
): Promise<unknown> {
  const query = params.query?.trim();
  if (!query) return { error: "query is required" };
  const limit = params.limit ?? 25;
  if (!Number.isInteger(limit) || limit < 1 || limit > 200) {
    return { error: "limit must be an integer between 1 and 200" };
  }
  const token = await getGraphAccessToken(config, filesReadScopes(config));
  const escaped = encodeURIComponent(query.replace(/'/g, "''"));
  const data = await graphJson<{ value?: DriveItem[] }>(
    "GET",
    `/me/drive/root/search(q='${escaped}')?$top=${limit}&$select=id,name,size,webUrl,createdDateTime,lastModifiedDateTime,file,folder,parentReference`,
    token,
  );
  const items = (data.value ?? []).map(formatItem);
  return { count: items.length, items };
}

export async function getOneDriveMetadata(
  config: OutlookConfig,
  params: DriveAddress,
): Promise<unknown> {
  validateAddress(params);
  const token = await getGraphAccessToken(config, filesReadScopes(config));
  const item = await graphJson<DriveItem>(
    "GET",
    `${itemPath(params)}?$select=id,name,size,webUrl,createdDateTime,lastModifiedDateTime,file,folder,parentReference`,
    token,
  );
  return formatItem(item);
}

export async function downloadOneDriveFile(
  config: OutlookConfig,
  params: DriveAddress & { output_path: string; overwrite?: boolean },
): Promise<unknown> {
  validateAddress(params, false);
  const outputPath = params.output_path?.trim();
  if (!outputPath) return { error: "output_path is required" };
  const destination = resolve(outputPath);
  if (existsSync(destination) && !params.overwrite) {
    return { error: `Local file already exists: ${destination}. Set overwrite=true to replace it.` };
  }
  const token = await getGraphAccessToken(config, filesReadScopes(config));
  const response = await request(
    "GET",
    `${GRAPH_BASE}${itemPath(params)}/content`,
    token,
    undefined,
    undefined,
    0,
    ONEDRIVE_DOWNLOAD_LIMIT,
  );
  if (response.status < 200 || response.status >= 300) throw graphError(response);
  await mkdir(dirname(destination), { recursive: true });
  await writeFile(destination, response.body, { flag: params.overwrite ? "w" : "wx" });
  return { ok: true, path: destination, size: response.body.length };
}

export async function uploadOneDriveFile(
  config: OutlookConfig,
  params: ParentAddress & {
    local_path: string;
    name?: string;
    overwrite?: boolean;
  },
): Promise<unknown> {
  validateParent(params);
  const localPath = params.local_path?.trim();
  if (!localPath) return { error: "local_path is required" };
  const source = resolve(localPath);
  const fileStat = await stat(source);
  if (!fileStat.isFile()) return { error: `local_path is not a file: ${source}` };
  if (fileStat.size > ONEDRIVE_SIMPLE_UPLOAD_LIMIT) {
    return {
      error: `File is ${fileStat.size} bytes; the simple upload limit is ${ONEDRIVE_SIMPLE_UPLOAD_LIMIT} bytes`,
    };
  }
  const name = params.name?.trim() || basename(source);
  const destinationPath = childItemPath(params, name);
  const token = await getGraphAccessToken(config, FILES_WRITE);
  if (!params.overwrite) {
    const existing = await request("GET", `${GRAPH_BASE}${destinationPath}`, token);
    if (existing.status >= 200 && existing.status < 300) {
      return { error: `A OneDrive item named '${name}' already exists. Set overwrite=true to replace it.` };
    }
    if (existing.status !== 404) throw graphError(existing);
  }
  const body = await readFile(source);
  const response = await request(
    "PUT",
    `${GRAPH_BASE}${destinationPath}:/content`,
    token,
    body,
    "application/octet-stream",
  );
  if (response.status < 200 || response.status >= 300) throw graphError(response);
  const item = JSON.parse(response.body.toString("utf8")) as DriveItem;
  return { ok: true, item: formatItem(item) };
}

export async function createOneDriveFolder(
  config: OutlookConfig,
  params: ParentAddress & { name: string },
): Promise<unknown> {
  validateParent(params);
  const name = params.name?.trim();
  if (!name) return { error: "name is required" };
  normalizeItemName(name);
  const token = await getGraphAccessToken(config, FILES_WRITE);
  const parent: DriveAddress = {
    item_id: params.parent_id,
    path: params.parent_path,
  };
  const item = await graphJson<DriveItem>("POST", childrenPath(parent), token, {
    name,
    folder: {},
    "@microsoft.graph.conflictBehavior": "fail",
  });
  return { ok: true, item: formatItem(item) };
}

export async function moveOneDriveItem(
  config: OutlookConfig,
  params: DriveAddress & ParentAddress & {
    new_name?: string;
  },
): Promise<unknown> {
  validateAddress(params, false);
  validateParent(params);
  const newName = params.new_name?.trim();
  if (!newName && !params.parent_id?.trim() && !params.parent_path?.trim()) {
    return { error: "new_name, parent_id, or parent_path is required" };
  }
  if (newName) normalizeItemName(newName);
  const token = await getGraphAccessToken(config, FILES_WRITE);
  const item = await resolveItem(token, params);
  if (!item.id) throw new Error("Microsoft Graph metadata response did not include an item ID");
  const patch: Record<string, unknown> = {};
  if (newName) patch.name = newName;
  if (params.parent_id?.trim()) {
    patch.parentReference = { id: params.parent_id.trim() };
  } else if (params.parent_path?.trim()) {
    const parent = await resolveItem(token, { path: params.parent_path });
    if (!parent.id) throw new Error("Destination folder metadata did not include an item ID");
    if (!parent.folder) return { error: "parent_path must identify a folder" };
    patch.parentReference = { id: parent.id };
  }
  const updated = await graphJson<DriveItem>(
    "PATCH",
    `/me/drive/items/${encodeURIComponent(item.id)}`,
    token,
    patch,
  );
  return { ok: true, item: formatItem(updated) };
}

export async function deleteOneDriveItem(
  config: OutlookConfig,
  params: DriveAddress,
): Promise<unknown> {
  validateAddress(params, false);
  const token = await getGraphAccessToken(config, FILES_WRITE);
  let id = params.item_id?.trim();
  if (!id) {
    const item = await resolveItem(token, params);
    id = item.id;
  }
  if (!id) throw new Error("Microsoft Graph metadata response did not include an item ID");
  await graphJson<Record<string, never>>(
    "DELETE",
    `/me/drive/items/${encodeURIComponent(id)}`,
    token,
  );
  return { ok: true, item_id: id };
}
