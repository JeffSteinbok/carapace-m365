import { createHmac, randomBytes, timingSafeEqual } from "node:crypto";
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
  parentReference?: { id?: string; path?: string; driveId?: string };
  remoteItem?: unknown;
  root?: unknown;
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
  maxResponseBytes = 4 * 1024 * 1024,
  allowRedirects = false,
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
        if (allowRedirects && status >= 300 && status < 400 && location && redirects < 5) {
          const next = new URL(location, url);
          if (next.protocol !== "https:" || next.username || next.password) {
            reject(new Error("Unsafe content redirect"));
            return;
          }
          const sameOrigin = next.origin === new URL(url).origin;
          void request(
            "GET",
            next.toString(),
            sameOrigin ? token : "",
            undefined,
            undefined,
            redirects + 1,
            maxResponseBytes,
            allowRedirects,
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

// Policy paths are deliberately conservative: decode repeatedly, normalize Unicode,
// separators and case, and reject ambiguous traversal rather than resolving it.
function policyPath(value: string): string {
  let decoded = value;
  for (let i = 0; i < 8 && /%[0-9a-f]{2}/i.test(decoded); i++) {
    decoded = decodeURIComponent(decoded);
  }
  if (/%[0-9a-f]{2}/i.test(decoded)) throw new Error("Ambiguous OneDrive path");
  return normalizeDrivePath(decoded.normalize("NFC")).toLowerCase();
}

function deniedPaths(config: OutlookConfig): string[] {
  const paths = config.oneDriveDeniedPaths;
  if (paths === undefined) return [];
  if (!Array.isArray(paths) || paths.some(p => typeof p !== "string")) {
    throw new Error("oneDriveDeniedPaths must be an array of paths");
  }
  return paths.map(policyPath);
}

function assertPathAllowed(path: string, denied: string[]): void {
  const normalized = policyPath(path);
  if (denied.some(d => !d || normalized === d || normalized.startsWith(`${d}/`))) {
    throw new Error("OneDrive access denied by path policy");
  }
}

const SELECT = "id,name,size,webUrl,createdDateTime,lastModifiedDateTime,file,folder,parentReference,remoteItem,root";

async function checkedItem(config: OutlookConfig, token: string, address: DriveAddress): Promise<DriveItem> {
  validateAddress(address);
  const denied = deniedPaths(config);
  if (denied.length && address.path !== undefined) assertPathAllowed(address.path, denied);
  const item = await graphJson<DriveItem>("GET", `${itemPath(address)}?$select=${SELECT}`, token);
  if (!item.id) throw new Error("Cannot establish OneDrive item identity");
  if (!denied.length) return item;
  // Never trust a supplied path or search parentReference.path as proof of ancestry.
  const root = await graphJson<DriveItem>("GET", "/me/drive/root?$select=id,root", token);
  if (!root.id || !root.root) throw new Error("Cannot establish OneDrive root");
  let current = item;
  const segments: string[] = [];
  const seen = new Set<string>();
  for (let depth = 0; depth < 128; depth++) {
    if (!current.id || seen.has(current.id) || current.remoteItem || current.deleted) break;
    seen.add(current.id);
    if (current.id === root.id) {
      assertPathAllowed(segments.join("/"), denied);
      return item;
    }
    if (!current.name || !current.parentReference?.id) break;
    // A Graph name must be one segment, including after policy decoding.
    const name = policyPath(current.name);
    if (!name || name.includes("/")) break;
    segments.unshift(name);
    current = await graphJson<DriveItem>("GET", `${itemPath({ item_id: current.parentReference.id })}?$select=${SELECT}`, token);
  }
  throw new Error("Cannot establish allowed OneDrive ancestry");
}

function assertWritesAllowed(config: OutlookConfig): void {
  if (deniedPaths(config).length) {
    throw new Error("OneDrive writes are disabled while oneDriveDeniedPaths is configured");
  }
}

// Authenticated, process-local cursors cannot be changed into arbitrary Graph routes.
const cursorKey = randomBytes(32);
function cursorMac(payload: string, scope: string): Buffer {
  return createHmac("sha256", cursorKey).update(scope).update("\0").update(payload).digest();
}
function encodeCursor(next: string, scope: string): string {
  const payload = Buffer.from(JSON.stringify({ next, expires: Date.now() + 3600_000 })).toString("base64url");
  return `${payload}.${cursorMac(payload, scope).toString("base64url")}`;
}
function decodeCursor(cursor: string, scope: string): string {
  if (cursor.length > 32768) throw new Error("Invalid OneDrive cursor");
  const [payload, signature, extra] = cursor.split(".");
  const actual = Buffer.from(signature ?? "", "base64url");
  const expected = cursorMac(payload ?? "", scope);
  if (!payload || extra || actual.length !== expected.length || !timingSafeEqual(actual, expected)) {
    throw new Error("Invalid OneDrive cursor; restart enumeration");
  }
  const data = JSON.parse(Buffer.from(payload, "base64url").toString("utf8")) as { next: string; expires: number };
  if (data.expires < Date.now()) throw new Error("Expired OneDrive cursor");
  return data.next;
}
function continuationPath(next: string, initial: string): string {
  const url = new URL(next);
  const base = new URL(`${GRAPH_BASE}${initial}`);
  if (url.origin !== base.origin || url.pathname !== base.pathname || url.username || url.password || url.hash) {
    throw new Error("Unsafe OneDrive continuation URL");
  }
  return url.pathname.slice("/v1.0".length) + url.search;
}

async function page(config: OutlookConfig, token: string, initial: string, cursor?: string): Promise<unknown> {
  const scope = JSON.stringify([config, initial]);
  const path = cursor ? continuationPath(decodeCursor(cursor, scope), initial) : initial;
  const data = await graphJson<{ value?: DriveItem[]; "@odata.nextLink"?: string }>("GET", path, token);
  const items: Record<string, unknown>[] = [];
  const protectedMode = deniedPaths(config).length > 0;
  for (const candidate of data.value ?? []) {
    if (!protectedMode) { items.push(formatItem(candidate)); continue; }
    try {
      if (!candidate.id || candidate.remoteItem) continue;
      const item = await checkedItem(config, token, { item_id: candidate.id });
      items.push(formatItem(item));
    } catch { /* Excluded or unverifiable entries never escape in results. */ }
  }
  const next = data["@odata.nextLink"];
  if (next) continuationPath(next, initial);
  return { count: items.length, items, next_cursor: next ? encodeCursor(next, scope) : null };
}

export async function listOneDrive(config: OutlookConfig, params: DriveAddress & { limit?: number; cursor?: string }): Promise<unknown> {
  validateAddress(params);
  const limit = params.limit ?? 50;
  if (!Number.isInteger(limit) || limit < 1 || limit > 200) return { error: "limit must be an integer between 1 and 200" };
  const denied = deniedPaths(config);
  const token = await getGraphAccessToken(config, filesReadScopes(config));
  let address: DriveAddress = params;
  if (denied.length) {
    const folder = await checkedItem(config, token, params);
    if (!folder.folder) throw new Error("Expected a OneDrive folder");
    address = { item_id: folder.id };
  }
  return page(config, token, `${childrenPath(address)}?$top=${limit}&$select=${SELECT}`, params.cursor);
}

export async function searchOneDrive(config: OutlookConfig, params: { query: string; limit?: number; cursor?: string }): Promise<unknown> {
  const query = params.query?.trim();
  if (!query) return { error: "query is required" };
  const limit = params.limit ?? 25;
  if (!Number.isInteger(limit) || limit < 1 || limit > 200) return { error: "limit must be an integer between 1 and 200" };
  deniedPaths(config);
  const token = await getGraphAccessToken(config, filesReadScopes(config));
  const escaped = encodeURIComponent(query.replace(/'/g, "''"));
  return page(config, token, `/me/drive/root/search(q='${escaped}')?$top=${limit}&$select=${SELECT}`, params.cursor);
}

export async function getOneDriveMetadata(config: OutlookConfig, params: DriveAddress): Promise<unknown> {
  validateAddress(params);
  const token = await getGraphAccessToken(config, filesReadScopes(config));
  return formatItem(await checkedItem(config, token, params));
}

async function fileContent(config: OutlookConfig, params: DriveAddress, maxBytes: number, textOnly = false): Promise<{ item: DriveItem; body: Buffer }> {
  validateAddress(params, false);
  const token = await getGraphAccessToken(config, filesReadScopes(config));
  const item = await checkedItem(config, token, params);
  if (!item.file || item.folder || item.remoteItem) throw new Error("Expected a local OneDrive file");
  if (textOnly && !/\.(txt|md|markdown|csv|tsv|json|jsonl|xml|yaml|yml|log|html|css|js|ts|py|sh|sql|ini|toml)$/i.test(item.name ?? "")) {
    throw new Error("Unsupported text format; use onedrive_download for binary documents");
  }
  if (typeof item.size !== "number" || item.size < 0 || item.size > maxBytes) throw new Error(`File size must be known and at most ${maxBytes} bytes`);
  // Use the verified ID: avoids both path substitution and the missing :/content bug.
  const response = await request("GET", `${GRAPH_BASE}${itemPath({ item_id: item.id })}/content`, token, undefined, undefined, 0, maxBytes, true);
  if (response.status < 200 || response.status >= 300) throw graphError(response);
  // Recheck after transfer to catch moves during the read before releasing bytes.
  if (deniedPaths(config).length) await checkedItem(config, token, { item_id: item.id });
  return { item, body: response.body };
}

export async function readOneDriveFile(config: OutlookConfig, params: DriveAddress & { offset?: number; max_chars?: number }): Promise<unknown> {
  const offset = params.offset ?? 0;
  const maxChars = params.max_chars ?? 16000;
  if (!Number.isSafeInteger(offset) || offset < 0 || !Number.isInteger(maxChars) || maxChars < 1 || maxChars > 100000) {
    throw new Error("offset must be a nonnegative integer; max_chars must be 1-100000");
  }
  const { item, body } = await fileContent(config, params, 1024 * 1024, true);
  const text = new TextDecoder("utf-8", { fatal: true }).decode(body);
  if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(text)) throw new Error("File contains binary control characters");
  const content = text.slice(offset, offset + maxChars);
  const next = offset + content.length;
  return { id: item.id, name: item.name, content, offset, total_chars: text.length, next_offset: next < text.length ? next : null };
}

export async function downloadOneDriveFile(config: OutlookConfig, params: DriveAddress & { output_path: string; overwrite?: boolean }): Promise<unknown> {
  validateAddress(params, false);
  const outputPath = params.output_path?.trim();
  if (!outputPath) return { error: "output_path is required" };
  const destination = resolve(outputPath);
  if (existsSync(destination) && !params.overwrite) return { error: `Local file already exists: ${destination}. Set overwrite=true to replace it.` };
  const { body } = await fileContent(config, params, ONEDRIVE_DOWNLOAD_LIMIT);
  await mkdir(dirname(destination), { recursive: true });
  await writeFile(destination, body, { flag: params.overwrite ? "w" : "wx" });
  return { ok: true, path: destination, size: body.length };
}

export async function uploadOneDriveFile(
  config: OutlookConfig,
  params: ParentAddress & {
    local_path: string;
    name?: string;
    overwrite?: boolean;
  },
): Promise<unknown> {
  assertWritesAllowed(config);
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
  assertWritesAllowed(config);
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
  assertWritesAllowed(config);
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
  assertWritesAllowed(config);
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
