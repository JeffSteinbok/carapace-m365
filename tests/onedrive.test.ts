import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { EventEmitter } from "node:events";
import { existsSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import https from "node:https";
import { join, resolve } from "node:path";

interface ToolDef {
  name: string;
  execute: (id: string, params: Record<string, unknown>) => Promise<unknown>;
}

const scratch = resolve("tests", ".onedrive-test");
let tokenSequence = 0;

function makeApi(oneDriveDeniedPaths?: string[]) {
  const tools: Record<string, ToolDef> = {};
  return {
    pluginConfig: { features: ["onedrive-write"], oneDriveDeniedPaths },
    registerTool(tool: unknown) {
      tools[(tool as ToolDef).name] = tool as ToolDef;
    },
    tools,
  };
}

function resultText(result: unknown): Record<string, unknown> {
  return JSON.parse((result as { content: Array<{ text: string }> }).content[0].text) as Record<string, unknown>;
}

function mockHttpsSeq(...responses: Array<[string | Buffer, number, Record<string, string>?]>) {
  const spy = vi.spyOn(https, "request");
  const writtenBodies: string[] = [];
  for (const [body, status, headers = {}] of responses) {
    const res = new EventEmitter() as NodeJS.EventEmitter & {
      statusCode: number;
      headers: Record<string, string>;
      resume: () => void;
    };
    res.statusCode = status;
    res.headers = headers;
    res.resume = () => {};
    (res as unknown as { destroy: () => void }).destroy = vi.fn();
    const req = new EventEmitter() as NodeJS.EventEmitter & {
      destroy: () => void;
      end: () => void;
      write: () => void;
    };
    req.destroy = vi.fn();
    req.end = vi.fn();
    req.write = vi.fn((value: string | Buffer) => {
      writtenBodies.push(value.toString());
    });
    spy.mockImplementationOnce((_url, _options, callback) => {
      callback?.(res as Parameters<NonNullable<typeof callback>>[0]);
      setTimeout(() => {
        res.emit("data", Buffer.isBuffer(body) ? body : Buffer.from(body));
        res.emit("end");
      }, 0);
      return req as unknown as ReturnType<typeof https.request>;
    });
  }
  return { spy, writtenBodies };
}

async function loadTools(denied?: string[]) {
  const { createEntry } = await import("../src/index.js");
  const api = makeApi(denied);
  createEntry().register(api);
  return api.tools;
}

beforeEach(() => {
  rmSync(scratch, { recursive: true, force: true });
  delete process.env.M365_CLIENT_ID;
  delete process.env.M365_CLIENT_SECRET;
  delete process.env.M365_REFRESH_TOKEN;
  delete process.env.M365_TENANT;
  delete process.env.M365_TOKEN_BROKER_URL;
  delete process.env.M365_TOKEN_BROKER_SECRET;
  delete process.env.M365_DIRECT_TOKEN_STATE_PATH;
  delete process.env.M365_FEATURES;
  delete process.env.M365_DIRECT_TOKEN_STATE_PATH;
  delete process.env.M365_FEATURES;
  const sequence = tokenSequence++;
  process.env.M365_CLIENT_ID = "client";
  process.env.M365_REFRESH_TOKEN = `onedrive-refresh-${sequence}`;
  process.env.M365_TENANT = "consumers";
  process.env.M365_DIRECT_TOKEN_STATE_PATH = join(
    scratch,
    `direct-token-${sequence}.json`,
  );
});

afterEach(() => {
  vi.restoreAllMocks();
  rmSync(scratch, { recursive: true, force: true });
});

const TOKEN = JSON.stringify({ access_token: "drive-token", expires_in: 3600 });
const ITEM = {
  id: "item-1",
  name: "Report.txt",
  size: 12,
  webUrl: "https://onedrive.test/report",
  file: { mimeType: "text/plain" },
  parentReference: { id: "root", path: "/drive/root:" },
};

describe("OneDrive tools", () => {
  it("lists the root and searches with Files.Read", async () => {
    const { spy, writtenBodies } = mockHttpsSeq(
      [TOKEN, 200],
      [JSON.stringify({ value: [ITEM] }), 200],
      [JSON.stringify({ value: [ITEM] }), 200],
    );
    const tools = await loadTools();

    const listed = resultText(await tools.onedrive_list.execute("id", {}));
    const searched = resultText(await tools.onedrive_search.execute("id", { query: "Report" }));

    expect(listed.count).toBe(1);
    expect(searched.count).toBe(1);
    expect(writtenBodies[0]).toContain("scope=Files.Read");
    expect(writtenBodies[0]).not.toContain("offline_access");
    expect(String((spy.mock.calls[0][1] as { method?: string }).method)).toBe("POST");
    expect(String(spy.mock.calls[2][0])).toContain("search(q='Report')");
  });

  it("gets metadata using safe path encoding", async () => {
    const { spy } = mockHttpsSeq([TOKEN, 200], [JSON.stringify(ITEM), 200]);
    const tools = await loadTools();
    const metadata = resultText(await tools.onedrive_metadata.execute("id", {
      path: "Work Plans/Report #1.txt",
    }));

    expect(metadata.id).toBe("item-1");
    expect(String(spy.mock.calls[1][0])).toContain("Work%20Plans/Report%20%231.txt");
  });

  it("downloads without silently overwriting local files", async () => {
    mkdirSync(scratch, { recursive: true });
    const output = join(scratch, "download.txt");
    writeFileSync(output, "existing");
    const tools = await loadTools();

    const refused = resultText(await tools.onedrive_download.execute("id", {
      item_id: "item-1",
      output_path: output,
    }));
    expect(refused.error).toContain("already exists");

    mockHttpsSeq([TOKEN, 200], [JSON.stringify(ITEM), 200], [Buffer.from("downloaded"), 200]);
    const saved = resultText(await tools.onedrive_download.execute("id", {
      item_id: "item-1",
      output_path: output,
      overwrite: true,
    }));
    expect(saved.ok).toBe(true);
    expect(existsSync(output)).toBe(true);
  });

  it("uploads a small file and refuses remote overwrite by default", async () => {
    mkdirSync(scratch, { recursive: true });
    const source = join(scratch, "upload.txt");
    writeFileSync(source, "upload body");
    mockHttpsSeq(
      [TOKEN, 200],
      [JSON.stringify({ error: { code: "itemNotFound", message: "missing" } }), 404],
      [JSON.stringify(ITEM), 201],
    );
    const tools = await loadTools();
    const uploaded = resultText(await tools.onedrive_upload.execute("id", {
      local_path: source,
      parent_path: "Documents",
    }));

    expect(uploaded.ok).toBe(true);
    expect((uploaded.item as Record<string, unknown>).id).toBe("item-1");
  });

  it("creates, moves, renames, and deletes items", async () => {
    const folder = { id: "folder-1", name: "Archive", folder: { childCount: 0 } };
    const moved = { ...ITEM, name: "Renamed.txt", parentReference: { id: "folder-1" } };
    const { spy } = mockHttpsSeq(
      [TOKEN, 200],
      [JSON.stringify(folder), 201],
      [JSON.stringify(ITEM), 200],
      [JSON.stringify(moved), 200],
      ["", 204],
    );
    const tools = await loadTools();

    const created = resultText(await tools.onedrive_create_folder.execute("id", { name: "Archive" }));
    const movedResult = resultText(await tools.onedrive_move.execute("id", {
      item_id: "item-1",
      new_name: "Renamed.txt",
      parent_id: "folder-1",
    }));
    const deleted = resultText(await tools.onedrive_delete.execute("id", { item_id: "item-1" }));

    expect(created.ok).toBe(true);
    expect(movedResult.ok).toBe(true);
    expect(deleted.ok).toBe(true);
    expect(String(spy.mock.calls[4][0])).toContain("/me/drive/items/item-1");
  });

  it("rejects ambiguous and traversal-style addresses", async () => {
    const tools = await loadTools();
    const ambiguous = resultText(await tools.onedrive_list.execute("id", {
      item_id: "item-1",
      path: "Documents",
    }));
    const unsafe = resultText(await tools.onedrive_metadata.execute("id", {
      path: "../Secrets",
    }));

    expect(ambiguous.error).toContain("either item_id or path");
    expect(unsafe.error).toContain("safe");
  });
});

const ROOT = { id: "root", name: "", root: {}, folder: {} };
const SECRET = { id: "secret", name: "Secrets", folder: {}, parentReference: { id: "root" } };
const HIDDEN = { ...ITEM, parentReference: { id: "secret", path: "/drive/root:/Public" } };
const json = (value: unknown): [string, number] => [JSON.stringify(value), 200];
const allowed = () => [json(ITEM), json(ROOT), json(ROOT)];
const hidden = () => [json(HIDDEN), json(ROOT), json(SECRET), json(ROOT)];

describe("OneDrive exclusion policy", () => {
  it.each(["onedrive_metadata", "onedrive_read", "onedrive_download", "onedrive_list"])("blocks denied descendants via IDs in %s", async name => {
    const { spy } = mockHttpsSeq([TOKEN, 200], ...hidden());
    const tools = await loadTools(["Secrets"]);
    const result = resultText(await tools[name].execute("id", { item_id: "item-1", output_path: join(scratch, "secret.txt") }));
    expect(result.error).toContain("denied");
    expect(spy.mock.calls.some(c => String(c[0]).endsWith("/content"))).toBe(false);
    expect(existsSync(join(scratch, "secret.txt"))).toBe(false);
  });

  it.each(["/SECRETS/Report.txt", "Secrets\\Report.txt", "%53ecrets/Report.txt", "%2553ecrets/Report.txt", "Secrets/%2e%2e/Public", "Secrets%2fReport.txt"])("blocks normalized or encoded path %s", async path => {
    const { spy } = mockHttpsSeq([TOKEN, 200]);
    const tools = await loadTools(["Secrets"]);
    const result = resultText(await tools.onedrive_read.execute("id", { path }));
    expect(result.error).toBeTruthy();
    expect(spy.mock.calls).toHaveLength(1);
  });

  it.each([
    { ...ITEM, remoteItem: { id: "elsewhere" } },
    { ...ITEM, parentReference: {} },
    { ...ITEM, deleted: {} },
    { ...ITEM, name: ".." },
    { ...ITEM, name: "Secrets%2fReport.txt" },
  ])("fails closed for unverifiable metadata %#", async item => {
    mockHttpsSeq([TOKEN, 200], json(item), json(ROOT));
    const tools = await loadTools(["Secrets"]);
    expect(resultText(await tools.onedrive_metadata.execute("id", { item_id: "item-1" })).error).toBeTruthy();
  });

  it("detects ancestry cycles", async () => {
    mockHttpsSeq([TOKEN, 200], json({ ...ITEM, parentReference: { id: ITEM.id } }), json(ROOT), json(ITEM));
    const tools = await loadTools(["Secrets"]);
    expect(resultText(await tools.onedrive_read.execute("id", { item_id: ITEM.id })).error).toContain("ancestry");
  });

  it("filters search results using fresh ancestry, never search paths", async () => {
    mockHttpsSeq([TOKEN, 200], json({ value: [HIDDEN, ITEM, { id: "remote", remoteItem: {} }] }), ...hidden(), ...allowed());
    const tools = await loadTools(["Secrets"]);
    const result = resultText(await tools.onedrive_search.execute("id", { query: "Report" }));
    expect(result.count).toBe(1);
    expect(result.items).toEqual([expect.objectContaining({ id: ITEM.id })]);
  });

  it("filters children and allows similarly prefixed siblings", async () => {
    const sibling = { ...ITEM, name: "Secrets-public.txt" };
    mockHttpsSeq([TOKEN, 200], json(ROOT), json(ROOT), json({ value: [HIDDEN, sibling] }), ...hidden(), json(sibling), json(ROOT), json(ROOT));
    const tools = await loadTools(["Secrets"]);
    const result = resultText(await tools.onedrive_list.execute("id", {}));
    expect(result.count).toBe(1);
    expect(result.items).toEqual([expect.objectContaining({ name: sibling.name })]);
  });

  it.each(["onedrive_upload", "onedrive_create_folder", "onedrive_move", "onedrive_delete"])("blocks alternate write route %s before access", async name => {
    const spy = vi.spyOn(https, "request");
    const tools = await loadTools(["Secrets"]);
    expect(resultText(await tools[name].execute("id", { item_id: ITEM.id, new_name: "public.txt", name: "x", local_path: "x" })).error).toContain("writes are disabled");
    expect(spy).not.toHaveBeenCalled();
  });

  it("rechecks policy after downloading and refuses a file moved into exclusions", async () => {
    mockHttpsSeq([TOKEN, 200], ...allowed(), ["hello", 200], ...hidden());
    const tools = await loadTools(["Secrets"]);
    expect(resultText(await tools.onedrive_download.execute("id", { item_id: ITEM.id, output_path: join(scratch, "moved.txt") })).error).toContain("denied");
    expect(existsSync(join(scratch, "moved.txt"))).toBe(false);
  });

  it("supports a root deny and fails closed on malformed policy", async () => {
    mockHttpsSeq([TOKEN, 200], ...allowed());
    const tools = await loadTools(["/"]);
    expect(resultText(await tools.onedrive_metadata.execute("id", { item_id: ITEM.id })).error).toContain("denied");
    const invalid = await loadTools(["../bad"]);
    expect(resultText(await invalid.onedrive_search.execute("id", { query: "x" })).error).toContain("safe");
  });
});

describe("OneDrive bounded reads and pagination", () => {
  it("reads a bounded text window by path using the resolved ID content route", async () => {
    const { spy } = mockHttpsSeq([TOKEN, 200], json(ITEM), ["hello world", 200]);
    const tools = await loadTools();
    const result = resultText(await tools.onedrive_read.execute("id", { path: "Report.txt", offset: 1, max_chars: 4 }));
    expect(result).toMatchObject({ content: "ello", next_offset: 5, total_chars: 11 });
    expect(String(spy.mock.calls[2][0])).toBe("https://graph.microsoft.com/v1.0/me/drive/items/item-1/content");
  });

  it("downloads a path through a correctly formed ID content URL", async () => {
    const { spy } = mockHttpsSeq([TOKEN, 200], json(ITEM), ["hello", 200]);
    const tools = await loadTools();
    expect(resultText(await tools.onedrive_download.execute("id", { path: "Report.txt", output_path: join(scratch, "path.txt") })).ok).toBe(true);
    expect(String(spy.mock.calls[2][0])).toMatch(/items\/item-1\/content$/);
  });

  it.each([{ ...ITEM, name: "report.pdf" }, { ...ITEM, size: 1024 * 1024 + 1 }, { ...ITEM, file: undefined, folder: {} }])("rejects unsupported or oversized content before download %#", async item => {
    const { spy } = mockHttpsSeq([TOKEN, 200], json(item));
    const tools = await loadTools();
    expect(resultText(await tools.onedrive_read.execute("id", { item_id: ITEM.id })).error).toBeTruthy();
    expect(spy.mock.calls).toHaveLength(2);
  });

  it.each([Buffer.from([0xff]), Buffer.from("binary\0data"), Buffer.alloc(1024 * 1024 + 1, 65)])("rejects malformed, binary or over-limit bodies %#", async body => {
    mockHttpsSeq([TOKEN, 200], json(ITEM), [body, 200]);
    const tools = await loadTools();
    expect(resultText(await tools.onedrive_read.execute("id", { item_id: ITEM.id })).error).toBeTruthy();
  });

  it.each(["onedrive_list", "onedrive_search"])("continues %s with a bound cursor", async name => {
    const route = name === "onedrive_list" ? "/me/drive/root/children" : "/me/drive/root/search(q='Report')";
    const next = `https://graph.microsoft.com/v1.0${route}?$skiptoken=abc`;
    const { spy } = mockHttpsSeq([TOKEN, 200], json({ value: [], "@odata.nextLink": next }), json({ value: [ITEM] }));
    const tools = await loadTools();
    const first = resultText(await tools[name].execute("id", { query: "Report" }));
    expect(first.count).toBe(0);
    expect(first.next_cursor).toBeTypeOf("string");
    const second = resultText(await tools[name].execute("id", { query: "Report", cursor: first.next_cursor }));
    expect(second).toMatchObject({ count: 1, next_cursor: null });
    expect(String(spy.mock.calls[2][0])).toBe(next);
    for (const args of [{ cursor: first.next_cursor + "x" }, { cursor: first.next_cursor, limit: 1 }]) {
      expect(resultText(await tools[name].execute("id", { query: "Report", ...args })).error).toContain("cursor");
    }
    expect(spy.mock.calls).toHaveLength(3);
  });

  it.each(["https://evil.test/v1.0/me/drive/root/children", "https://graph.microsoft.com/v1.0/me/messages", "https://graph.microsoft.com/v1.0/me/drive/root/children#fragment"])("rejects unsafe Graph continuation %s", async next => {
    mockHttpsSeq([TOKEN, 200], json({ value: [], "@odata.nextLink": next }));
    const tools = await loadTools();
    expect(resultText(await tools.onedrive_list.execute("id", {})).error).toContain("Unsafe");
  });

  it("strips authorization on content CDN redirects", async () => {
    const { spy } = mockHttpsSeq([TOKEN, 200], json(ITEM), ["", 302, { location: "https://cdn.example.test/file" }], ["hello", 200]);
    const tools = await loadTools();
    expect(resultText(await tools.onedrive_read.execute("id", { item_id: ITEM.id })).content).toBe("hello");
    expect((spy.mock.calls[3][1] as { headers: Record<string, string> }).headers.Authorization).toBeUndefined();
  });
});

it("reads an allowed file with exclusions and verifies ancestry before and after content", async () => {
  mockHttpsSeq([TOKEN, 200], ...allowed(), ["hello", 200], ...allowed());
  const tools = await loadTools(["Secrets"]);
  expect(resultText(await tools.onedrive_read.execute("id", { item_id: ITEM.id })).content).toBe("hello");
});

it("rechecks exclusions on subsequent search pages", async () => {
  const next = "https://graph.microsoft.com/v1.0/me/drive/root/search(q='Report')?$skiptoken=next";
  mockHttpsSeq([TOKEN, 200], json({ value: [], "@odata.nextLink": next }), json({ value: [HIDDEN] }), ...hidden());
  const tools = await loadTools(["Secrets"]);
  const first = resultText(await tools.onedrive_search.execute("id", { query: "Report" }));
  const second = resultText(await tools.onedrive_search.execute("id", { query: "Report", cursor: first.next_cursor }));
  expect(second).toEqual({ count: 0, items: [], next_cursor: null });
});

it("rejects cursor reuse for a different query, tool or policy", async () => {
  const next = "https://graph.microsoft.com/v1.0/me/drive/root/search(q='Report')?$skiptoken=next";
  const { spy } = mockHttpsSeq([TOKEN, 200], json({ value: [], "@odata.nextLink": next }));
  const tools = await loadTools();
  const first = resultText(await tools.onedrive_search.execute("id", { query: "Report" }));
  expect(resultText(await tools.onedrive_search.execute("id", { query: "Other", cursor: first.next_cursor })).error).toContain("cursor");
  expect(resultText(await tools.onedrive_list.execute("id", { cursor: first.next_cursor })).error).toContain("cursor");
  const protectedTools = await loadTools(["Secrets"]);
  expect(resultText(await protectedTools.onedrive_search.execute("id", { query: "Report", cursor: first.next_cursor })).error).toContain("cursor");
  expect(spy.mock.calls).toHaveLength(2);
});

it("rejects insecure content redirects and does not follow JSON redirects", async () => {
  const { spy } = mockHttpsSeq([TOKEN, 200], json(ITEM), ["", 302, { location: "http://cdn.example.test/file" }], ["", 302, { location: "https://graph.microsoft.com/v1.0/me/messages" }]);
  const tools = await loadTools();
  expect(resultText(await tools.onedrive_read.execute("id", { item_id: ITEM.id })).error).toContain("Unsafe");
  expect(resultText(await tools.onedrive_metadata.execute("id", { item_id: ITEM.id })).error).toContain("302");
  expect(spy.mock.calls).toHaveLength(4);
});
