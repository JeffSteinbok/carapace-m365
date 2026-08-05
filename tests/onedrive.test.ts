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

function makeApi() {
  const tools: Record<string, ToolDef> = {};
  return {
    pluginConfig: { features: ["onedrive-write"] },
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

async function loadTools() {
  const { createEntry } = await import("../src/index.js");
  const api = makeApi();
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

    mockHttpsSeq([TOKEN, 200], [Buffer.from("downloaded"), 200]);
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
