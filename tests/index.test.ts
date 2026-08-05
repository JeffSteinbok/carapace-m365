import { describe, it, expect, vi, afterEach, beforeEach } from "vitest";
import https from "node:https";
import { EventEmitter } from "node:events";
import { mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, resolve } from "node:path";

const scratch = resolve("tests", ".test-state");
let directTokenStatePath = "";

afterEach(() => {
  vi.restoreAllMocks();
  rmSync(scratch, { recursive: true, force: true });
});

function mockHttpsSeq(...responses: Array<[string, number]>) {
  const spy = vi.spyOn(https, "request");
  for (const [body, status] of responses) {
    const res = new EventEmitter() as NodeJS.EventEmitter & { statusCode: number; resume: () => void };
    res.statusCode = status;
    res.resume = () => {};
    const req = new EventEmitter() as NodeJS.EventEmitter & { destroy: () => void; end: () => void; write: () => void };
    req.destroy = vi.fn(); req.end = vi.fn(); req.write = vi.fn();
    spy.mockImplementationOnce((_url, _opts, cb) => {
      if (cb) cb(res as Parameters<typeof cb>[0]);
      setTimeout(() => { res.emit("data", Buffer.from(body)); res.emit("end"); }, 0);
      return req as unknown as ReturnType<typeof https.request>;
    });
  }
  return spy;
}

interface ToolDef { name: string; parameters: { properties: Record<string, unknown> }; execute: (id: string, params: Record<string, unknown>) => Promise<unknown> }
function makeApi() {
  const tools: Record<string, ToolDef> = {};
  return { pluginConfig: {}, registerTool(t: unknown) { tools[(t as ToolDef).name] = t as ToolDef; }, tools };
}
function resultText(r: unknown) { return JSON.parse((r as { content: Array<{ text: string }> }).content[0].text); }

// Re-import fresh module each time to avoid mock state bleed between tests.
async function loadPlugin() {
  const { createEntry } = await import("../src/index.js");
  const entry = createEntry();
  const api = makeApi();
  entry.register(api);
  return { entry, api };
}

const TOKEN = JSON.stringify({ access_token: "test-token" });
const MESSAGES = JSON.stringify({
  value: [
    { id: "msg1", subject: "Hello", from: { emailAddress: { name: "Alice", address: "alice@test.com" } }, receivedDateTime: "2026-05-02T10:00:00Z", isRead: false, hasAttachments: false, bodyPreview: "Hi" },
    { id: "msg2", subject: "Re: Hello", from: { emailAddress: { name: "Bob", address: "bob@test.com" } }, receivedDateTime: "2026-05-02T09:00:00Z", isRead: true, hasAttachments: false, bodyPreview: "Thanks" },
  ],
});
const CALENDARS = JSON.stringify({ value: [{ name: "Calendar", id: "cal-1" }, { name: "Your Family", id: "fam-1" }] });
const EVENTS = JSON.stringify({
  value: [{
    id: "evt1", subject: "Team Standup",
    start: { dateTime: "2026-05-03T17:00:00Z", timeZone: "UTC" },
    end: { dateTime: "2026-05-03T17:30:00Z", timeZone: "UTC" },
    location: { displayName: "Zoom" },
    organizer: { emailAddress: { name: "Jeff", address: "jeff@test.com" } },
    attendees: [], responseStatus: { response: "accepted" }, showAs: "busy",
    body: { contentType: "html", content: "<div>Agenda item</div>" },
  }],
});
const TASK_LISTS = JSON.stringify({
  value: [
    { id: "list-1", displayName: "Tasks", isDefault: true },
    { id: "list-2", displayName: "Work" },
  ],
});
const TASKS = JSON.stringify({
  value: [
    {
      id: "task-1",
      title: "Pay rent",
      status: "notStarted",
      importance: "normal",
      createdDateTime: "2026-05-02T08:00:00Z",
      lastModifiedDateTime: "2026-05-02T08:05:00Z",
      dueDateTime: { dateTime: "2026-05-03T00:00:00", timeZone: "America/Los_Angeles" },
      reminderDateTime: { dateTime: "2026-05-02T16:00:00", timeZone: "America/Los_Angeles" },
      body: { contentType: "text", content: "May rent" },
    },
  ],
});
const TASK_CREATED = JSON.stringify({
  id: "task-new",
  title: "Buy milk",
  status: "notStarted",
  importance: "high",
  createdDateTime: "2026-05-02T08:10:00Z",
  lastModifiedDateTime: "2026-05-02T08:10:00Z",
  dueDateTime: { dateTime: "2026-05-03T00:00:00", timeZone: "America/Los_Angeles" },
  body: { contentType: "text", content: "Whole milk" },
});

let testTokenSequence = 0;

beforeEach(() => {
  delete process.env.M365_CLIENT_ID;
  delete process.env.M365_CLIENT_SECRET;
  delete process.env.M365_REFRESH_TOKEN;
  delete process.env.M365_TENANT;
  delete process.env.M365_TOKEN_BROKER_URL;
  delete process.env.M365_TOKEN_BROKER_SECRET;
  delete process.env.M365_DIRECT_TOKEN_STATE_PATH;
  process.env.OUTLOOK_CLIENT_ID = "cid";
  process.env.OUTLOOK_CLIENT_SECRET = "csec";
  process.env.OUTLOOK_REFRESH_TOKEN = `rtoken-${testTokenSequence++}`;
  process.env.OUTLOOK_TENANT = "consumers";
  directTokenStatePath = resolve(scratch, `direct-${testTokenSequence}.json`);
  process.env.OUTLOOK_DIRECT_TOKEN_STATE_PATH = directTokenStatePath;
});

// ---------------------------------------------------------------------------
// Plugin registration
// ---------------------------------------------------------------------------

describe("plugin entry", () => {
  it("has correct id and name", async () => {
    const { entry } = await loadPlugin();
    expect(entry.id).toBe("outlook");
    expect(entry.name).toBe("Microsoft 365");
  });

  it("registers all Outlook and OneDrive tools", async () => {
    const { api } = await loadPlugin();
    expect(Object.keys(api.tools).sort()).toEqual([
      "onedrive_create_folder",
      "onedrive_delete",
      "onedrive_download",
      "onedrive_list",
      "onedrive_metadata",
      "onedrive_move",
      "onedrive_search",
      "onedrive_upload",
      "outlook_calendar_fetch",
      "outlook_complete_task",
      "outlook_create_event",
      "outlook_create_task",
      "outlook_delete_event",
      "outlook_delete_task",
      "outlook_flag",
      "outlook_forward",
      "outlook_inbox",
      "outlook_meeting",
      "outlook_move",
      "outlook_query_events",
      "outlook_read",
      "outlook_reply",
      "outlook_save_attachments",
      "outlook_search",
      "outlook_send",
      "outlook_task_lists",
      "outlook_tasks",
      "outlook_update_event",
      "outlook_update_task",
    ]);
  });
});

// ---------------------------------------------------------------------------
// Mail: outlook_inbox
// ---------------------------------------------------------------------------

describe("outlook_inbox", () => {
  it("returns error when credentials missing", async () => {
    delete process.env.OUTLOOK_REFRESH_TOKEN;
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_inbox"].execute("id", {}));
    expect(data).toHaveProperty("error");
    process.env.OUTLOOK_REFRESH_TOKEN = "rtoken-restored";
  });

  it("returns inbox messages", async () => {
    mockHttpsSeq([TOKEN, 200], [MESSAGES, 200]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_inbox"].execute("id", { limit: 10 })) as Record<string, unknown>;
    expect(data.count).toBe(2);
    expect((data.messages as Array<Record<string, unknown>>)[0].subject).toBe("Hello");
  });

  it("prefers a persisted direct-mode refresh token", async () => {
    mkdirSync(dirname(directTokenStatePath), { recursive: true });
    writeFileSync(
      directTokenStatePath,
      `${JSON.stringify({ refreshToken: "persisted-refresh" })}\n`,
      "utf8",
    );
    const spy = mockHttpsSeq([TOKEN, 200], [MESSAGES, 200]);

    const { api } = await loadPlugin();
    await api.tools["outlook_inbox"].execute("id", {});

    const tokenRequest = spy.mock.results[0]?.value as {
      write: ReturnType<typeof vi.fn>;
    };
    expect(String(tokenRequest.write.mock.calls[0]?.[0])).toContain(
      "refresh_token=persisted-refresh",
    );
    expect(String(tokenRequest.write.mock.calls[0]?.[0])).not.toContain(
      process.env.OUTLOOK_REFRESH_TOKEN,
    );
  });

  it("atomically persists direct-mode refresh-token rotation", async () => {
    mockHttpsSeq([
      JSON.stringify({
        access_token: "test-token",
        expires_in: 3600,
        refresh_token: "rotated-refresh",
      }),
      200,
    ], [MESSAGES, 200]);

    const { api } = await loadPlugin();
    await api.tools["outlook_inbox"].execute("id", {});

    expect(JSON.parse(readFileSync(directTokenStatePath, "utf8"))).toEqual({
      refreshToken: "rotated-refresh",
    });
  });

  it("does not read or write direct-token state in broker mode", async () => {
    mkdirSync(dirname(directTokenStatePath), { recursive: true });
    writeFileSync(directTokenStatePath, "not valid json", "utf8");
    delete process.env.OUTLOOK_REFRESH_TOKEN;
    process.env.M365_TOKEN_BROKER_URL = "https://broker.example.test/token";
    process.env.M365_TOKEN_BROKER_SECRET = "broker-secret";
    mockHttpsSeq([TOKEN, 200], [MESSAGES, 200]);

    const { api } = await loadPlugin();
    const data = resultText(
      await api.tools["outlook_inbox"].execute("id", {}),
    ) as Record<string, unknown>;

    expect(data.count).toBe(2);
    expect(readFileSync(directTokenStatePath, "utf8")).toBe("not valid json");
  });

  it("surfaces HTTP errors", async () => {
    mockHttpsSeq([TOKEN, 200], ["Forbidden", 403]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_inbox"].execute("id", {}));
    expect(data).toHaveProperty("error");
  });
});

// ---------------------------------------------------------------------------
// Mail: outlook_search
// ---------------------------------------------------------------------------

describe("outlook_search", () => {
  it("returns search results", async () => {
    mockHttpsSeq([TOKEN, 200], [MESSAGES, 200]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_search"].execute("id", { subject: "Hello" })) as { count: number };
    expect(data.count).toBe(2);
  });
});

// ---------------------------------------------------------------------------
// Mail: outlook_read
// ---------------------------------------------------------------------------

describe("outlook_read", () => {
  it("returns error when message_id missing", async () => {
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_read"].execute("id", {}));
    expect(data).toHaveProperty("error");
  });

  it("reads a message by id", async () => {
    const msg = {
      id: "msg1", subject: "Hello",
      from: { emailAddress: { name: "Alice", address: "alice@test.com" } },
      receivedDateTime: "2026-05-02T10:00:00Z", isRead: true, hasAttachments: false,
      bodyPreview: "", body: { content: "Full body text", contentType: "text" },
    };
    mockHttpsSeq([TOKEN, 200], [JSON.stringify(msg), 200]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_read"].execute("id", { message_id: "msg1" })) as Record<string, unknown>;
    expect(data.subject).toBe("Hello");
    expect(data.body).toBe("Full body text");
  });
});

// ---------------------------------------------------------------------------
// Mail: outlook_send
// ---------------------------------------------------------------------------

describe("outlook_send", () => {
  it("returns error when to missing", async () => {
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_send"].execute("id", { subject: "Hi", body: "Test" }));
    expect(data).toHaveProperty("error");
  });

  it("sends a message and returns success", async () => {
    mockHttpsSeq([TOKEN, 200], ["", 202]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_send"].execute("id", {
      to: "octo@steinbok.net",
      subject: "Test",
      body: "Hello",
    })) as Record<string, unknown>;
    expect(data.ok).toBe(true);
    expect(data.message).toContain("octo@steinbok.net");
  });
});

// ---------------------------------------------------------------------------
// Mail: outlook_reply
// ---------------------------------------------------------------------------

describe("outlook_reply", () => {
  it("returns error when message_id missing", async () => {
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_reply"].execute("id", { body: "Thanks" }));
    expect(data).toHaveProperty("error");
  });

  it("replies to a message", async () => {
    mockHttpsSeq([TOKEN, 200], ["", 202]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_reply"].execute("id", {
      message_id: "msg1",
      body: "Thanks!",
    })) as Record<string, unknown>;
    expect(data.ok).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Calendar: outlook_calendar_fetch
// ---------------------------------------------------------------------------

describe("outlook_calendar_fetch", () => {
  it("returns error when credentials missing", async () => {
    delete process.env.OUTLOOK_REFRESH_TOKEN;
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_calendar_fetch"].execute("id", {}));
    expect(data).toHaveProperty("error");
    process.env.OUTLOOK_REFRESH_TOKEN = "rtoken-restored";
  });

  it("returns calendar events", async () => {
    mockHttpsSeq([TOKEN, 200], [CALENDARS, 200], [EVENTS, 200], [EVENTS, 200]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_calendar_fetch"].execute("id", { calendar: "all", days: 7 })) as Record<string, { events: unknown[] }>;
    // Returns { personal: { events: [...] }, family: { events: [...] } }
    expect(data.personal).toBeDefined();
    expect(Array.isArray(data.personal.events)).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Calendar: outlook_create_event
// ---------------------------------------------------------------------------

describe("outlook_create_event", () => {
  it("returns error when subject missing", async () => {
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_create_event"].execute("id", { start: "2026-08-01T10:00" }));
    expect(data).toHaveProperty("error");
  });

  it("creates an event", async () => {
    const created = { id: "evt-new", subject: "New Event", start: { dateTime: "2026-08-01T17:00:00Z", timeZone: "UTC" }, end: { dateTime: "2026-08-01T18:00:00Z", timeZone: "UTC" }, webLink: "" };
    mockHttpsSeq([TOKEN, 200], [CALENDARS, 200], [JSON.stringify(created), 201]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_create_event"].execute("id", {
      subject: "New Event",
      start: "2026-08-01T10:00",
    })) as Record<string, unknown>;
    expect(data.event_id).toBe("evt-new");
    expect(data.success).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// Calendar: outlook_meeting
// ---------------------------------------------------------------------------

describe("outlook_meeting", () => {
  it("returns error when subject missing", async () => {
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_meeting"].execute("id", { to: "a@b.com", start: "2026-08-01T10:00" }));
    expect(data).toHaveProperty("error");
  });

  it("creates a meeting and returns ok", async () => {
    const CREATED = JSON.stringify({
      id: "mtg-1", iCalUId: "uid-1", subject: "Team Sync",
      start: { dateTime: "2026-08-01T17:00:00.0000000", timeZone: "UTC" },
      end: { dateTime: "2026-08-01T17:30:00.0000000", timeZone: "UTC" },
      webLink: "https://outlook.live.com/owa/?itemid=mtg-1",
    });
    mockHttpsSeq([TOKEN, 200], [CREATED, 201]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_meeting"].execute("id", {
      to: "jeff@steinbok.net",
      subject: "Team Sync",
      start: "2026-08-01T10:00",
      duration: "30m",
    })) as Record<string, unknown>;
    expect(data.ok).toBe(true);
    expect(data.subject).toBe("Team Sync");
    expect(data.id).toBe("mtg-1");
  });
});

// ---------------------------------------------------------------------------
// Calendar: outlook_query_events
// ---------------------------------------------------------------------------

describe("outlook_query_events", () => {
  it("returns events matching text filter", async () => {
    mockHttpsSeq([TOKEN, 200], [CALENDARS, 200], [EVENTS, 200], [EVENTS, 200]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_query_events"].execute("id", { text: "Standup" })) as { count: number };
    expect(data.count).toBeGreaterThan(0);
  });
});

// ---------------------------------------------------------------------------
// Tasks: outlook_task_lists / outlook_tasks / outlook_create_task
// ---------------------------------------------------------------------------

describe("outlook tasks", () => {
  it("lists task lists", async () => {
    mockHttpsSeq([TOKEN, 200], [TASK_LISTS, 200]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_task_lists"].execute("id", {})) as { count: number };
    expect(data.count).toBe(2);
  });

  it("lists tasks", async () => {
    mockHttpsSeq([TOKEN, 200], [TASK_LISTS, 200], [TASKS, 200]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_tasks"].execute("id", { task_list: "Tasks" })) as { count: number; tasks: Array<Record<string, unknown>> };
    expect(data.count).toBe(1);
    expect(data.tasks[0].title).toBe("Pay rent");
  });

  it("creates a task", async () => {
    mockHttpsSeq([TOKEN, 200], [TASK_LISTS, 200], [TASK_CREATED, 201]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_create_task"].execute("id", {
      title: "Buy milk",
      notes: "Whole milk",
      importance: "high",
      due: "2026-05-03T00:00:00",
    })) as Record<string, unknown>;
    expect(data.ok).toBe(true);
    expect((data.task as Record<string, unknown>).title).toBe("Buy milk");
  });

  it("updates a task", async () => {
    const updated = JSON.stringify({
      id: "task-1",
      title: "Pay rent",
      status: "completed",
      importance: "high",
      createdDateTime: "2026-05-02T08:00:00Z",
      lastModifiedDateTime: "2026-05-02T08:15:00Z",
      body: { contentType: "text", content: "May rent" },
    });
    mockHttpsSeq([TOKEN, 200], [TASK_LISTS, 200], [updated, 200]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_update_task"].execute("id", {
      task_id: "task-1",
      task_list: "Tasks",
      status: "completed",
    })) as Record<string, unknown>;
    expect(data.ok).toBe(true);
  });

  it("completes a task", async () => {
    const updated = JSON.stringify({
      id: "task-1",
      title: "Pay rent",
      status: "completed",
      importance: "normal",
      createdDateTime: "2026-05-02T08:00:00Z",
      lastModifiedDateTime: "2026-05-02T08:15:00Z",
      body: { contentType: "text", content: "May rent" },
    });
    mockHttpsSeq([TOKEN, 200], [TASK_LISTS, 200], [updated, 200]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_complete_task"].execute("id", {
      task_id: "task-1",
      task_list: "Tasks",
    })) as Record<string, unknown>;
    expect(data.ok).toBe(true);
  });

  it("deletes a task", async () => {
    mockHttpsSeq([TOKEN, 200], [TASK_LISTS, 200], ["", 204]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_delete_task"].execute("id", {
      task_id: "task-1",
      task_list: "Tasks",
    })) as Record<string, unknown>;
    expect(data.ok).toBe(true);
  });
});
