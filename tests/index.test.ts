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
function makeApi(pluginConfig: Record<string, unknown> = {}) {
  const tools: Record<string, ToolDef> = {};
  return { pluginConfig, registerTool(t: unknown) { tools[(t as ToolDef).name] = t as ToolDef; }, tools };
}
function resultText(r: unknown) { return JSON.parse((r as { content: Array<{ text: string }> }).content[0].text); }

// Re-import fresh module each time to avoid mock state bleed between tests.
async function loadPlugin(pluginConfig: Record<string, unknown> = {}) {
  const { createEntry } = await import("../src/index.js");
  const entry = createEntry();
  const api = makeApi(pluginConfig);
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
  delete process.env.M365_FEATURES;
  process.env.M365_CLIENT_ID = "cid";
  process.env.M365_CLIENT_SECRET = "csec";
  process.env.M365_REFRESH_TOKEN = `rtoken-${testTokenSequence++}`;
  process.env.M365_TENANT = "consumers";
  directTokenStatePath = resolve(scratch, `direct-${testTokenSequence}.json`);
  process.env.M365_DIRECT_TOKEN_STATE_PATH = directTokenStatePath;
});

// ---------------------------------------------------------------------------
// Plugin registration
// ---------------------------------------------------------------------------

describe("plugin entry", () => {
  it("has correct id and name", async () => {
    const { entry } = await loadPlugin();
    expect(entry.id).toBe("m365");
    expect(entry.name).toBe("Microsoft 365");
  });

  it("keeps every possible tool in the manifest contract", async () => {
    const { entry } = await loadPlugin();
    expect(entry.contracts?.tools.sort()).toEqual([
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
      "outlook_save_draft",
      "outlook_search",
      "outlook_search_store",
      "outlook_send",
      "outlook_task_lists",
      "outlook_tasks",
      "outlook_update_event",
      "outlook_update_task",
    ]);
  });

  it("preserves Outlook defaults while hiding OneDrive tools", async () => {
    const { api } = await loadPlugin();
    expect(Object.keys(api.tools).sort()).toEqual([
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
      "outlook_save_draft",
      "outlook_search",
      "outlook_search_store",
      "outlook_send",
      "outlook_task_lists",
      "outlook_tasks",
      "outlook_update_event",
      "outlook_update_task",
    ]);
  });

  it("registers only read tools for a read-only feature", async () => {
    const { api } = await loadPlugin({ features: ["onedrive-read"] });
    expect(Object.keys(api.tools).sort()).toEqual([
      "onedrive_download",
      "onedrive_list",
      "onedrive_metadata",
      "onedrive_search",
    ]);
  });

  it("registers read and write tools when a write feature is selected", async () => {
    const { api } = await loadPlugin({ features: ["onedrive-write"] });
    expect(Object.keys(api.tools).sort()).toEqual([
      "onedrive_create_folder",
      "onedrive_delete",
      "onedrive_download",
      "onedrive_list",
      "onedrive_metadata",
      "onedrive_move",
      "onedrive_search",
      "onedrive_upload",
    ]);
  });

  it("does not let mail-send expose unrelated Outlook tools", async () => {
    const { api } = await loadPlugin({ features: ["mail-send"] });
    expect(Object.keys(api.tools).sort()).toEqual([
      "outlook_forward",
      "outlook_reply",
      "outlook_send",
    ]);
  });

  it("uses M365_FEATURES and rejects unknown feature names", async () => {
    process.env.M365_FEATURES = "tasks-read";
    const { api } = await loadPlugin();
    expect(Object.keys(api.tools).sort()).toEqual([
      "outlook_task_lists",
      "outlook_tasks",
    ]);
    await expect(loadPlugin({ features: ["unknown"] })).rejects.toThrow(
      "Unknown Microsoft 365 feature(s): unknown",
    );
  });
});

// ---------------------------------------------------------------------------
// Mail: outlook_inbox
// ---------------------------------------------------------------------------

describe("outlook_inbox", () => {
  it("returns error when credentials missing", async () => {
    delete process.env.M365_REFRESH_TOKEN;
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_inbox"].execute("id", {}));
    expect(data).toHaveProperty("error");
    process.env.M365_REFRESH_TOKEN = "rtoken-restored";
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
      process.env.M365_REFRESH_TOKEN,
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
    delete process.env.M365_REFRESH_TOKEN;
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
    delete process.env.M365_REFRESH_TOKEN;
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_calendar_fetch"].execute("id", {}));
    expect(data).toHaveProperty("error");
    process.env.M365_REFRESH_TOKEN = "rtoken-restored";
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

describe("all-day and free/busy", () => {
  function lastBody(spy: ReturnType<typeof mockHttpsSeq>) {
    const req = spy.mock.results.at(-1)?.value as { write: { mock: { calls: unknown[][] } } };
    return JSON.parse(String(req.write.mock.calls.at(-1)?.[0]));
  }

  it("creates an all-day event on midnight boundaries with an exclusive end", async () => {
    const spy = mockHttpsSeq([TOKEN, 200], [CALENDARS, 200], [JSON.stringify({ id: "e1", subject: "Move-In" }), 201]);
    const { api } = await loadPlugin();
    await api.tools["outlook_create_event"].execute("id", {
      subject: "Move-In", start: "2027-09-19", is_all_day: true, calendar: "family",
    });
    const body = lastBody(spy);
    expect(body.isAllDay).toBe(true);
    expect(body.start.dateTime).toBe("2027-09-19T00:00:00");
    // Exclusive end: the midnight after the single day covered.
    expect(body.end.dateTime).toBe("2027-09-20T00:00:00");
    // All-day events default to Free, matching Outlook's own behaviour.
    expect(body.showAs).toBe("free");
  });

  it("treats all-day end as the last day covered", async () => {
    const spy = mockHttpsSeq([TOKEN, 200], [CALENDARS, 200], [JSON.stringify({ id: "e1" }), 201]);
    const { api } = await loadPlugin();
    await api.tools["outlook_create_event"].execute("id", {
      subject: "Orientation", start: "2027-09-19", end: "2027-09-21", is_all_day: true,
    });
    const body = lastBody(spy);
    expect(body.start.dateTime).toBe("2027-09-19T00:00:00");
    expect(body.end.dateTime).toBe("2027-09-22T00:00:00");
  });

  it("honours an explicit show_as on an all-day event", async () => {
    const spy = mockHttpsSeq([TOKEN, 200], [CALENDARS, 200], [JSON.stringify({ id: "e1" }), 201]);
    const { api } = await loadPlugin();
    await api.tools["outlook_create_event"].execute("id", {
      subject: "Conference", start: "2027-09-19", is_all_day: true, show_as: "oof",
    });
    expect(lastBody(spy).showAs).toBe("oof");
  });

  it("sets free/busy without cancelling the event", async () => {
    const spy = mockHttpsSeq([TOKEN, 200], [JSON.stringify({ id: "e1" }), 200]);
    const { api } = await loadPlugin();
    await api.tools["outlook_update_event"].execute("id", { event_id: "e1", show_as: "free" });
    const body = lastBody(spy);
    expect(body.showAs).toBe("free");
    expect(body.isCancelled).toBeUndefined();
  });

  it("preserves the span when converting a multi-day event to all-day", async () => {
    // 09-19T00:00 -> 09-23T00:00 covers 19..22 (end midnight is exclusive).
    const current = JSON.stringify({ start: { dateTime: "2027-09-19T00:00:00" }, end: { dateTime: "2027-09-23T00:00:00" } });
    const spy = mockHttpsSeq([TOKEN, 200], [current, 200], [JSON.stringify({ id: "e1" }), 200]);
    const { api } = await loadPlugin();
    await api.tools["outlook_update_event"].execute("id", { event_id: "e1", is_all_day: true });
    const req = spy.mock.results.at(-1)?.value as { write: { mock: { calls: unknown[][] } } };
    const body = JSON.parse(String(req.write.mock.calls.at(-1)?.[0]));
    expect(body.start.dateTime).toBe("2027-09-19T00:00:00");
    expect(body.end.dateTime).toBe("2027-09-23T00:00:00");
  });

  it("reads the current event in the calendar timezone before converting", async () => {
    // Without Prefer, Graph answers in UTC and the exclusive-end correction
    // silently shifts multi-day events one day later.
    const current = JSON.stringify({ start: { dateTime: "2027-09-19T00:00:00" }, end: { dateTime: "2027-09-23T00:00:00" } });
    const spy = mockHttpsSeq([TOKEN, 200], [current, 200], [JSON.stringify({ id: "e1" }), 200]);
    const { api } = await loadPlugin();
    await api.tools["outlook_update_event"].execute("id", { event_id: "e1", is_all_day: true });
    const readOpts = spy.mock.calls[1]?.[1] as { headers: Record<string, string> };
    expect(readOpts.headers.Prefer).toBe('outlook.timezone="America/Los_Angeles"');
  });

  it("converts an existing timed event to all-day using its current date", async () => {
    const current = JSON.stringify({ start: { dateTime: "2027-09-19T00:00:00" }, end: { dateTime: "2027-09-20T00:00:00" } });
    const spy = mockHttpsSeq([TOKEN, 200], [current, 200], [JSON.stringify({ id: "e1" }), 200]);
    const { api } = await loadPlugin();
    await api.tools["outlook_update_event"].execute("id", { event_id: "e1", is_all_day: true });
    const body = lastBody(spy);
    expect(body.isAllDay).toBe(true);
    expect(body.start.dateTime).toBe("2027-09-19T00:00:00");
    expect(body.end.dateTime).toBe("2027-09-20T00:00:00");
    expect(body.showAs).toBe("free");
  });
});

// ---------------------------------------------------------------------------
// Calendar: outlook_query_events
// ---------------------------------------------------------------------------

describe("outlook_query_events", () => {
  it("returns events matching text filter", async () => {
    // Mock order matters: token, then the /me/events response. (This previously
    // fed the calendar-folder list in as events, so the text filter "passed"
    // on two "No subject" entries.)
    mockHttpsSeq([TOKEN, 200], [EVENTS, 200]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_query_events"].execute("id", { text: "Standup" })) as { count: number };
    expect(data.count).toBeGreaterThan(0);
  });

  it("excludes events that do not match the text filter", async () => {
    mockHttpsSeq([TOKEN, 200], [EVENTS, 200]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_query_events"].execute("id", { text: "Nonexistent" })) as { count: number };
    expect(data.count).toBe(0);
  });

  // A date range must use calendarView, not /me/events: calendarView expands
  // recurring series into occurrences and returns events overlapping the
  // window, so all-day and multi-day events are not dropped.
  it("queries a date range via calendarView with an inclusive end day", async () => {
    const spy = mockHttpsSeq([TOKEN, 200], [EVENTS, 200]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_query_events"].execute("id", {
      after: "2026-05-01", before: "2026-05-03",
    })) as { start_date: string; end_date: string; count: number };

    const url = String(spy.mock.calls.at(-1)?.[0]);
    expect(url).toContain("/me/calendarView");
    expect(url).not.toContain("/me/events");
    expect(decodeURIComponent(url)).toContain("startDateTime=2026-05-01T00:00:00");
    // `before` is inclusive, so the half-open window ends at the next midnight.
    expect(decodeURIComponent(url)).toContain("endDateTime=2026-05-04T00:00:00");
    expect(data.start_date).toBe("2026-05-01");
    expect(data.end_date).toBe("2026-05-03");
    expect(data.count).toBe(1);
  });

  it("treats after === before as a single-day query", async () => {
    const spy = mockHttpsSeq([TOKEN, 200], [EVENTS, 200]);
    const { api } = await loadPlugin();
    await api.tools["outlook_query_events"].execute("id", { after: "2026-05-03", before: "2026-05-03" });
    const url = decodeURIComponent(String(spy.mock.calls.at(-1)?.[0]));
    expect(url).toContain("startDateTime=2026-05-03T00:00:00");
    expect(url).toContain("endDateTime=2026-05-04T00:00:00");
  });

  it("requests results in the calendar timezone", async () => {
    const spy = mockHttpsSeq([TOKEN, 200], [EVENTS, 200]);
    const { api } = await loadPlugin();
    await api.tools["outlook_query_events"].execute("id", { after: "2026-05-01" });
    const opts = spy.mock.calls.at(-1)?.[1] as { headers: Record<string, string> };
    expect(opts.headers.Prefer).toBe('outlook.timezone="America/Los_Angeles"');
  });

  it("still uses /me/events when no date bounds are given", async () => {
    const spy = mockHttpsSeq([TOKEN, 200], [EVENTS, 200]);
    const { api } = await loadPlugin();
    await api.tools["outlook_query_events"].execute("id", { text: "Standup" });
    expect(String(spy.mock.calls.at(-1)?.[0])).toContain("/me/events");
  });

  it("queries a named calendar by resolving its id", async () => {
    const spy = mockHttpsSeq([TOKEN, 200], [CALENDARS, 200], [EVENTS, 200]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_query_events"].execute("id", {
      calendar: "family", after: "2027-09-01", before: "2027-09-30",
    })) as { calendar: string; count: number };

    const url = decodeURIComponent(String(spy.mock.calls.at(-1)?.[0]));
    expect(url).toContain("/me/calendars/fam-1/calendarView");
    expect(url).toContain("startDateTime=2027-09-01T00:00:00");
    expect(url).toContain("endDateTime=2027-10-01T00:00:00");
    expect(data.calendar).toBe("family");
    expect(data.count).toBe(1);
  });

  it("reports available calendars when the requested one is missing", async () => {
    const unknownCals = JSON.stringify({ value: [{ name: "Work", id: "w-1" }] });
    mockHttpsSeq([TOKEN, 200], [unknownCals, 200]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_query_events"].execute("id", {
      calendar: "family", after: "2027-09-01",
    })) as { error: string };
    expect(data.error).toContain("not found");
    expect(data.error).toContain("work");
  });

  it("filters a date-range result by attendee", async () => {
    mockHttpsSeq([TOKEN, 200], [EVENTS, 200]);
    const { api } = await loadPlugin();
    const data = resultText(await api.tools["outlook_query_events"].execute("id", {
      after: "2026-05-01", before: "2026-05-03", attendee: "nobody@test.com",
    })) as { count: number };
    expect(data.count).toBe(0);
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
