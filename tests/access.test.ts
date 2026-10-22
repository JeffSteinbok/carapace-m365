import { describe, it, expect, vi, afterEach, beforeEach } from "vitest";
import https from "node:https";
import { EventEmitter } from "node:events";
import { rmSync } from "node:fs";
import { resolve } from "node:path";

// Calendar access control: tools are registered as factories and receive the
// trusted OpenClaw tool context; config.access.calendars gates each calendar.

const scratch = resolve("tests", ".test-state-access");
let seq = 0;

beforeEach(() => {
  delete process.env.M365_TOKEN_BROKER_URL;
  delete process.env.M365_TOKEN_BROKER_SECRET;
  delete process.env.M365_FEATURES;
  process.env.M365_CLIENT_ID = "cid";
  process.env.M365_CLIENT_SECRET = "csec";
  process.env.M365_REFRESH_TOKEN = `acl-rtoken-${seq++}`;
  process.env.M365_TENANT = "consumers";
  process.env.M365_DIRECT_TOKEN_STATE_PATH = resolve(scratch, `direct-${seq}.json`);
});

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

type Concrete = { name: string; execute: (id: string, params: Record<string, unknown>) => Promise<unknown> };
type Factory = ((ctx?: Record<string, unknown>) => Concrete) & Concrete;

async function loadPlugin(pluginConfig: Record<string, unknown>) {
  const { createEntry } = await import("../src/index.js");
  const tools: Record<string, Factory> = {};
  const opts: Record<string, unknown> = {};
  createEntry().register({
    pluginConfig,
    registerTool(t: unknown, o?: { name?: string }) {
      const f = t as Factory;
      tools[f.name] = f;
      opts[f.name] = o;
    },
  });
  return { tools, opts };
}

const text = (r: unknown) => JSON.parse((r as { content: Array<{ text: string }> }).content[0].text);

const TOKEN = JSON.stringify({ access_token: "test-token" });
const CALENDARS = JSON.stringify({ value: [{ name: "Calendar", id: "cal-1" }, { name: "Your Family", id: "fam-1" }, { name: "Birthdays", id: "bday-1" }] });
const EVENTS = JSON.stringify({ value: [{ id: "evt1", subject: "Standup", start: { dateTime: "2026-05-03T17:00:00Z" }, end: { dateTime: "2026-05-03T17:30:00Z" } }] });

const accessGroups = {
  jeff: { type: "message.senders", members: { imessage: ["+14255550100"] } },
};
const ACCESS = {
  calendars: {
    family: { read: ["*"], write: ["*"] },
    personal: { read: [{ agent: "main" }, { sender: "jeff" }], write: [{ agent: "main" }] },
  },
};

const main = { agentId: "main" };
const family = { agentId: "family", config: { accessGroups } };
const familyJeff = { agentId: "family", messageChannel: "imessage", requesterSenderId: "+14255550100", config: { accessGroups } };

describe("registration", () => {
  it("forwards the tool name through the feature filter", async () => {
    const { tools, opts } = await loadPlugin({});
    expect(typeof tools.outlook_calendar_fetch).toBe("function");
    expect(opts.outlook_calendar_fetch).toEqual({ name: "outlook_calendar_fetch" });
  });
});

describe("without an access block", () => {
  it("allows every calendar to every caller", async () => {
    mockHttpsSeq([TOKEN, 200], [CALENDARS, 200], [EVENTS, 200], [EVENTS, 200]);
    const { tools } = await loadPlugin({});
    const data = text(await tools.outlook_calendar_fetch(family).execute("id", { calendar: "all" }));
    expect(Object.keys(data).sort()).toEqual(["family", "personal"]);
  });

  it("does not look up the event calendar before update/delete", async () => {
    const spy = mockHttpsSeq([TOKEN, 200], ["", 204]);
    const { tools } = await loadPlugin({});
    const data = text(await tools.outlook_delete_event(family).execute("id", { event_id: "evt1" }));
    expect(data.success).toBe(true);
    expect(spy.mock.calls.map((c) => String(c[0]))).not.toContainEqual(expect.stringContaining("/calendar?"));
  });
});

describe("outlook_calendar_fetch", () => {
  it("denies a calendar the caller may not read, without calling Graph", async () => {
    const spy = mockHttpsSeq();
    const { tools } = await loadPlugin({ access: ACCESS });
    const data = text(await tools.outlook_calendar_fetch(family).execute("id", { calendar: "personal" }));
    expect(data).toEqual({ error: 'Access denied: agent "family" (no sender) may not read calendar "personal"' });
    expect(spy).not.toHaveBeenCalled();
  });

  it("'all' returns only the readable calendars", async () => {
    const spy = mockHttpsSeq([TOKEN, 200], [CALENDARS, 200], [EVENTS, 200]);
    const { tools } = await loadPlugin({ access: ACCESS });
    const data = text(await tools.outlook_calendar_fetch(family).execute("id", { calendar: "all" }));
    expect(data.family.count).toBe(1);
    expect(data.personal).toBeUndefined();
    expect(data.omitted_calendars).toEqual(["personal"]);
    expect(String(spy.mock.calls[2][0])).toContain("/me/calendars/fam-1/calendarView");
  });

  it("'all' errors only when nothing is readable", async () => {
    const { tools } = await loadPlugin({ access: { calendars: {} } });
    const data = text(await tools.outlook_calendar_fetch(main).execute("id", {}));
    expect(data.error).toContain('may not read calendar "personal"');
    expect(data.error).toContain('may not read calendar "family"');
  });

  it("allows by sender group", async () => {
    mockHttpsSeq([TOKEN, 200], [CALENDARS, 200], [EVENTS, 200]);
    const { tools } = await loadPlugin({ access: ACCESS });
    const data = text(await tools.outlook_calendar_fetch(familyJeff).execute("id", { calendar: "personal" }));
    expect(data.personal.count).toBe(1);
  });

  it("the standalone/static execute (no context) is treated as a senderless, agentless caller", async () => {
    const { tools } = await loadPlugin({ access: ACCESS });
    const data = text(await tools.outlook_calendar_fetch.execute("id", { calendar: "personal" }));
    expect(data.error).toBe('Access denied: unknown agent (no sender) may not read calendar "personal"');
  });
});

describe("outlook_query_events", () => {
  it("treats a missing calendar as the personal (default) calendar", async () => {
    const { tools } = await loadPlugin({ access: ACCESS });
    const data = text(await tools.outlook_query_events(family).execute("id", { text: "x" }));
    expect(data.error).toContain('read calendar "personal"');
  });

  it("checks uid lookups too, and scopes them to the named calendar", async () => {
    const { tools } = await loadPlugin({ access: ACCESS });
    expect(text(await tools.outlook_query_events(family).execute("id", { uid: "u1" })).error).toContain("Access denied");

    const spy = mockHttpsSeq([TOKEN, 200], [CALENDARS, 200], [EVENTS, 200]);
    const data = text(await tools.outlook_query_events(family).execute("id", { uid: "u1", calendar: "family" }));
    expect(data.events).toHaveLength(1);
    expect(String(spy.mock.calls[2][0])).toContain("/me/calendars/fam-1/events?$filter=iCalUId");
  });
});

describe("writes", () => {
  it("outlook_create_event defaults to personal and requires write", async () => {
    const { tools } = await loadPlugin({ access: ACCESS });
    const data = text(await tools.outlook_create_event(familyJeff).execute("id", { subject: "x", start: "2026-08-01T10:00" }));
    expect(data.error).toBe('Access denied: agent "family" / sender +14255550100 may not write calendar "personal"');
  });

  it("outlook_create_event on family is allowed for anyone", async () => {
    const created = { id: "evt-new", subject: "x", start: { dateTime: "2026-08-01T17:00:00Z" } };
    mockHttpsSeq([TOKEN, 200], [CALENDARS, 200], [JSON.stringify(created), 201]);
    const { tools } = await loadPlugin({ access: ACCESS });
    const data = text(await tools.outlook_create_event(family).execute("id", { subject: "x", start: "2026-08-01T10:00", calendar: "family" }));
    expect(data.success).toBe(true);
  });

  it("outlook_meeting requires write on personal", async () => {
    const { tools } = await loadPlugin({ access: ACCESS });
    const data = text(await tools.outlook_meeting(family).execute("id", { to: "a@b.c", subject: "x", start: "2026-08-01T10:00" }));
    expect(data.error).toContain('write calendar "personal"');
  });

  it("update/delete look up the event's calendar and deny personal events for other agents", async () => {
    const spy = mockHttpsSeq([TOKEN, 200], [JSON.stringify({ id: "cal-1", name: "Calendar" }), 200], [CALENDARS, 200]);
    const { tools } = await loadPlugin({ access: ACCESS });
    const data = text(await tools.outlook_delete_event(family).execute("id", { event_id: "evt1" }));
    expect(data.error).toBe('Access denied: agent "family" (no sender) may not write calendar "personal"');
    expect(String(spy.mock.calls[1][0])).toContain("/me/events/evt1/calendar?$select=id,name");
    // No DELETE was issued.
    expect(spy).toHaveBeenCalledTimes(3);
  });

  it("update/delete allow family events", async () => {
    mockHttpsSeq([TOKEN, 200], [JSON.stringify({ id: "fam-1", name: "Your Family" }), 200], [CALENDARS, 200], ["", 204]);
    const { tools } = await loadPlugin({ access: ACCESS });
    const data = text(await tools.outlook_delete_event(family).execute("id", { event_id: "evt2" }));
    expect(data.success).toBe(true);
  });

  it("update/delete deny events on unlisted calendars", async () => {
    mockHttpsSeq([TOKEN, 200], [JSON.stringify({ id: "bday-1", name: "Birthdays" }), 200], [CALENDARS, 200]);
    const { tools } = await loadPlugin({ access: ACCESS });
    const data = text(await tools.outlook_update_event(main).execute("id", { event_id: "evt3", subject: "y" }));
    expect(data.error).toBe('Access denied: agent "main" (no sender) may not write calendar "other:Birthdays"');
  });

  it("update/delete deny when the event cannot be resolved", async () => {
    mockHttpsSeq([TOKEN, 200], [JSON.stringify({ error: { message: "The specified object was not found in the store." } }), 404]);
    const { tools } = await loadPlugin({ access: ACCESS });
    const data = text(await tools.outlook_update_event(main).execute("id", { event_id: "nope", subject: "y" }));
    expect(data.error).toContain("could not determine the calendar");
  });
});
