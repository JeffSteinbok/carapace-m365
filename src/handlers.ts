/**
 * Outlook Calendar — core handlers.
 *
 * Pure logic with no knowledge of how it's invoked (plugin vs CLI).
 * Config is passed in; handlers never read env vars or plugin APIs directly.
 */

import https from "node:https";
import { createHash } from "node:crypto";
import {
  GraphTokenManager,
  isM365FeatureEnabled,
  type GraphTokenManagerOptions,
} from "@carapace/m365-graph-auth";
import { DirectTokenStateStore } from "./direct-token-state.js";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface OutlookCalendarConfig {
  clientId: string;
  clientSecret: string;
  refreshToken: string;
  tenant: string;
  tokenBrokerUrl: string;
  tokenBrokerSecret: string;
  directTokenStatePath: string;
  features: string[];
  personalCalendarNames: string[];
  familyCalendarNames: string[];
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const SCOPES = {
  calendarRead: ["Calendars.Read"],
  calendarWrite: ["Calendars.ReadWrite"],
  mailRead: ["Mail.Read"],
  mailSend: ["Mail.Send"],
  mailWrite: ["Mail.ReadWrite"],
  tasksRead: ["Tasks.Read"],
  tasksWrite: ["Tasks.ReadWrite"],
} as const;
const CALENDAR_DEFAULTS: Record<string, string[]> = {
  personal: ["calendar", "personal"],
  family: ["family v2", "your family", "family"],
};
const TASK_LIST_DEFAULTS = ["tasks", "to do"];

function taskReadScopes(config: OutlookCalendarConfig): readonly string[] {
  return isM365FeatureEnabled(config.features, "tasks-write")
    ? SCOPES.tasksWrite
    : SCOPES.tasksRead;
}

// ---------------------------------------------------------------------------
// HTTP helpers
// ---------------------------------------------------------------------------

function httpGet(url: string, token: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const req = https.request(url, { method: "GET", headers: { Authorization: `Bearer ${token}`, Accept: "application/json" }, timeout: 30_000 }, res => {
      let data = ""; res.on("data", (c: Buffer) => data += c); res.on("end", () => resolve(data));
    });
    req.on("error", reject); req.on("timeout", () => { req.destroy(); reject(new Error("timeout")); }); req.end();
  });
}

function httpRequest(
  method: "PATCH" | "DELETE",
  url: string,
  token: string,
  body?: string,
): Promise<{ status: number; data: string }> {
  return new Promise((resolve, reject) => {
    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    };
    if (body) {
      headers["Content-Type"] = "application/json";
      headers["Content-Length"] = String(Buffer.byteLength(body));
    }
    const req = https.request(url, { method, headers, timeout: 30_000 }, res => {
      let data = ""; res.on("data", (c: Buffer) => data += c); res.on("end", () => resolve({ status: res.statusCode ?? 0, data }));
    });
    req.on("error", reject); req.on("timeout", () => { req.destroy(); reject(new Error("timeout")); });
    if (body) req.write(body);
    req.end();
  });
}

function httpPostJson(
  url: string,
  token: string,
  body: string,
): Promise<{ status: number; data: string }> {
  return new Promise((resolve, reject) => {
    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      "Content-Type": "application/json",
      "Content-Length": String(Buffer.byteLength(body)),
    };
    const req = https.request(url, { method: "POST", headers, timeout: 30_000 }, res => {
      let data = ""; res.on("data", (c: Buffer) => data += c); res.on("end", () => resolve({ status: res.statusCode ?? 0, data }));
    });
    req.on("error", reject); req.on("timeout", () => { req.destroy(); reject(new Error("timeout")); });
    req.write(body);
    req.end();
  });
}

// ---------------------------------------------------------------------------
// Token / Graph
// ---------------------------------------------------------------------------

const tokenManagers = new Map<string, GraphTokenManager>();

function getAuthConfigError(config: OutlookCalendarConfig): string | undefined {
  if (config.tokenBrokerUrl) {
    return config.tokenBrokerSecret
      ? undefined
      : "M365_TOKEN_BROKER_SECRET must be set when a token broker is configured";
  }
  if (!config.clientId) {
    return "Set M365_CLIENT_ID and M365_REFRESH_TOKEN, or configure the token broker";
  }
  return undefined;
}

function tokenManagerKey(config: OutlookCalendarConfig): string {
  const mode = config.tokenBrokerUrl ? "broker" : "direct";
  const identity = config.tokenBrokerUrl
    ? {
        tokenBrokerUrl: config.tokenBrokerUrl,
        tokenBrokerSecret: config.tokenBrokerSecret,
      }
    : {
        clientId: config.clientId,
        clientSecret: config.clientSecret,
        tenant: config.tenant,
        directTokenStatePath: config.directTokenStatePath,
      };
  const digest = createHash("sha256").update(JSON.stringify(identity)).digest("hex");
  return `${mode}:${digest}`;
}

function getTokenManager(config: OutlookCalendarConfig): GraphTokenManager {
  const key = tokenManagerKey(config);
  let manager = tokenManagers.get(key);
  if (!manager) {
    let options: GraphTokenManagerOptions;
    if (config.tokenBrokerUrl) {
      options = {
        tokenBrokerUrl: config.tokenBrokerUrl,
        tokenBrokerSecret: config.tokenBrokerSecret,
      };
    } else {
      const store = new DirectTokenStateStore(config.directTokenStatePath);
      const persistedToken = store.load().refreshToken;
      const refreshToken = persistedToken || config.refreshToken;
      if (!refreshToken) {
        throw new Error(
          "Set M365_REFRESH_TOKEN or configure the token broker",
        );
      }
      if (!persistedToken) store.save({ refreshToken });
      options = {
        clientId: config.clientId,
        clientSecret: config.clientSecret || undefined,
        refreshToken,
        tenant: config.tenant,
        onRefreshToken: async (rotatedToken) => {
          store.save({ refreshToken: rotatedToken });
        },
      };
    }
    manager = new GraphTokenManager(options);
    tokenManagers.set(key, manager);
  }
  return manager;
}

export async function getGraphAccessToken(
  config: OutlookCalendarConfig,
  scopes: readonly string[],
): Promise<string> {
  const configError = getAuthConfigError(config);
  if (configError) throw new Error(configError);
  return getTokenManager(config).getAccessToken(scopes);
}

function httpPatch(url: string, body: string, token: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const headers: Record<string, string | number> = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      Accept: "application/json",
      "Content-Length": Buffer.byteLength(body),
    };
    const req = https.request(url, { method: "PATCH", headers: headers as Record<string, string>, timeout: 30_000 }, res => {
      let data = ""; res.on("data", (c: Buffer) => data += c); res.on("end", () => resolve(data));
    });
    req.on("error", reject); req.on("timeout", () => { req.destroy(); reject(new Error("timeout")); });
    req.write(body); req.end();
  });
}

async function graphGet(token: string, path: string): Promise<unknown> {
  const res = await httpGet(`${GRAPH_BASE}${path}`, token);
  return JSON.parse(res);
}

async function graphPost(token: string, path: string, body: unknown): Promise<unknown> {
  const bodyStr = JSON.stringify(body);
  const res = await httpPostJson(`${GRAPH_BASE}${path}`, token, bodyStr);
  if (!res.data || res.data.trim() === "") return { success: true };
  if (res.status >= 400) return { error: res.data };
  return JSON.parse(res.data);
}

async function graphPatch(token: string, path: string, body: unknown): Promise<unknown> {
  const bodyStr = JSON.stringify(body);
  const res = await httpPatch(`${GRAPH_BASE}${path}`, bodyStr, token);
  if (!res || res.trim() === "") return { success: true };
  return JSON.parse(res);
}

function esc(s: string): string { return s.replace(/'/g, "''"); }

function formatMessage(m: Record<string, unknown>, includeBody = false): Record<string, unknown> {
  const from = (m.from as Record<string, Record<string, string>>)?.emailAddress ?? {};
  const result: Record<string, unknown> = {
    id: m.id,
    subject: m.subject ?? "(no subject)",
    from: `${from.name ?? ""}${from.address ? ` <${from.address}>` : ""}`.trim(),
    received: String(m.receivedDateTime ?? "").slice(0, 10),
    is_read: m.isRead,
    has_attachments: m.hasAttachments,
  };
  if (includeBody) result.body_preview = (m.bodyPreview as string ?? "").slice(0, 500);
  return result;
}

function stripHtml(input: string): string {
  return input.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
}

function formatTask(task: Record<string, unknown>): Record<string, unknown> {
  const due = task.dueDateTime as Record<string, string> | undefined;
  const reminder = task.reminderDateTime as Record<string, string> | undefined;
  const body = task.body as Record<string, string> | undefined;
  const bodyContent = body?.content ?? "";
  return {
    id: String(task.id ?? ""),
    title: String(task.title ?? "No title"),
    status: String(task.status ?? "notStarted"),
    importance: String(task.importance ?? "normal"),
    created: String(task.createdDateTime ?? ""),
    last_modified: String(task.lastModifiedDateTime ?? ""),
    due: due?.dateTime ? { dateTime: due.dateTime, timeZone: due.timeZone ?? "" } : null,
    reminder: reminder?.dateTime ? { dateTime: reminder.dateTime, timeZone: reminder.timeZone ?? "" } : null,
    notes: bodyContent ? (body?.contentType === "html" ? stripHtml(bodyContent) : bodyContent) : "",
  };
}

async function listTaskLists(token: string): Promise<Array<Record<string, unknown>>> {
  const data = await graphGet(token, "/me/todo/lists?$top=100") as { value: Array<Record<string, unknown>> };
  return data.value ?? [];
}

async function resolveTaskList(
  token: string,
  selector?: string,
): Promise<{ id: string; name: string; lists: Array<Record<string, unknown>> }> {
  const lists = await listTaskLists(token);
  if (!lists.length) {
    throw new Error("No Microsoft To Do lists found");
  }
  const sel = selector?.trim().toLowerCase();
  const byId = selector ? lists.find(l => String(l.id ?? "") === selector) : undefined;
  const byName = sel ? lists.find(l => String(l.displayName ?? "").trim().toLowerCase() === sel) : undefined;
  const byDefault = lists.find(l => TASK_LIST_DEFAULTS.includes(String(l.displayName ?? "").trim().toLowerCase()));
  const chosen = byId ?? byName ?? byDefault ?? lists[0];
  if (!chosen?.id) {
    throw new Error("Unable to resolve Microsoft To Do list");
  }
  return {
    id: String(chosen.id),
    name: String(chosen.displayName ?? chosen.id),
    lists,
  };
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function utcToLocal(s: string): string {
  try {
    const dt = new Date(s.slice(0, 19) + "Z");
    return dt.toLocaleString("en-US", { year: "numeric", month: "2-digit", day: "2-digit", hour: "numeric", minute: "2-digit", hour12: true });
  } catch { return s.slice(0, 16); }
}

function formatEvent(e: Record<string, unknown>): Record<string, unknown> {
  const start = (e.start as Record<string, string>); const end = (e.end as Record<string, string>);
  const tz = start?.timeZone ?? "UTC";
  const fmtTime = (s: string) => tz === "UTC" ? utcToLocal(s) : s.slice(0, 16);
  const attendees = ((e.attendees ?? []) as Array<Record<string, unknown>>).map(a => {
    const ea = (a.emailAddress ?? {}) as Record<string, string>;
    return { name: ea.name ?? "", email: ea.address ?? "", status: ((a.status as Record<string, string>)?.response ?? "none"), type: String(a.type ?? "required") };
  });
  const result: Record<string, unknown> = {
    id: String(e.id ?? ""),
    subject: String(e.subject ?? "No subject"), start: fmtTime(start?.dateTime ?? ""), end: fmtTime(end?.dateTime ?? ""),
    location: ((e.location as Record<string, string>)?.displayName || "No location"),
    organizer: ((e.organizer as Record<string, Record<string, string>>)?.emailAddress?.name || ((e.organizer as Record<string, Record<string, string>>)?.emailAddress?.address ?? "")),
    my_status: ((e.responseStatus as Record<string, string>)?.response ?? "none"), show_as: String(e.showAs ?? "busy"),
  };
  if (attendees.length) result.attendees = attendees;
  const bodyObj = e.body as Record<string, string> | undefined;
  const bodyContent = bodyObj?.content;
  const bodyType = bodyObj?.contentType ?? "html";
  if (bodyContent && bodyContent.trim()) {
    let plainText: string;
    if (bodyType === "text") {
      plainText = bodyContent.replace(/\s+/g, " ").trim();
    } else {
      // HTML body: strip tags first, then collapse whitespace only.
      // Entity sequences (&amp; &lt; etc.) are left as-is — they are HTML
      // artefacts that don't need unescaping for plain-text display.
      plainText = bodyContent.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
    }
    if (plainText.length > 0) result.body = plainText;
  }
  return result;
}

function calendarSearchNames(config: OutlookCalendarConfig, key: string): string[] {
  const extraNames = key === "personal" ? config.personalCalendarNames : config.familyCalendarNames;
  return [...extraNames, ...CALENDAR_DEFAULTS[key]];
}


// ---------------------------------------------------------------------------
// Unified config (superset of mail + calendar configs)
// ---------------------------------------------------------------------------

export type OutlookConfig = OutlookCalendarConfig;
export type OutlookMailConfig = OutlookConfig;
/** Resolve ISO datetime + optional timezone into a Graph dateTimeTimeZone object. */
function toGraphDateTime(isoStr: string, timezone: string): { dateTime: string; timeZone: string } {
  // Strip any trailing Z or offset — Graph wants wall-clock in the given tz
  const dt = isoStr.replace(/Z$/, "").replace(/[+-]\d{2}:\d{2}$/, "");
  return { dateTime: dt.length === 16 ? `${dt}:00` : dt, timeZone: timezone };
}

/** Parse duration string like "1h", "30m", "1.5h", "2h30m" into minutes. */
function parseDurationMinutes(s: string): number {
  const hours = s.match(/(\d+(?:\.\d+)?)\s*h/i);
  const mins = s.match(/(\d+(?:\.\d+)?)\s*m(?!o)/i); // avoid matching 'month'
  let total = 0;
  if (hours) total += parseFloat(hours[1]) * 60;
  if (mins) total += parseFloat(mins[1]);
  if (!hours && !mins) {
    const num = parseFloat(s);
    if (!isNaN(num)) total = num; // bare number treated as minutes
  }
  return Math.round(total);
}

/** Add minutes to an ISO datetime string (wall-clock, no tz offset). */
function addMinutes(isoStr: string, minutes: number): string {
  const clean = isoStr.replace(/Z$/, "").replace(/[+-]\d{2}:\d{2}$/, "");
  const d = new Date(clean + "Z");
  d.setUTCMinutes(d.getUTCMinutes() + minutes);
  return d.toISOString().slice(0, 19);
}

// ---------------------------------------------------------------------------
// Handlers
// ---------------------------------------------------------------------------

export async function fetchCalendar(
  config: OutlookCalendarConfig,
  params: { calendar?: string; days?: number },
): Promise<unknown> {
  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  const calendar = params.calendar ?? "all";
  const days = params.days ?? 7;
  const token = await getGraphAccessToken(config, SCOPES.calendarRead);
  const calData = await graphGet(token, "/me/calendars?$select=id,name&$top=50") as { value: Array<{ name: string; id: string }> };
  const calMap: Record<string, string> = {};
  for (const c of calData.value ?? []) calMap[c.name.toLowerCase()] = c.id;
  const start = new Date().toISOString().slice(0, 10);
  const end = new Date(Date.now() + days * 86400000).toISOString().slice(0, 10);
  const keys = calendar === "all" ? ["personal", "family"] : [calendar];
  const results: Record<string, unknown> = {};
  for (const key of keys) {
    const searchNames = calendarSearchNames(config, key);
    const calId = searchNames.map(n => calMap[n]).find(Boolean);
    if (!calId) { results[key] = { label: key, error: `Calendar not found. Available: ${Object.keys(calMap).join(", ")}`, events: [] }; continue; }
    const qp = new URLSearchParams({ "$select": "id,subject,start,end,location,organizer,attendees,responseStatus,showAs,body", "$orderby": "start/dateTime", "$top": "100", "startDateTime": `${start}T00:00:00`, "endDateTime": `${end}T00:00:00` }).toString();
    const evData = await graphGet(token, `/me/calendars/${calId}/calendarView?${qp}`) as { value: Array<Record<string, unknown>> };
    const events = (evData.value ?? []).map(formatEvent);
    results[key] = { label: key === "personal" ? "Personal" : "Family", count: events.length, start_date: start, end_date: end, events };
  }
  return results;
}

export interface CreateEventParams {
  subject: string;
  start: string;
  duration?: string;
  end?: string;
  timezone?: string;
  location?: string;
  description?: string;
  attendees?: string[];
  calendar?: string;
}

export async function createEvent(
  config: OutlookCalendarConfig,
  params: CreateEventParams,
): Promise<unknown> {
  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  if (!params.subject) return { error: "subject is required" };
  if (!params.start) return { error: "start is required" };

  const timezone = params.timezone ?? "America/Los_Angeles";
  const token = await getGraphAccessToken(config, SCOPES.calendarWrite);

  // Resolve end time
  let endIso: string;
  if (params.end) {
    endIso = params.end;
  } else {
    const durationMins = parseDurationMinutes(params.duration ?? "1h");
    endIso = addMinutes(params.start, durationMins);
  }

  // Resolve calendar ID
  let calId: string | undefined;
  const calKey = params.calendar ?? "personal";
  const calData = await graphGet(token, "/me/calendars?$select=id,name&$top=50") as { value: Array<{ name: string; id: string }> };
  const calMap: Record<string, string> = {};
  for (const c of calData.value ?? []) calMap[c.name.toLowerCase()] = c.id;
  const searchNames = calendarSearchNames(config, calKey);
  calId = searchNames.map(n => calMap[n]).find(Boolean);
  if (!calId) return { error: `Calendar '${calKey}' not found. Available: ${Object.keys(calMap).join(", ")}` };

  const body: Record<string, unknown> = {
    subject: params.subject,
    start: toGraphDateTime(params.start, timezone),
    end: toGraphDateTime(endIso, timezone),
  };
  if (params.location) body.location = { displayName: params.location };
  if (params.description) body.body = { contentType: "text", content: params.description };
  if (params.attendees?.length) {
    body.attendees = params.attendees.map(email => ({
      emailAddress: { address: email },
      type: "required",
    }));
  }

  const res = await httpPostJson(`${GRAPH_BASE}/me/calendars/${calId}/events`, token, JSON.stringify(body));
  if (res.status < 200 || res.status >= 300) {
    const err = JSON.parse(res.data ?? "{}");
    return { error: `Graph API error ${res.status}: ${err?.error?.message ?? res.data}` };
  }
  const created = JSON.parse(res.data) as Record<string, unknown>;
  return {
    success: true,
    event_id: String(created.id ?? ""),
    subject: String(created.subject ?? ""),
    start: (created.start as Record<string, string>)?.dateTime ?? "",
    end: (created.end as Record<string, string>)?.dateTime ?? "",
    timezone,
    calendar: calKey,
    web_link: String(created.webLink ?? ""),
  };
}

export interface UpdateEventParams {
  event_id: string;
  subject?: string;
  start?: string;
  end?: string;
  duration?: string;
  timezone?: string;
  location?: string;
  description?: string;
  add_attendees?: string[];
  remove_attendees?: string[];
  status?: "confirmed" | "tentative" | "cancelled";
}

export async function updateEvent(
  config: OutlookCalendarConfig,
  params: UpdateEventParams,
): Promise<unknown> {
  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  if (!params.event_id) return { error: "event_id is required" };

  const token = await getGraphAccessToken(config, SCOPES.calendarWrite);
  const patch: Record<string, unknown> = {};

  const timezone = params.timezone ?? "America/Los_Angeles";

  if (params.subject) patch.subject = params.subject;
  if (params.location !== undefined) patch.location = { displayName: params.location };
  if (params.description !== undefined) patch.body = { contentType: "text", content: params.description };

  if (params.status) {
    const showAs: Record<string, string> = { confirmed: "busy", tentative: "tentative", cancelled: "free" };
    patch.showAs = showAs[params.status] ?? "busy";
    if (params.status === "cancelled") patch.isCancelled = true;
  }

  if (params.start) {
    patch.start = toGraphDateTime(params.start, timezone);
    if (params.end) {
      patch.end = toGraphDateTime(params.end, timezone);
    } else if (params.duration) {
      patch.end = toGraphDateTime(addMinutes(params.start, parseDurationMinutes(params.duration)), timezone);
    }
  } else if (params.end) {
    patch.end = toGraphDateTime(params.end, timezone);
  }

  // Attendee merge: fetch current, diff, patch full list
  if (params.add_attendees?.length || params.remove_attendees?.length) {
    const evData = await graphGet(token, `/me/events/${params.event_id}?$select=attendees`) as Record<string, unknown>;
    const current = ((evData.attendees ?? []) as Array<Record<string, unknown>>).map(a => {
      const ea = (a.emailAddress as Record<string, string> | undefined) ?? {};
      return { email: (ea.address ?? "").toLowerCase(), type: String(a.type ?? "required") };
    });
    const removeSet = new Set((params.remove_attendees ?? []).map(e => e.toLowerCase()));
    const kept = current.filter(a => !removeSet.has(a.email));
    const keptEmails = new Set(kept.map(a => a.email));
    const added = (params.add_attendees ?? [])
      .filter(e => !keptEmails.has(e.toLowerCase()))
      .map(e => ({ email: e.toLowerCase(), type: "required" }));
    patch.attendees = [...kept, ...added].map(a => ({
      emailAddress: { address: a.email },
      type: a.type,
    }));
  }

  if (Object.keys(patch).length === 0) {
    return { error: "No fields to update were provided" };
  }

  const res = await httpRequest("PATCH", `${GRAPH_BASE}/me/events/${params.event_id}`, token, JSON.stringify(patch));
  if (res.status < 200 || res.status >= 300) {
    const err = JSON.parse(res.data ?? "{}");
    return { error: `Graph API error ${res.status}: ${err?.error?.message ?? res.data}` };
  }
  const updated = JSON.parse(res.data) as Record<string, unknown>;
  return {
    success: true,
    event_id: String(updated.id ?? params.event_id),
    subject: String(updated.subject ?? ""),
    start: (updated.start as Record<string, string>)?.dateTime ?? "",
    end: (updated.end as Record<string, string>)?.dateTime ?? "",
  };
}

export interface DeleteEventParams {
  event_id: string;
}

export async function deleteEvent(
  config: OutlookCalendarConfig,
  params: DeleteEventParams,
): Promise<unknown> {
  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  if (!params.event_id) return { error: "event_id is required" };

  const token = await getGraphAccessToken(config, SCOPES.calendarWrite);
  const res = await httpRequest("DELETE", `${GRAPH_BASE}/me/events/${params.event_id}`, token);
  if (res.status === 204) return { success: true, event_id: params.event_id };
  if (res.status < 200 || res.status >= 300) {
    const err = (() => { try { return JSON.parse(res.data ?? "{}"); } catch { return {}; } })();
    return { error: `Graph API error ${res.status}: ${err?.error?.message ?? res.data}` };
  }
  return { success: true, event_id: params.event_id };
}

// ---------------------------------------------------------------------------
// createMeeting — send invite to attendees
// ---------------------------------------------------------------------------

export interface CreateMeetingParams {
  to: string | string[];
  cc?: string[];
  subject: string;
  start: string;
  duration?: string;
  end?: string;
  location?: string;
  description?: string;
  timezone?: string;
  signature?: string;
}

export async function createMeeting(
  config: OutlookCalendarConfig,
  params: CreateMeetingParams,
): Promise<unknown> {
  if (!params.subject?.trim()) return { error: "subject is required" };
  if (!params.start?.trim()) return { error: "start is required" };
  const toList = Array.isArray(params.to) ? params.to : [params.to];
  if (!toList.length || !toList.some((address) => address?.trim())) {
    return { error: "to is required" };
  }
  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  const token = await getGraphAccessToken(config, SCOPES.calendarWrite);
  const tz = params.timezone ?? "America/Los_Angeles";
  const start = toGraphDateTime(params.start, tz);
  const endDt = params.end
    ? toGraphDateTime(params.end, tz)
    : { dateTime: addMinutes(params.start, parseDurationMinutes(params.duration ?? "1h")), timeZone: tz };

  const ccList = params.cc ?? [];
  const attendees = [
    ...toList.map(e => ({ emailAddress: { address: e }, type: "required" })),
    ...ccList.map(e => ({ emailAddress: { address: e }, type: "optional" })),
  ];

  const body: Record<string, unknown> = {
    subject: params.subject,
    start,
    end: endDt,
    attendees,
    ...(params.location ? { location: { displayName: params.location } } : {}),
    ...(params.description ? { body: { contentType: "Text", content: params.description } } : {}),
  };

  const res = await httpPostJson(`${GRAPH_BASE}/me/events`, token, JSON.stringify(body));
  if (res.status < 200 || res.status >= 300) {
    const err = JSON.parse(res.data ?? "{}");
    return { error: `Graph API error ${res.status}: ${err?.error?.message ?? res.data}` };
  }
  const created = JSON.parse(res.data) as Record<string, unknown>;
  if (created.error) return { error: JSON.stringify(created.error) };
  return {
    ok: true,
    id: created.id,
    iCalUId: created.iCalUId,
    subject: created.subject,
    start: (created.start as Record<string, string>)?.dateTime,
    end: (created.end as Record<string, string>)?.dateTime,
    webLink: created.webLink,
    message: `✓ Meeting created: ${params.subject}`,
  };
}

// ---------------------------------------------------------------------------
// queryEvents — filter events by date range / text / attendee / UID
// ---------------------------------------------------------------------------

export interface QueryEventsParams {
  after?: string;
  before?: string;
  text?: string;
  attendee?: string;
  uid?: string;
}

export async function queryEvents(
  config: OutlookCalendarConfig,
  params: QueryEventsParams,
): Promise<unknown> {
  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  const token = await getGraphAccessToken(config, SCOPES.calendarRead);

  const esc = (s: string) => s.replace(/'/g, "''");

  if (params.uid) {
    const res = await graphGet(token, `/me/events?$filter=iCalUId eq '${esc(params.uid)}'`) as { value: Array<Record<string, unknown>> };
    return { events: (res.value ?? []).map(formatEvent) };
  }

  const filters: string[] = [];
  if (params.after) filters.push(`start/dateTime ge '${params.after}T00:00:00'`);
  if (params.before) filters.push(`end/dateTime le '${params.before}T23:59:59'`);
  if (params.text) filters.push(`contains(subject,'${esc(params.text)}')`);

  const qs = filters.length
    ? `?$filter=${encodeURIComponent(filters.join(" and "))}&$top=50&$orderby=start/dateTime`
    : `?$top=20&$orderby=start/dateTime`;

  const res = await graphGet(token, `/me/events${qs}`) as { value: Array<Record<string, unknown>> };
  let events = (res.value ?? []).map(formatEvent);

  if (params.attendee) {
    const att = params.attendee.toLowerCase();
    events = events.filter(e =>
      (e.attendees as Array<Record<string, string>>)?.some(
        (a: Record<string, string>) => a.email?.toLowerCase() === att,
      ),
    );
  }

  return { count: events.length, events };
}
export async function getInbox(
  config: OutlookMailConfig,
  params: { limit?: number; unread?: boolean; folder?: string },
): Promise<unknown> {
  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  const token = await getGraphAccessToken(config, SCOPES.mailRead);
  const limit = params.limit ?? 10;
  const folder = params.folder ?? "inbox";
  let path = `/me/mailFolders/${encodeURIComponent(folder)}/messages?$top=${limit}&$select=subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview&$orderby=receivedDateTime%20desc`;
  if (params.unread) path += "&$filter=isRead%20eq%20false";
  const data = await graphGet(token, path) as { value: Array<Record<string, unknown>> };
  return { messages: (data.value ?? []).map(m => formatMessage(m, true)), count: data.value?.length ?? 0 };
}

export async function searchMail(
  config: OutlookMailConfig,
  params: { query?: string; from?: string; subject?: string; since?: string; before?: string; limit?: number },
): Promise<unknown> {
  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  const token = await getGraphAccessToken(config, SCOPES.mailRead);
  const limit = params.limit ?? 10;
  const filters: string[] = [];
  if (params.from) filters.push(`from/emailAddress/address eq '${esc(String(params.from))}'`);
  if (params.subject) filters.push(`contains(subject,'${esc(String(params.subject))}')`);
  if (params.since) filters.push(`receivedDateTime ge ${params.since}T00:00:00Z`);
  if (params.before) filters.push(`receivedDateTime le ${params.before}T00:00:00Z`);
  const base = `/me/messages?$top=${limit}&$select=subject,from,receivedDateTime,isRead,bodyPreview&$orderby=receivedDateTime%20desc`;
  const path = filters.length ? `${base}&$filter=${encodeURIComponent(filters.join(" and "))}` : base;
  const data = await graphGet(token, path) as { value: Array<Record<string, unknown>> };
  return { messages: (data.value ?? []).map(m => formatMessage(m, true)), count: data.value?.length ?? 0 };
}

export async function readMessage(
  config: OutlookMailConfig,
  params: { message_id: string },
): Promise<unknown> {
  const msgId = params.message_id?.trim();
  if (!msgId) return { error: "message_id is required" };
  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  const token = await getGraphAccessToken(config, SCOPES.mailRead);
  const data = await graphGet(token, `/me/messages/${encodeURIComponent(msgId)}`) as Record<string, unknown>;
  const body = (data.body as Record<string, string> | undefined);
  return { ...formatMessage(data), body: body?.content ?? "", content_type: body?.contentType ?? "" };
}

export async function saveAttachments(
  config: OutlookMailConfig,
  params: { message_id: string; output_dir: string; content_types?: string[] },
): Promise<unknown> {
  const { default: fs } = await import("node:fs");
  const { default: path } = await import("node:path");
  const msgId = params.message_id?.trim();
  const outputDir = params.output_dir?.trim();
  if (!msgId) return { error: "message_id is required" };
  if (!outputDir) return { error: "output_dir is required" };
  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  const token = await getGraphAccessToken(config, SCOPES.mailRead);
  const filters = params.content_types ?? ["image/*"];
  fs.mkdirSync(outputDir, { recursive: true });
  const attData = await graphGet(token, `/me/messages/${encodeURIComponent(msgId)}/attachments`) as { value: Array<Record<string, unknown>> };
  const saved: string[] = [];
  for (const att of attData.value ?? []) {
    const ct = String(att.contentType ?? "");
    const matches = filters.some(f => { if (f.endsWith("/*")) return ct.startsWith(f.slice(0, -1)); return ct === f; });
    if (!matches) continue;
    const name = String(att.name ?? "attachment");
    const safe = path.basename(name).replace(/[/\\]/g, "_") || "attachment";
    const dest = path.join(outputDir, safe);
    const content = String(att.contentBytes ?? "");
    fs.writeFileSync(dest, Buffer.from(content, "base64"));
    saved.push(dest);
  }
  return { saved, count: saved.length };
}

// ---------------------------------------------------------------------------
// Send / Reply / Forward / Move / Flag handlers
// ---------------------------------------------------------------------------

export async function saveDraft(
  config: OutlookMailConfig,
  params: {
    to?: string | string[];
    cc?: string[];
    subject?: string;
    body?: string;
    signature?: string;
    attachment?: string[];
    in_reply_to?: string;
  },
): Promise<unknown> {
  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  const token = await getGraphAccessToken(config, SCOPES.mailWrite);

  const toList = params.to ? (Array.isArray(params.to) ? params.to : [params.to]) : [];
  const bodyText = params.signature ? `${params.body ?? ""}\n\n${params.signature}` : (params.body ?? "");

  const message: Record<string, unknown> = {
    subject: params.subject ?? "",
    body: { contentType: "Text", content: bodyText },
  };
  if (toList.length) {
    message.toRecipients = toList.map(e => ({ emailAddress: { address: e.trim() } }));
  }
  if (params.cc?.length) {
    message.ccRecipients = params.cc.map(e => ({ emailAddress: { address: e.trim() } }));
  }
  if (params.in_reply_to) {
    message.internetMessageHeaders = [{ name: "In-Reply-To", value: params.in_reply_to }];
  }
  if (params.attachment?.length) {
    const { readFile } = await import("node:fs/promises");
    const { basename } = await import("node:path");
    const attachments: Array<Record<string, unknown>> = [];
    for (const filepath of params.attachment) {
      const data = await readFile(filepath);
      attachments.push({
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: basename(filepath),
        contentBytes: data.toString("base64"),
      });
    }
    message.attachments = attachments;
  }

  const res = await graphPost(token, "/me/messages", message) as Record<string, unknown>;
  if (res.error) return { error: JSON.stringify(res.error) };
  return { ok: true, draft_id: res.id, message: `✓ Draft saved: ${params.subject ?? "(no subject)"}` };
}

export async function sendMessage(
  config: OutlookMailConfig,
  params: {
    to: string | string[];
    cc?: string[];
    subject: string;
    body: string;
    signature?: string;
    attachment?: string[];
    in_reply_to?: string;
    references?: string;
  },
): Promise<unknown> {
  if (!params.subject?.trim()) return { error: "subject is required" };
  if (!params.body?.trim()) return { error: "body is required" };

  const toList = Array.isArray(params.to) ? params.to : [params.to];
  if (!toList.length || !toList.some((address) => address?.trim())) return { error: "to is required" };
  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  const token = await getGraphAccessToken(config, SCOPES.mailSend);

  const bodyText = params.signature ? `${params.body}\n\n${params.signature}` : params.body;
  const message: Record<string, unknown> = {
    subject: params.subject,
    body: { contentType: "Text", content: bodyText },
    toRecipients: toList.map(e => ({ emailAddress: { address: e.trim() } })),
  };
  if (params.cc?.length) {
    message.ccRecipients = params.cc.map(e => ({ emailAddress: { address: e.trim() } }));
  }
  if (params.in_reply_to) {
    message.internetMessageHeaders = [{ name: "In-Reply-To", value: params.in_reply_to }];
  }

  if (params.attachment?.length) {
    const { readFile } = await import("node:fs/promises");
    const { basename } = await import("node:path");
    const attachments: Array<Record<string, unknown>> = [];
    for (const filepath of params.attachment) {
      const data = await readFile(filepath);
      attachments.push({
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: basename(filepath),
        contentBytes: data.toString("base64"),
      });
    }
    message.attachments = attachments;
  }

  const sendRes = await graphPost(token, "/me/sendMail", {
    message,
    saveToSentItems: true,
  }) as Record<string, unknown>;
  if (sendRes.error) return { error: JSON.stringify(sendRes.error) };

  const attNote = params.attachment?.length ? ` (${params.attachment.length} attachment(s))` : "";
  return { ok: true, message: `✓ Sent to ${toList.join(", ")}: ${params.subject}${attNote}` };
}

export async function replyToMessage(
  config: OutlookMailConfig,
  params: {
    message_id: string;
    body: string;
    reply_all?: boolean;
    signature?: string;
  },
): Promise<unknown> {
  const msgId = params.message_id?.trim();
  if (!msgId) return { error: "message_id is required" };
  if (!params.body?.trim()) return { error: "body is required" };

  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  const token = await getGraphAccessToken(config, SCOPES.mailSend);
  const bodyText = params.signature ? `${params.body}\n\n${params.signature}` : params.body;
  const endpoint = params.reply_all
    ? `/me/messages/${encodeURIComponent(msgId)}/replyAll`
    : `/me/messages/${encodeURIComponent(msgId)}/reply`;

  const payload = {
    message: { body: { contentType: "Text", content: bodyText } },
    comment: bodyText,
  };
  const res = await graphPost(token, endpoint, payload) as Record<string, unknown>;
  if (res.error) return { error: JSON.stringify(res.error) };
  return { ok: true, message: `✓ Reply sent${params.reply_all ? " (reply-all)" : ""}` };
}

export async function forwardMessage(
  config: OutlookMailConfig,
  params: {
    message_id: string;
    to: string | string[];
    comment?: string;
  },
): Promise<unknown> {
  const msgId = params.message_id?.trim();
  if (!msgId) return { error: "message_id is required" };
  const toList = Array.isArray(params.to) ? params.to : [params.to];
  if (!toList.length || !toList.some((address) => address?.trim())) return { error: "to is required" };

  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  const token = await getGraphAccessToken(config, SCOPES.mailSend);
  const payload: Record<string, unknown> = {
    toRecipients: toList.map(e => ({ emailAddress: { address: e.trim() } })),
  };
  if (params.comment) payload.comment = params.comment;

  const res = await graphPost(token, `/me/messages/${encodeURIComponent(msgId)}/forward`, payload) as Record<string, unknown>;
  if (res.error) return { error: JSON.stringify(res.error) };
  return { ok: true, message: `✓ Forwarded to ${toList.join(", ")}` };
}

export async function moveMessage(
  config: OutlookMailConfig,
  params: { message_id: string; destination_folder: string },
): Promise<unknown> {
  const msgId = params.message_id?.trim();
  if (!msgId) return { error: "message_id is required" };
  if (!params.destination_folder?.trim()) return { error: "destination_folder is required" };

  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  const token = await getGraphAccessToken(config, SCOPES.mailWrite);
  const res = await graphPost(token, `/me/messages/${encodeURIComponent(msgId)}/move`, { destinationId: params.destination_folder }) as Record<string, unknown>;
  if (res.error) return { error: JSON.stringify(res.error) };
  return { ok: true, new_id: res.id ?? null };
}

export async function flagMessage(
  config: OutlookMailConfig,
  params: { message_id: string; flag_status: "flagged" | "complete" | "notFlagged" },
): Promise<unknown> {
  const msgId = params.message_id?.trim();
  if (!msgId) return { error: "message_id is required" };

  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  const token = await getGraphAccessToken(config, SCOPES.mailWrite);
  const res = await graphPatch(token, `/me/messages/${encodeURIComponent(msgId)}`, { flag: { flagStatus: params.flag_status } }) as Record<string, unknown>;
  if (res.error) return { error: JSON.stringify(res.error) };
  return { ok: true, flag_status: params.flag_status };
}

// ---------------------------------------------------------------------------
// markMessage — patch read/flag/importance in one call
// ---------------------------------------------------------------------------

export async function markMessage(
  config: OutlookMailConfig,
  params: {
    message_id: string;
    read?: boolean;
    flag_status?: "flagged" | "notFlagged" | "complete";
    importance?: "normal" | "high" | "low";
  },
): Promise<unknown> {
  const { clientId, clientSecret, refreshToken } = config;
  if (!clientId) return { error: "OUTLOOK_CLIENT_ID not set" };
  const msgId = params.message_id?.trim();
  if (!msgId) return { error: "message_id is required" };

  const patch: Record<string, unknown> = {};
  if (params.read !== undefined) patch.isRead = params.read;
  if (params.flag_status !== undefined) patch.flag = { flagStatus: params.flag_status };
  if (params.importance !== undefined) patch.importance = params.importance;
  if (Object.keys(patch).length === 0) return { error: "At least one of read, flag_status, or importance must be provided" };

  const token = await getAccessToken(clientId, clientSecret, refreshToken);
  const res = await graphPatch(token, `/me/messages/${encodeURIComponent(msgId)}`, patch) as Record<string, unknown>;
  if (res.error) return { error: JSON.stringify(res.error) };
  return { ok: true, ...patch };
}

// ---------------------------------------------------------------------------
// labelMessage — add/remove Outlook categories
// ---------------------------------------------------------------------------

export async function labelMessage(
  config: OutlookMailConfig,
  params: {
    message_id: string;
    add?: string[];
    remove?: string[];
  },
): Promise<unknown> {
  const { clientId, clientSecret, refreshToken } = config;
  if (!clientId) return { error: "OUTLOOK_CLIENT_ID not set" };
  const msgId = params.message_id?.trim();
  if (!msgId) return { error: "message_id is required" };
  if (!params.add?.length && !params.remove?.length) return { error: "At least one of add or remove must be provided" };

  const token = await getAccessToken(clientId, clientSecret, refreshToken);

  // Fetch current categories
  const current = await graphGet(token, `/me/messages/${encodeURIComponent(msgId)}?$select=categories`) as Record<string, unknown>;
  const currentCats: string[] = (current.categories as string[] | undefined) ?? [];

  // Compute updated set (case-insensitive deduplication)
  const removeSet = new Set((params.remove ?? []).map(c => c.toLowerCase()));
  const kept = currentCats.filter(c => !removeSet.has(c.toLowerCase()));
  const keptLower = new Set(kept.map(c => c.toLowerCase()));
  const toAdd = (params.add ?? []).filter(c => !keptLower.has(c.toLowerCase()));
  const newCats = [...kept, ...toAdd];

  const res = await graphPatch(token, `/me/messages/${encodeURIComponent(msgId)}`, { categories: newCats }) as Record<string, unknown>;
  if (res.error) return { error: JSON.stringify(res.error) };
  return { ok: true, categories: newCats };
}

// ---------------------------------------------------------------------------
// deleteMessage — soft or permanent delete
// ---------------------------------------------------------------------------

export async function deleteMessage(
  config: OutlookMailConfig,
  params: { message_id: string; permanent?: boolean },
): Promise<unknown> {
  const { clientId, clientSecret, refreshToken } = config;
  if (!clientId) return { error: "OUTLOOK_CLIENT_ID not set" };
  const msgId = params.message_id?.trim();
  if (!msgId) return { error: "message_id is required" };

  const token = await getAccessToken(clientId, clientSecret, refreshToken);

  if (params.permanent) {
    // Hard delete — irrecoverable
    const res = await graphPost(token, `/me/messages/${encodeURIComponent(msgId)}/permanentDelete`, {}) as Record<string, unknown>;
    if (res.error) return { error: JSON.stringify(res.error) };
    return { ok: true, permanent: true, message_id: msgId };
  } else {
    // Soft delete — moves to Deleted Items
    const res = await httpRequest("DELETE", `${GRAPH_BASE}/me/messages/${encodeURIComponent(msgId)}`, token);
    if (res.status === 204 || (res.status >= 200 && res.status < 300)) {
      return { ok: true, permanent: false, message_id: msgId };
    }
    const err = (() => { try { return JSON.parse(res.data ?? "{}"); } catch { return {}; } })();
    return { error: `Graph API error ${res.status}: ${(err as Record<string, Record<string, string>>)?.error?.message ?? res.data}` };
  }
}

// ---------------------------------------------------------------------------
// listOrCreateFolder — list mail folders or create one
// ---------------------------------------------------------------------------

export async function listOrCreateFolder(
  config: OutlookMailConfig,
  params: { create_name?: string },
): Promise<unknown> {
  const { clientId, clientSecret, refreshToken } = config;
  if (!clientId) return { error: "OUTLOOK_CLIENT_ID not set" };

  const token = await getAccessToken(clientId, clientSecret, refreshToken);

  if (params.create_name?.trim()) {
    const res = await graphPost(token, "/me/mailFolders", { displayName: params.create_name.trim() }) as Record<string, unknown>;
    if (res.error) return { error: JSON.stringify(res.error) };
    return {
      ok: true,
      created: true,
      folder: {
        id: String(res.id ?? ""),
        name: String(res.displayName ?? params.create_name),
        total_items: Number(res.totalItemCount ?? 0),
        unread_items: Number(res.unreadItemCount ?? 0),
      },
    };
  }

  // List all top-level mail folders
  const data = await graphGet(token, "/me/mailFolders?$top=100&$select=id,displayName,totalItemCount,unreadItemCount") as { value: Array<Record<string, unknown>> };
  const folders = (data.value ?? []).map(f => ({
    id: String(f.id ?? ""),
    name: String(f.displayName ?? ""),
    total_items: Number(f.totalItemCount ?? 0),
    unread_items: Number(f.unreadItemCount ?? 0),
  }));
  return { count: folders.length, folders };
}

// ---------------------------------------------------------------------------
// To Do / Tasks handlers
// ---------------------------------------------------------------------------

export interface ListTasksParams {
  task_list?: string;
}

export async function listTaskListsHandler(
  config: OutlookMailConfig,
): Promise<unknown> {
  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  const token = await getGraphAccessToken(config, taskReadScopes(config));
  const lists = await listTaskLists(token);
  return {
    count: lists.length,
    lists: lists.map(list => ({
      id: String(list.id ?? ""),
      display_name: String(list.displayName ?? ""),
      is_default: Boolean(list.wellknownListName || list.isDefault),
    })),
  };
}

export async function listTasks(
  config: OutlookMailConfig,
  params: { task_list?: string; include_completed?: boolean; limit?: number },
): Promise<unknown> {
  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  const token = await getGraphAccessToken(config, taskReadScopes(config));
  const resolved = await resolveTaskList(token, params.task_list);
  const data = await graphGet(token, `/me/todo/lists/${encodeURIComponent(resolved.id)}/tasks?$top=100`) as { value: Array<Record<string, unknown>> };
  let tasks = (data.value ?? []).map(formatTask);
  if (!params.include_completed) {
    tasks = tasks.filter(t => String((t as Record<string, unknown>).status ?? "") !== "completed");
  }
  const limit = params.limit ?? 20;
  return {
    list_id: resolved.id,
    list_name: resolved.name,
    count: tasks.slice(0, limit).length,
    tasks: tasks.slice(0, limit),
  };
}

export interface CreateTaskParams {
  title: string;
  task_list?: string;
  due?: string;
  reminder?: string;
  notes?: string;
  importance?: "low" | "normal" | "high";
}

export async function createTask(
  config: OutlookMailConfig,
  params: CreateTaskParams,
): Promise<unknown> {
  if (!params.title?.trim()) return { error: "title is required" };
  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  const token = await getGraphAccessToken(config, SCOPES.tasksWrite);
  const resolved = await resolveTaskList(token, params.task_list);
  const payload: Record<string, unknown> = { title: params.title.trim() };
  if (params.notes) payload.body = { contentType: "text", content: params.notes };
  if (params.importance) payload.importance = params.importance;
  if (params.due) payload.dueDateTime = toGraphDateTime(params.due, "America/Los_Angeles");
  if (params.reminder) payload.reminderDateTime = toGraphDateTime(params.reminder, "America/Los_Angeles");
  const res = await graphPost(token, `/me/todo/lists/${encodeURIComponent(resolved.id)}/tasks`, payload) as Record<string, unknown>;
  if (res.error) return { error: JSON.stringify(res.error) };
  return {
    ok: true,
    list_id: resolved.id,
    list_name: resolved.name,
    task: formatTask(res),
  };
}

export interface UpdateTaskParams {
  task_id: string;
  task_list?: string;
  title?: string;
  due?: string;
  reminder?: string;
  notes?: string;
  importance?: "low" | "normal" | "high";
  status?: "notStarted" | "inProgress" | "completed";
}

export async function updateTask(
  config: OutlookMailConfig,
  params: UpdateTaskParams,
): Promise<unknown> {
  if (!params.task_id?.trim()) return { error: "task_id is required" };
  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  const token = await getGraphAccessToken(config, SCOPES.tasksWrite);
  const resolved = await resolveTaskList(token, params.task_list);
  const payload: Record<string, unknown> = {};
  if (params.title !== undefined) payload.title = params.title.trim();
  if (params.notes !== undefined) payload.body = { contentType: "text", content: params.notes };
  if (params.importance) payload.importance = params.importance;
  if (params.status) payload.status = params.status;
  if (params.due !== undefined) payload.dueDateTime = params.due ? toGraphDateTime(params.due, "America/Los_Angeles") : null;
  if (params.reminder !== undefined) payload.reminderDateTime = params.reminder ? toGraphDateTime(params.reminder, "America/Los_Angeles") : null;
  if (!Object.keys(payload).length) return { error: "No fields to update were provided" };
  const res = await graphPatch(
    token,
    `/me/todo/lists/${encodeURIComponent(resolved.id)}/tasks/${encodeURIComponent(params.task_id)}`,
    payload,
  ) as Record<string, unknown>;
  if (res.error) return { error: JSON.stringify(res.error) };
  return {
    ok: true,
    list_id: resolved.id,
    list_name: resolved.name,
    task: formatTask(res),
  };
}

export async function completeTask(
  config: OutlookMailConfig,
  params: { task_id: string; task_list?: string },
): Promise<unknown> {
  return updateTask(config, { task_id: params.task_id, task_list: params.task_list, status: "completed" });
}

export async function deleteTask(
  config: OutlookMailConfig,
  params: { task_id: string; task_list?: string },
): Promise<unknown> {
  if (!params.task_id?.trim()) return { error: "task_id is required" };
  const authError = getAuthConfigError(config);
  if (authError) return { error: authError };
  const token = await getGraphAccessToken(config, SCOPES.tasksWrite);
  const resolved = await resolveTaskList(token, params.task_list);
  const res = await httpRequest(
    "DELETE",
    `${GRAPH_BASE}/me/todo/lists/${encodeURIComponent(resolved.id)}/tasks/${encodeURIComponent(params.task_id)}`,
    token,
  );
  if (res.status === 204) {
    return { ok: true, list_id: resolved.id, list_name: resolved.name, task_id: params.task_id };
  }
  if (res.status < 200 || res.status >= 300) {
    const err = (() => { try { return JSON.parse(res.data ?? "{}"); } catch { return {}; } })();
    return { error: `Graph API error ${res.status}: ${err?.error?.message ?? res.data}` };
  }
  return { ok: true, list_id: resolved.id, list_name: resolved.name, task_id: params.task_id };
}
