export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

export function calendarEventPath(eventId: string): string {
  if (typeof eventId !== "string" || !/^[A-Za-z0-9_+=-]+$/.test(eventId)) {
    throw new Error("A single Microsoft Graph event ID is required");
  }
  return `/me/events/${encodeURIComponent(eventId)}`;
}

function graphContinuationUrl(next: string, errorMessage: string): URL {
  let url: URL;
  try {
    url = new URL(next);
  } catch {
    throw new Error(errorMessage);
  }
  if (url.origin !== new URL(GRAPH_BASE).origin || url.username || url.password || url.hash || url.href.endsWith("#")) {
    throw new Error(errorMessage);
  }
  return url;
}

export function continuationPath(next: string, initial: string, errorMessage: string): string {
  const url = graphContinuationUrl(next, errorMessage);
  if (url.pathname !== new URL(`${GRAPH_BASE}${initial}`).pathname) throw new Error(errorMessage);
  return url.pathname.slice("/v1.0".length) + url.search;
}

function normalizeCalendarView(pathname: string): string[] {
  // Decode individual segments so an escaped slash in a calendar ID remains
  // part of that ID. Graph also represents the same ID as an OData string key.
  const segments = pathname.split("/").map(segment => decodeURIComponent(segment));
  const key = /^calendars\('((?:[^']|'')+)'\)$/.exec(segments[3] ?? "");
  if (segments.length === 5 && key) {
    segments.splice(3, 1, "calendars", key[1].replace(/''/g, "'"));
  }
  segments[1] = (segments[1] ?? "").toLowerCase();
  return segments;
}

export function calendarContinuationPath(next: string, initial: string, errorMessage: string): string {
  const url = graphContinuationUrl(next, errorMessage);
  let actual: string[];
  let expected: string[];
  try {
    actual = normalizeCalendarView(url.pathname);
    expected = normalizeCalendarView(new URL(`${GRAPH_BASE}${initial}`).pathname);
  } catch {
    throw new Error(errorMessage);
  }
  if (actual.length !== expected.length || actual.some((segment, index) => segment !== expected[index])) {
    throw new Error(errorMessage);
  }
  const resourceStart = url.pathname.indexOf("/", 1);
  return url.pathname.slice(resourceStart) + url.search;
}
