import { describe, expect, it } from "vitest";
import { calendarContinuationPath } from "../src/graph-url.js";

describe("calendar continuation paths", () => {
  const initial = "/me/calendars/Ab%2Fc%3D/calendarView";
  const error = "invalid continuation";

  it("compares decoded paths with case-insensitive host and version and preserves the wire encoding", () => {
    const next = "/me/calendars(%27Ab%2fc=%27)/calendarView?%24skip=200";
    expect(calendarContinuationPath(`https://GRAPH.MICROSOFT.COM/V1.0${next}`, initial, error)).toBe(next);
  });

  it.each([
    "/beta/me/calendars('Ab%2Fc=')/calendarView",
    "/v1.0/me/calendars('different=')/calendarView",
    "/v1.0/ME/calendars('Ab%2Fc=')/calendarView",
    "/v1.0/me/calendars('Ab%2Fc=')/calendarview",
    "/v1.0/me/calendars('Ab%2Fc=')/events",
    "/v1.0/me/calendars/Ab/c=/calendarView",
    "/v1.0/me/calendars('Ab%ZZc=')/calendarView",
  ])("rejects a different or malformed resource %s", path => {
    expect(() => calendarContinuationPath(`https://graph.microsoft.com${path}`, initial, error)).toThrow(error);
  });
});
