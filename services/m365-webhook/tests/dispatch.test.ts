import { afterEach, describe, expect, it, vi } from "vitest";

const execFile = vi.hoisted(() => vi.fn());

vi.mock("node:child_process", () => ({ execFile }));

import { deliver, handoffToAgent } from "../src/dispatch.js";

afterEach(() => {
  execFile.mockReset();
});

describe("m365 webhook command dispatch", () => {
  it("starts message delivery without blocking the event loop", () => {
    deliver("hello", "discord", "target");

    expect(execFile).toHaveBeenCalledOnce();
    expect(execFile.mock.calls[0]?.[0]).toBe("openclaw");
    expect(execFile.mock.calls[0]?.[1]).toEqual([
      "message",
      "send",
      "--channel",
      "discord",
      "--target",
      "target",
      "--message",
      "hello",
    ]);
    expect(execFile.mock.calls[0]?.[2]).toMatchObject({ timeout: 30_000 });
    expect(execFile.mock.calls[0]?.[3]).toBeTypeOf("function");
  });

  it("keeps agent handoff bounded without waiting synchronously", () => {
    handoffToAgent("main", "review this", "mail:1");

    expect(execFile).toHaveBeenCalledOnce();
    expect(execFile.mock.calls[0]?.[1]).toContain("--session-key");
    expect(execFile.mock.calls[0]?.[2]).toMatchObject({ timeout: 150_000 });
    expect(execFile.mock.calls[0]?.[3]).toBeTypeOf("function");
  });
});
