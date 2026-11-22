import { afterEach, describe, expect, it, vi } from "vitest";

const execFile = vi.hoisted(() => vi.fn());

vi.mock("node:child_process", () => ({ execFile }));

import { deliver, dispatchResults, handoffToAgent } from "../src/dispatch.js";

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

  it("delivers a handoff reply to the owner's channel when one is given", () => {
    handoffToAgent("mail-reader", "summarize", "agent:mail-reader:imessage:direct:owner", {
      channel: "imessage",
      target: "owner@example.com",
    });

    const args = execFile.mock.calls[0]?.[1] as string[];
    expect(args.slice(args.indexOf("--reply-channel"), args.indexOf("--reply-channel") + 4)).toEqual([
      "--reply-channel",
      "imessage",
      "--reply-to",
      "owner@example.com",
    ]);
    expect(args.at(-2)).toBe("--message");
  });

  it("leaves the route to the session when no reply target is given", () => {
    handoffToAgent("main", "review this");

    expect(execFile.mock.calls[0]?.[1]).not.toContain("--reply-channel");
  });

  it("routes rule handoffs to the configured notify channel", () => {
    dispatchResults(
      [{ kind: "agent_handoff", payload: { agent: "mail-reader", message: "summarize" } }],
      { channel: "imessage", target: "owner@example.com" },
    );

    const args = execFile.mock.calls[0]?.[1] as string[];
    expect(args).toContain("mail-reader");
    expect(args).toContain("--reply-to");
    expect(args).toContain("owner@example.com");
  });
});
