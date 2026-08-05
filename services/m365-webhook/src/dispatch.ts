import { execFile } from "node:child_process";
import {
  dispatchResults as dispatchActionResults,
  type ActionResult,
} from "carapace-mail-runtime";
import { log } from "./config.js";

export function deliver(message: string, channel: string, target: string): void {
  execFile(
    "openclaw",
    ["message", "send", "--channel", channel, "--target", target, "--message", message],
    { timeout: 30_000, maxBuffer: 256 * 1024, windowsHide: true },
    (error, _stdout, stderr) => {
      if (!error) {
        log(`delivered: ${message}`);
        return;
      }
      const detail = String(stderr).trim().slice(0, 200) || error.message;
      const timeout = error.killed ? " (timed out after 30s)" : "";
      log(`delivery failed${timeout}: ${detail}`);
    },
  );
}

export function handoffToAgent(
  agent: string,
  message: string,
  sessionKey?: string,
): void {
  execFile(
    "openclaw",
    [
      "agent",
      "--agent",
      agent,
      "--json",
      "--deliver",
      "--timeout",
      "120",
      ...(sessionKey ? ["--session-key", sessionKey] : []),
      "--message",
      message,
    ],
    { timeout: 150_000, maxBuffer: 256 * 1024, windowsHide: true },
    (error, _stdout, stderr) => {
      if (!error) {
        log(`handoff delivered to agent ${agent}`);
        return;
      }
      const detail = String(stderr).trim().slice(0, 200) || error.message;
      const timeout = error.killed ? " (timed out after 150s)" : "";
      log(`agent handoff failed${timeout}: ${detail}`);
    },
  );
}

export function dispatchResults(
  results: ActionResult[],
  config: { channel: string; target: string },
): void {
  dispatchActionResults(results, {
    logger: log,
    handlers: {
      message: (payload) => deliver(
        payload.message as string,
        config.channel,
        config.target,
      ),
      agent_handoff: (payload) => handoffToAgent(
        (payload.agent as string) ?? "main",
        payload.message as string,
        payload.session as string | undefined,
      ),
    },
  });
}
