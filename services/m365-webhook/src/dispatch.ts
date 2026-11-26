import { execFile } from "node:child_process";
import {
  dispatchResults as dispatchActionResults,
  type ActionResult,
} from "carapace-mail-runtime";
import { log } from "./config.js";

export function deliver(message: string, channel: string, target: string, logger = log): void {
  execFile(
    "openclaw",
    ["message", "send", "--channel", channel, "--target", target, "--message", message],
    { timeout: 30_000, maxBuffer: 256 * 1024, windowsHide: true },
    (error, _stdout, stderr) => {
      if (!error) {
        logger(`delivered: ${message}`);
        return;
      }
      const detail = String(stderr).trim().slice(0, 200) || error.message;
      const timeout = error.killed ? " (timed out after 30s)" : "";
      logger(`delivery failed${timeout}: ${detail}`);
    },
  );
}

/**
 * Hand an email to an agent and deliver its reply.
 *
 * `reply` names where the reply goes. Without it, OpenClaw delivers to the
 * session's own route, which only exists for agents bound to a channel (such
 * as the default agent); a dedicated agent with no bindings fails with
 * "Channel is required when multiple channels are configured".
 */
export function handoffToAgent(
  agent: string,
  message: string,
  sessionKey?: string,
  reply?: { channel: string; target: string },
  logger = log,
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
      ...(reply ? ["--reply-channel", reply.channel, "--reply-to", reply.target] : []),
      "--message",
      message,
    ],
    { timeout: 150_000, maxBuffer: 256 * 1024, windowsHide: true },
    (error, _stdout, stderr) => {
      if (!error) {
        logger(`handoff delivered to agent ${agent}`);
        return;
      }
      const detail = String(stderr).trim().slice(0, 200) || error.message;
      const timeout = error.killed ? " (timed out after 150s)" : "";
      logger(`agent handoff failed${timeout}: ${detail}`);
    },
  );
}

export function dispatchResults(
  results: ActionResult[],
  config: { channel: string; target: string; logger?: typeof log; allowedAgents: readonly string[] | undefined },
): void {
  const logger = config.logger ?? log;
  const permitted = results.filter((result) => {
    if (result.kind !== "agent_handoff" || config.allowedAgents === undefined) return true;
    const agent = result.payload.agent ?? "main";
    const session = result.payload.session;
    if (typeof agent === "string" && config.allowedAgents.includes(agent)
        && (session === undefined || (typeof session === "string" && session.startsWith(`agent:${agent}:`)))) return true;
    logger("refused agent handoff outside the account allowlist");
    return false;
  });
  dispatchActionResults(permitted, {
    logger,
    handlers: {
      message: (payload) => deliver(
        payload.message as string,
        config.channel,
        config.target,
        logger,
      ),
      agent_handoff: (payload) => handoffToAgent(
        (payload.agent as string) ?? "main",
        payload.message as string,
        payload.session as string | undefined,
        { channel: config.channel, target: config.target },
        logger,
      ),
    },
  });
}
