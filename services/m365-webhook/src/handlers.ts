import { timingSafeEqual } from "node:crypto";
import type { IncomingMessage, ServerResponse } from "node:http";
import {
  ActionRegistry,
  executeRules,
  type MailEnvelope,
} from "carapace-mail-runtime";
import { log, type MailRule, type TrustedForwarder } from "./config.js";
import { dispatchResults } from "./dispatch.js";
import { GraphClient, type GraphMessage } from "./graph.js";
import { MailStore } from "@carapace/m365-mail-store";
import { OutlookProviderClient } from "./provider.js";
import { StateStore } from "./state.js";

export interface GraphNotificationValue {
  subscriptionId?: string;
  clientState?: string;
  resourceData?: { id?: string };
}

interface GraphNotificationBody {
  value?: GraphNotificationValue[];
}

function extractAuthResult(header: string, protocol: string): string | undefined {
  const match = header.match(
    new RegExp(`(?:^|[;\\n])\\s*${protocol}=([a-zA-Z0-9-]+)`, "i"),
  );
  return match?.[1]?.toLowerCase();
}

function authservId(header: string): string {
  return header.split(";", 1)[0].trim().split(/\s+/)[0] ?? "";
}

function isDomainOrSubdomain(host: string, domain: string): boolean {
  const low = host.toLowerCase();
  return low === domain || low.endsWith(`.${domain}`);
}

/**
 * The verdict a trusted forwarder recorded before relaying mail here.
 *
 * Forwarding with an SRS envelope sender breaks DMARC alignment at Exchange
 * for unsigned mail, but the forwarder's own verdict survives as
 * Authentication-Results-Original. A sender could also write such a header, so
 * it is trusted only when Exchange's SPF check proves the message arrived from
 * the forwarder's SRS domain. Forwarders prepend their headers, so only the
 * leading run stamped by that forwarder counts, and the first one carrying a
 * DMARC result is its own.
 */
function forwardedAuth(
  exchangeResults: string | undefined,
  originals: string[],
  forwarders: TrustedForwarder[],
): MailEnvelope["auth_results"] {
  if (!exchangeResults || !forwarders.length) return undefined;
  if (extractAuthResult(exchangeResults, "spf") !== "pass") return undefined;
  const mailFrom = exchangeResults.match(/\bsmtp\.mailfrom=([^\s;]+)/i)?.[1] ?? "";
  const mailFromDomain = mailFrom.split("@").pop()!.toLowerCase();
  const forwarder = forwarders.find((f) => f.srs_domain === mailFromDomain);
  if (!forwarder) return undefined;

  for (const original of originals) {
    if (!isDomainOrSubdomain(authservId(original), forwarder.authserv_id)) break;
    const dmarc = extractAuthResult(original, "dmarc");
    if (!dmarc) continue;
    return {
      dkim: extractAuthResult(original, "dkim"),
      spf: extractAuthResult(original, "spf"),
      dmarc,
      raw: original,
    };
  }
  return undefined;
}

export function messageToEnvelope(
  message: GraphMessage,
  trustedForwarders: TrustedForwarder[] = [],
): MailEnvelope {
  const from = message.from?.emailAddress;
  const headers: Record<string, string> = {};
  let authenticationResults: string | undefined;
  const originalResults: string[] = [];
  for (const header of message.internetMessageHeaders ?? []) {
    const key = header.name.toLowerCase();
    // Exchange renames inbound copies to Authentication-Results-Original, so
    // this is its own verdict; take the first in case more than one appears.
    if (key === "authentication-results" && authenticationResults === undefined) {
      authenticationResults = header.value;
      headers[key] = header.value;
    }
    if (key === "authentication-results-original") {
      originalResults.push(header.value);
    }
    if (key === "to" || key === "cc") {
      headers[key] = header.value;
    }
  }

  let authResults = forwardedAuth(authenticationResults, originalResults, trustedForwarders);
  if (!authResults && authenticationResults) {
    const dkim = extractAuthResult(authenticationResults, "dkim");
    const spf = extractAuthResult(authenticationResults, "spf");
    const dmarc = extractAuthResult(authenticationResults, "dmarc");
    if (dkim || spf || dmarc) {
      authResults = { dkim, spf, dmarc, raw: authenticationResults };
    }
  }

  const contentType = message.body?.contentType?.toLowerCase();
  return {
    message_id: message.id,
    provider: "outlook",
    account_id: "m365",
    mailbox_id: "inbox",
    sender_name: from?.name ?? "",
    sender_email: from?.address ?? "unknown",
    subject: (message.subject ?? "(no subject)").slice(0, 150),
    received_at: message.receivedDateTime ?? null,
    body_text: contentType === "text"
      ? (message.body?.content ?? null)
      : (message.bodyPreview ?? null),
    body_html: contentType === "html" ? (message.body?.content ?? null) : null,
    has_attachments: message.hasAttachments ?? false,
    auth_results: authResults,
    headers: Object.keys(headers).length ? headers : undefined,
    raw: message as unknown as Record<string, unknown>,
  };
}

export function handleValidation(
  req: IncomingMessage,
  res: ServerResponse,
): boolean {
  const url = new URL(req.url ?? "/", "http://localhost");
  const token = url.searchParams.get("validationToken");
  if (!token) return false;
  res.writeHead(200, { "Content-Type": "text/plain" });
  res.end(token);
  log("completed Microsoft Graph validation handshake");
  return true;
}

export function clientStateMatches(value: unknown, expected: string): boolean {
  if (typeof value !== "string" || !expected) return false;
  const provided = Buffer.from(value);
  const wanted = Buffer.from(expected);
  return provided.length === wanted.length && timingSafeEqual(provided, wanted);
}

export async function handleNotification(
  body: string,
  options: {
    graph: GraphClient;
    clientState: string;
    accountId?: string;
    logger?: typeof log;
    allowedAgents?: string[];
    pipelineRules: MailRule[];
    registry: ActionRegistry;
    runtimeConfig: Record<string, unknown>;
    notifyChannel: string;
    notifyTarget: string;
    pipelineWorkspace: string;
    notificationStore?: StateStore;
    mailStore?: MailStore;
    trustedForwarders?: TrustedForwarder[];
  },
): Promise<void> {
  const logger = options.logger ?? log;
  const notifications = parseNotifications(body, logger);
  const provider = new OutlookProviderClient(options.graph, logger);
  for (const notification of notifications) {
    if (!clientStateMatches(notification.clientState, options.clientState)) {
      logger("rejected Graph notification with invalid clientState");
      continue;
    }
    const messageId = notification.resourceData?.id;
    if (!messageId) {
      logger("Graph notification did not include resourceData.id");
      continue;
    }
    if (options.notificationStore && !options.notificationStore.claimNotification(messageId)) {
      logger(`skipping duplicate Graph notification for message ${messageId}`);
      continue;
    }
    try {
      const message = await options.graph.fetchMessage(messageId);
      const envelope = messageToEnvelope(message, options.trustedForwarders);
      if (options.accountId) envelope.account_id = options.accountId;
      logger(`received: "${envelope.subject}" from ${envelope.sender_email}`);

      if (options.mailStore?.enabled) {
        const internetMessageId = message.internetMessageHeaders
          ?.find((h) => h.name.toLowerCase() === "message-id")
          ?.value ?? null;
        options.mailStore.store({
          graph_id: message.id,
          internet_message_id: internetMessageId,
          subject: envelope.subject,
          sender_email: envelope.sender_email,
          sender_name: envelope.sender_name,
          received_at: envelope.received_at ?? null,
          has_attachments: envelope.has_attachments ? 1 : 0,
          body_preview: message.bodyPreview ?? null,
          raw_json: JSON.stringify(message),
          stored_at: new Date().toISOString(),
        });
        logger(`mail-store: stored message ${message.id}`);
      }

      const [, results] = await executeRules(
        envelope,
        options.pipelineRules as Record<string, unknown>[],
        options.registry,
        provider,
        {
          workspace: options.pipelineWorkspace,
          logger,
          config: options.runtimeConfig,
        },
      );
      dispatchResults(results, {
        channel: options.notifyChannel,
        target: options.notifyTarget,
        logger,
        allowedAgents: options.allowedAgents,
      });
      options.notificationStore?.completeNotification(messageId);
    } catch (error) {
      options.notificationStore?.releaseNotification(messageId);
      logger(`error processing message ${messageId}: ${error}`);
    }
  }
}

export function readBody(
  req: IncomingMessage,
  maxBytes = 1_000_000,
): Promise<string> {
  return new Promise((resolve, reject) => {
    let data = "";
    req.on("data", (chunk: Buffer | string) => {
      data += chunk.toString();
      if (Buffer.byteLength(data) > maxBytes) {
        req.destroy();
        reject(new Error("request body too large"));
      }
    });
    req.on("end", () => resolve(data));
    req.on("error", reject);
  });
}

export function parseNotifications(body: string, logger = log): GraphNotificationValue[] {
  try {
    const parsed = JSON.parse(body) as GraphNotificationBody | null;
    if (Array.isArray(parsed?.value)) {
      return parsed.value.filter((notification) => notification && typeof notification === "object");
    }
  } catch {
    logger("invalid JSON in Graph notification body");
    return [];
  }
  logger("received empty notification batch");
  return [];
}
