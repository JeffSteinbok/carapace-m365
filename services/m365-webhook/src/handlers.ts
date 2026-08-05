import type { IncomingMessage, ServerResponse } from "node:http";
import {
  ActionRegistry,
  executeRules,
  type MailEnvelope,
} from "carapace-mail-runtime";
import { log, type MailRule } from "./config.js";
import { dispatchResults } from "./dispatch.js";
import { GraphClient, type GraphMessage } from "./graph.js";
import { OutlookProviderClient } from "./provider.js";
import { StateStore } from "./state.js";

interface GraphNotificationValue {
  clientState?: string;
  resourceData?: { id?: string };
}

interface GraphNotificationBody {
  value?: GraphNotificationValue[];
}

export function messageToEnvelope(message: GraphMessage): MailEnvelope {
  const from = message.from?.emailAddress;
  const headers: Record<string, string> = {};
  let authenticationResults: string | undefined;
  for (const header of message.internetMessageHeaders ?? []) {
    const key = header.name.toLowerCase();
    if (key === "authentication-results") authenticationResults = header.value;
    if (key === "authentication-results" || key === "to" || key === "cc") {
      headers[key] = header.value;
    }
  }

  let authResults: MailEnvelope["auth_results"];
  if (authenticationResults) {
    const extract = (protocol: string): string | undefined => {
      const match = authenticationResults?.match(
        new RegExp(`(?:^|[;\\n])\\s*${protocol}=([a-zA-Z0-9-]+)`, "i"),
      );
      return match?.[1]?.toLowerCase();
    };
    const dkim = extract("dkim");
    const spf = extract("spf");
    const dmarc = extract("dmarc");
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

export async function handleNotification(
  body: string,
  options: {
    graph: GraphClient;
    clientState: string;
    pipelineRules: MailRule[];
    registry: ActionRegistry;
    runtimeConfig: Record<string, unknown>;
    notifyChannel: string;
    notifyTarget: string;
    pipelineWorkspace: string;
    notificationStore?: StateStore;
  },
): Promise<void> {
  let parsed: GraphNotificationBody;
  try {
    parsed = JSON.parse(body) as GraphNotificationBody;
  } catch {
    log("invalid JSON in Graph notification body");
    return;
  }
  const notifications = parsed.value ?? [];
  if (!notifications.length) {
    log("received empty notification batch");
    return;
  }

  const provider = new OutlookProviderClient(options.graph, log);
  for (const notification of notifications) {
    if (notification.clientState !== options.clientState) {
      log("rejected Graph notification with invalid clientState");
      continue;
    }
    const messageId = notification.resourceData?.id;
    if (!messageId) {
      log("Graph notification did not include resourceData.id");
      continue;
    }
    if (options.notificationStore && !options.notificationStore.claimNotification(messageId)) {
      log(`skipping duplicate Graph notification for message ${messageId}`);
      continue;
    }
    try {
      const message = await options.graph.fetchMessage(messageId);
      const envelope = messageToEnvelope(message);
      log(`received: "${envelope.subject}" from ${envelope.sender_email}`);
      const [, results] = await executeRules(
        envelope,
        options.pipelineRules as Record<string, unknown>[],
        options.registry,
        provider,
        {
          workspace: options.pipelineWorkspace,
          logger: log,
          config: options.runtimeConfig,
        },
      );
      dispatchResults(results, {
        channel: options.notifyChannel,
        target: options.notifyTarget,
      });
      options.notificationStore?.completeNotification(messageId);
    } catch (error) {
      options.notificationStore?.releaseNotification(messageId);
      log(`error processing message ${messageId}: ${error}`);
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
