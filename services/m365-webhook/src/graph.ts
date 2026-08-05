import https from "node:https";
import type { GraphTokenManager } from "@carapace/m365-graph-auth";
import { GRAPH_BASE } from "./config.js";

const MAIL_READ = ["Mail.Read"] as const;

export interface GraphMessage {
  id: string;
  subject?: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  receivedDateTime?: string;
  bodyPreview?: string;
  body?: { content?: string; contentType?: string };
  hasAttachments?: boolean;
  internetMessageHeaders?: Array<{ name: string; value: string }>;
}

export interface GraphAttachment {
  id: string;
  name?: string;
  contentType?: string;
  isInline?: boolean;
  contentId?: string;
  contentBytes?: string;
  size?: number;
}

export interface GraphSubscription {
  id: string;
  expirationDateTime: string;
  resource: string;
  changeType: string;
  notificationUrl: string;
  clientState: string;
}

interface GraphResponse {
  status: number;
  body: string;
}

function httpRequest(
  method: string,
  url: string,
  headers: Record<string, string>,
  body?: string,
): Promise<GraphResponse> {
  return new Promise((resolve, reject) => {
    const req = https.request(url, {
      method,
      headers,
      timeout: 30_000,
    }, (res) => {
      let data = "";
      res.on("data", (chunk: Buffer | string) => {
        data += chunk.toString();
      });
      res.on("end", () => resolve({ status: res.statusCode ?? 0, body: data }));
    });
    req.on("error", reject);
    req.on("timeout", () => {
      req.destroy();
      reject(new Error("Microsoft Graph request timed out"));
    });
    if (body !== undefined) req.write(body);
    req.end();
  });
}

function graphError(operation: string, response: GraphResponse): Error {
  let detail = response.body.slice(0, 300);
  try {
    const parsed = JSON.parse(response.body) as {
      error?: { code?: string; message?: string };
    };
    detail = [parsed.error?.code, parsed.error?.message].filter(Boolean).join(": ");
  } catch {
    // Keep the bounded plain-text response.
  }
  return new Error(`${operation} failed (HTTP ${response.status})${detail ? `: ${detail}` : ""}`);
}

export class GraphClient {
  constructor(private readonly tokens: GraphTokenManager) {}

  async fetchMessage(messageId: string): Promise<GraphMessage> {
    const select =
      "id,subject,from,receivedDateTime,bodyPreview,body,hasAttachments,internetMessageHeaders";
    return this.requestJson<GraphMessage>(
      "GET",
      `/me/messages/${encodeURIComponent(messageId)}?$select=${select}`,
      undefined,
      "Graph message fetch",
    );
  }

  async fetchAttachments(messageId: string): Promise<GraphAttachment[]> {
    const response = await this.requestJson<{ value?: GraphAttachment[] }>(
      "GET",
      `/me/messages/${encodeURIComponent(messageId)}/attachments`,
      undefined,
      "Graph attachment fetch",
    );
    return response.value ?? [];
  }

  async createSubscription(options: {
    notificationUrl: string;
    clientState: string;
    expirationDateTime: string;
  }): Promise<GraphSubscription> {
    return this.requestJson<GraphSubscription>(
      "POST",
      "/subscriptions",
      {
        changeType: "created",
        notificationUrl: options.notificationUrl,
        resource: "me/mailFolders/inbox/messages",
        expirationDateTime: options.expirationDateTime,
        clientState: options.clientState,
      },
      "Graph subscription creation",
      201,
    );
  }

  async renewSubscription(
    subscriptionId: string,
    expirationDateTime: string,
  ): Promise<GraphSubscription> {
    return this.requestJson<GraphSubscription>(
      "PATCH",
      `/subscriptions/${encodeURIComponent(subscriptionId)}`,
      { expirationDateTime },
      "Graph subscription renewal",
    );
  }

  async deleteSubscription(subscriptionId: string): Promise<void> {
    const token = await this.tokens.getAccessToken(MAIL_READ);
    const response = await httpRequest(
      "DELETE",
      `${GRAPH_BASE}/subscriptions/${encodeURIComponent(subscriptionId)}`,
      {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      },
    );
    if (response.status === 401) this.tokens.invalidate(MAIL_READ);
    if (response.status !== 204 && response.status !== 404) {
      throw graphError("Graph subscription deletion", response);
    }
  }

  private async requestJson<T>(
    method: string,
    path: string,
    body: unknown,
    operation: string,
    expectedStatus = 200,
  ): Promise<T> {
    const token = await this.tokens.getAccessToken(MAIL_READ);
    const serialized = body === undefined ? undefined : JSON.stringify(body);
    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    };
    if (serialized !== undefined) {
      headers["Content-Type"] = "application/json";
      headers["Content-Length"] = String(Buffer.byteLength(serialized));
    }
    const response = await httpRequest(
      method,
      `${GRAPH_BASE}${path}`,
      headers,
      serialized,
    );
    if (response.status === 401) this.tokens.invalidate(MAIL_READ);
    if (response.status !== expectedStatus) throw graphError(operation, response);
    try {
      return JSON.parse(response.body) as T;
    } catch {
      throw new Error(`${operation} returned invalid JSON`);
    }
  }
}
