import { afterEach, describe, expect, it, vi } from "vitest";
import http from "node:http";
import { rmSync } from "node:fs";
import { resolve } from "node:path";
import {
  DEFAULT_M365_FEATURES,
  GraphTokenManager,
  type HttpRequest,
} from "@carapace/m365-graph-auth";
import { ActionRegistry } from "carapace-mail-runtime";
import {
  createServiceServer,
  ensureSubscription,
} from "../src/index.js";
import { handleNotification } from "../src/handlers.js";
import type { GraphClient } from "../src/graph.js";
import { StateStore } from "../src/state.js";

const scratch = resolve("services", "m365-webhook", ".test-state");
const validAuthorization = ["Bearer", "expected-secret"].join(" ");

afterEach(() => {
  rmSync(scratch, { recursive: true, force: true });
});

function callServer(
  port: number,
  options: {
    authorization?: string;
    body?: unknown;
  },
): Promise<{ status: number; body: string }> {
  return new Promise((resolveRequest, reject) => {
    const body = JSON.stringify(options.body ?? {});
    const req = http.request({
      hostname: "127.0.0.1",
      port,
      path: "/token",
      method: "POST",
      headers: {
        ...(options.authorization ? { Authorization: options.authorization } : {}),
        "Content-Type": "application/json",
        "Content-Length": String(Buffer.byteLength(body)),
      },
    }, (res) => {
      let responseBody = "";
      res.on("data", (chunk: Buffer | string) => {
        responseBody += chunk.toString();
      });
      res.on("end", () => resolveRequest({
        status: res.statusCode ?? 0,
        body: responseBody,
      }));
    });
    req.on("error", reject);
    req.end(body);
  });
}

describe("m365 webhook token broker", () => {
  it("requires the configured bearer secret", async () => {
    const getToken = vi.fn().mockResolvedValue({
      accessToken: "broker-token",
      expiresAt: Date.now() + 3_600_000,
    });
    const server = createServiceServer({
      config: {
        tokenPath: "/token",
        tokenBrokerSecret: "expected-secret",
        webhookPath: "/m365/webhook",
        features: [...DEFAULT_M365_FEATURES],
      },
      tokenSource: { getToken },
    });
    await new Promise<void>((resolveListen) => server.listen(0, "127.0.0.1", resolveListen));
    const address = server.address();
    const port = typeof address === "object" && address ? address.port : 0;

    const unauthorized = await callServer(port, {
      authorization: "Bearer wrong-secret",
      body: { scopes: ["Mail.Read"] },
    });
    const authorized = await callServer(port, {
      authorization: "Bearer expected-secret",
      body: { scopes: ["Mail.Read"] },
    });
    await new Promise<void>((resolveClose) => server.close(() => resolveClose()));

    expect(unauthorized.status).toBe(401);
    expect(authorized.status).toBe(200);
    expect(JSON.parse(authorized.body).access_token).toBe("broker-token");
    expect(getToken).toHaveBeenCalledOnce();
    expect(getToken).toHaveBeenCalledWith(["Mail.Read"]);
  });

  it("allows only scopes authorized by the service feature allowlist", async () => {
    const getToken = vi.fn().mockResolvedValue({
      accessToken: "broker-token",
      expiresAt: Date.now() + 3_600_000,
    });
    const server = createServiceServer({
      config: {
        tokenPath: "/token",
        tokenBrokerSecret: "expected-secret",
        webhookPath: "/m365/webhook",
        features: ["onedrive-read"],
      },
      tokenSource: { getToken },
    });
    await new Promise<void>((resolveListen) => server.listen(0, "127.0.0.1", resolveListen));
    const address = server.address();
    const port = typeof address === "object" && address ? address.port : 0;

    const allowed = await callServer(port, {
      authorization: validAuthorization,
      body: { scopes: ["files.read"] },
    });
    const denied = await callServer(port, {
      authorization: validAuthorization,
      body: { scopes: ["Files.ReadWrite"] },
    });
    const unknown = await callServer(port, {
      authorization: validAuthorization,
      body: { scopes: ["User.Read"] },
    });
    await new Promise<void>((resolveClose) => server.close(() => resolveClose()));

    expect(allowed.status).toBe(200);
    expect(denied.status).toBe(403);
    expect(JSON.parse(denied.body)).toMatchObject({
      error: "scope_not_allowed",
    });
    expect(unknown.status).toBe(403);
    expect(getToken).toHaveBeenCalledOnce();
    expect(getToken).toHaveBeenCalledWith(["files.read"]);
  });

  it("atomically persists refresh-token rotation without losing subscription state", async () => {
    const path = resolve(scratch, "state.json");
    const store = new StateStore(path);
    store.save({
      refreshToken: "initial",
      subscriptionId: "subscription-1",
      expirationDateTime: "2026-08-07T00:00:00.000Z",
    });
    const request: HttpRequest = async () => ({
      status: 200,
      headers: {},
      body: Buffer.from(JSON.stringify({
        access_token: "access",
        expires_in: 3600,
        refresh_token: "rotated",
      })),
    });
    const manager = new GraphTokenManager({
      clientId: "client",
      refreshToken: "initial",
      request,
      onRefreshToken: async (refreshToken) => {
        store.update({ refreshToken });
      },
    });

    await manager.getAccessToken(["Mail.Read"]);
    expect(store.load()).toEqual({
      refreshToken: "rotated",
      subscriptionId: "subscription-1",
      expirationDateTime: "2026-08-07T00:00:00.000Z",
    });
  });

  it("reuses a valid persisted Graph subscription after restart", async () => {
    const store = new StateStore(resolve(scratch, "subscription-state.json"));
    const state = {
      refreshToken: "refresh",
      subscriptionId: "subscription-1",
      expirationDateTime: new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString(),
      notificationUrl: "https://example.test/m365/webhook",
      clientState: "client-state",
    };
    store.save(state);
    const createSubscription = vi.fn();
    const renewSubscription = vi.fn();
    const deleteSubscription = vi.fn();
    const graph = {
      createSubscription,
      renewSubscription,
      deleteSubscription,
    } as unknown as GraphClient;

    await ensureSubscription(graph, store, {
      webhookUrl: "https://example.test/m365/webhook",
      clientState: "client-state",
    });

    expect(createSubscription).not.toHaveBeenCalled();
    expect(renewSubscription).not.toHaveBeenCalled();
    expect(deleteSubscription).not.toHaveBeenCalled();
    expect(store.load()).toEqual(state);
  });

  it("replaces a persisted Graph subscription when its configuration changes", async () => {
    const store = new StateStore(resolve(scratch, "changed-subscription-state.json"));
    store.save({
      refreshToken: "refresh",
      subscriptionId: "subscription-old",
      expirationDateTime: new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString(),
      notificationUrl: "https://old.example.test/m365/webhook",
      clientState: "old-client-state",
    });
    const deleteSubscription = vi.fn().mockResolvedValue(undefined);
    const renewSubscription = vi.fn();
    const createSubscription = vi.fn().mockResolvedValue({
      id: "subscription-new",
      expirationDateTime: "2026-08-07T00:00:00.000Z",
    });
    const graph = {
      createSubscription,
      renewSubscription,
      deleteSubscription,
    } as unknown as GraphClient;

    await ensureSubscription(graph, store, {
      webhookUrl: "https://new.example.test/m365/webhook",
      clientState: "new-client-state",
    });

    expect(deleteSubscription).toHaveBeenCalledWith("subscription-old");
    expect(renewSubscription).not.toHaveBeenCalled();
    expect(createSubscription).toHaveBeenCalledWith(expect.objectContaining({
      notificationUrl: "https://new.example.test/m365/webhook",
      clientState: "new-client-state",
    }));
    expect(store.load()).toEqual({
      refreshToken: "refresh",
      subscriptionId: "subscription-new",
      expirationDateTime: "2026-08-07T00:00:00.000Z",
      notificationUrl: "https://new.example.test/m365/webhook",
      clientState: "new-client-state",
    });
  });

  it("replaces a legacy Graph subscription without persisted configuration", async () => {
    const store = new StateStore(resolve(scratch, "legacy-subscription-state.json"));
    store.save({
      refreshToken: "refresh",
      subscriptionId: "subscription-legacy",
      expirationDateTime: new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString(),
    });
    const deleteSubscription = vi.fn().mockResolvedValue(undefined);
    const createSubscription = vi.fn().mockResolvedValue({
      id: "subscription-new",
      expirationDateTime: "2026-08-07T00:00:00.000Z",
    });
    const graph = {
      createSubscription,
      renewSubscription: vi.fn(),
      deleteSubscription,
    } as unknown as GraphClient;

    await ensureSubscription(graph, store, {
      webhookUrl: "https://example.test/m365/webhook",
      clientState: "client-state",
    });

    expect(deleteSubscription).toHaveBeenCalledWith("subscription-legacy");
    expect(createSubscription).toHaveBeenCalledOnce();
    expect(store.load()).toEqual({
      refreshToken: "refresh",
      subscriptionId: "subscription-new",
      expirationDateTime: "2026-08-07T00:00:00.000Z",
      notificationUrl: "https://example.test/m365/webhook",
      clientState: "client-state",
    });
  });

  it("does not create a replacement when old subscription cleanup fails", async () => {
    const store = new StateStore(resolve(scratch, "failed-cleanup-state.json"));
    store.save({
      subscriptionId: "subscription-old",
      expirationDateTime: new Date(Date.now() + 24 * 60 * 60 * 1000).toISOString(),
      notificationUrl: "https://old.example.test/m365/webhook",
      clientState: "old-client-state",
    });
    const deleteSubscription = vi.fn().mockRejectedValue(new Error("Graph unavailable"));
    const createSubscription = vi.fn();
    const graph = {
      createSubscription,
      renewSubscription: vi.fn(),
      deleteSubscription,
    } as unknown as GraphClient;

    await expect(ensureSubscription(graph, store, {
      webhookUrl: "https://new.example.test/m365/webhook",
      clientState: "new-client-state",
    })).rejects.toThrow("refusing replacement");

    expect(createSubscription).not.toHaveBeenCalled();
    expect(store.load().subscriptionId).toBe("subscription-old");
  });

  it("deduplicates persisted Graph message notifications", async () => {
    const store = new StateStore(resolve(scratch, "notification-state.json"));
    const fetchMessage = vi.fn().mockResolvedValue({
      id: "message-1",
      subject: "Test message",
      from: { emailAddress: { address: "sender@example.test" } },
    });
    const graph = { fetchMessage } as unknown as GraphClient;
    const notification = JSON.stringify({
      value: [
        { clientState: "client-state", resourceData: { id: "message-1" } },
        { clientState: "client-state", resourceData: { id: "message-1" } },
      ],
    });
    const options = {
      graph,
      clientState: "client-state",
      pipelineRules: [],
      registry: new ActionRegistry(),
      runtimeConfig: {},
      notifyChannel: "test",
      notifyTarget: "target",
      pipelineWorkspace: scratch,
      notificationStore: store,
    };

    await handleNotification(notification, options);
    await handleNotification(notification, options);

    expect(fetchMessage).toHaveBeenCalledOnce();
    expect(store.load().processedNotificationIds).toEqual(["message-1"]);
  });
});
