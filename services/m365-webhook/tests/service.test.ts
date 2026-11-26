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
  type AccountContext,
  validateAgentRoutes,
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
    path?: string;
    body?: unknown;
  },
): Promise<{ status: number; body: string }> {
  return new Promise((resolveRequest, reject) => {
    const body = JSON.stringify(options.body ?? {});
    const req = http.request({
      hostname: "127.0.0.1",
      port,
      path: options.path ?? "/token",
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

describe("multiple broker accounts", () => {
  it("isolates bearer credentials, routes, notification state and logs", async () => {
    const account = (id: string, legacyTokenPath: boolean) => {
      const store = new StateStore(resolve(scratch, id + ".json"));
      store.save({ subscriptionId: id + "-subscription" });
      const fetchMessage = vi.fn().mockRejectedValue(new Error(id + "-mail-error"));
      const logger = vi.fn();
      const config = {
        tokenPath: "/token",
        tokenBrokerSecret: id + "-secret",
        webhookPath: "/m365/webhook",
        features: [...DEFAULT_M365_FEATURES],
      };
      const context: AccountContext = {
        config,
        accountId: id,
        legacyTokenPath,
        tokenSource: {
          getToken: vi.fn().mockResolvedValue({ accessToken: id + "-token", expiresAt: Date.now() + 60_000 }),
        },
        clientState: id + "-state",
        graph: { fetchMessage } as unknown as GraphClient,
        getPipelineRules: () => [],
        registry: new ActionRegistry(),
        notifyChannel: "test",
        notifyTarget: id,
        pipelineWorkspace: scratch,
        notificationStore: store,
        logger,
        checkSubscription: async () => {},
        stopWatching: () => {},
      };
      return { context, fetchMessage, logger, store };
    };
    const first = account("first", true);
    const second = account("second", false);
    const server = createServiceServer(first.context, [first.context, second.context]);
    await new Promise<void>((done) => server.listen(0, "127.0.0.1", done));
    const port = (server.address() as { port: number }).port;
    try {
      const body = { scopes: ["Mail.Read"] };
      expect((await callServer(port, { body, path: "/token/second", authorization: "Bearer first-secret" })).status).toBe(403);
      expect((await callServer(port, { body, path: "/token/unknown", authorization: "Bearer first-secret" })).status).toBe(403);
      expect(JSON.parse((await callServer(port, { body, authorization: "Bearer first-secret" })).body).access_token).toBe("first-token");
      expect(JSON.parse((await callServer(port, { body, path: "/token/second", authorization: "Bearer second-secret" })).body).access_token).toBe("second-token");

      const item = (subscriptionId: string, clientState: string, id: string) => ({ subscriptionId, clientState, resourceData: { id } });
      first.store.completeNotification("seen");
      const result = await callServer(port, { path: "/m365/webhook", body: { value: [
        item("unknown", "first-state", "ignore"),
        item("first-subscription", "second-state", "crossed"),
        item("second-subscription", "first-state", "crossed"),
        item("first-subscription", "first-state", "seen"),
        item("second-subscription", "second-state", "seen"),
        item("first-subscription", "first-state", "fresh"),
      ] } });
      expect(result.status).toBe(503);
      await vi.waitFor(() => expect(first.logger).toHaveBeenCalledWith(expect.stringContaining("first-mail-error")));
      expect(first.fetchMessage.mock.calls).toEqual([["fresh"]]);
      expect(second.fetchMessage.mock.calls).toEqual([["seen"]]);
      expect(first.logger).not.toHaveBeenCalledWith(expect.stringContaining("second-mail-error"));
      expect(second.logger).not.toHaveBeenCalledWith(expect.stringContaining("first-mail-error"));

      const sharedLog = vi.spyOn(console, "log").mockImplementation(() => {});
      try {
        const forgedId = "forged-notification-id";
        const forged = await callServer(port, {
          path: "/m365/webhook", body: { value: [item(forgedId, "first-statz", "forged")] },
        });
        expect(forged.status).toBe(202);
        expect(first.fetchMessage).toHaveBeenCalledOnce();
        expect(second.fetchMessage).toHaveBeenCalledOnce();
        expect(sharedLog).toHaveBeenCalledOnce();
        expect(sharedLog.mock.calls[0]?.[0]).not.toContain(forgedId);
        expect(sharedLog.mock.calls[0]?.[0]).not.toContain("first-statz");

        const mixed = await callServer(port, {
          path: "/m365/webhook", body: { value: [
            item(forgedId, "invalid", "ignore"),
            item("first-subscription", "first-state", "independent"),
          ] },
        });
        expect(mixed.status).toBe(202);
        await vi.waitFor(() => expect(first.fetchMessage).toHaveBeenCalledWith("independent"));
      } finally {
        sharedLog.mockRestore();
      }

      first.store.update({ subscriptionId: "duplicated" });
      second.store.update({ subscriptionId: "duplicated" });
      const duplicate = await callServer(port, {
        path: "/m365/webhook", body: { value: [item("duplicated", "first-state", "duplicate")] },
      });
      expect(duplicate.status).toBe(503);
      expect(first.fetchMessage).toHaveBeenCalledTimes(2);
      expect(second.fetchMessage).toHaveBeenCalledOnce();

      second.store.update({ subscriptionId: "new-subscription" });
      const retried = await callServer(port, {
        path: "/m365/webhook", body: { value: [item("new-subscription", "second-state", "retried")] },
      });
      expect(retried.status).toBe(202);
      await vi.waitFor(() => expect(second.fetchMessage).toHaveBeenCalledWith("retried"));
    } finally {
      await new Promise<void>((done) => server.close(() => done()));
    }
  });
});

describe("account rule routes", () => {
  const rules = (agent: string, session: string) => ({
    mail_rules: [{ id: "summary", actions: [{ name: "prompt_email", params: { agent, session } }] }],
  });

  it("allows agent and session routes inside the same account", () => {
    const config = rules("reader", "agent:reader:mail");
    expect(validateAgentRoutes(config, ["reader"])).toBe(config);
  });

  it("refuses prompt actions, session overrides and decisions outside the account", () => {
    expect(() => validateAgentRoutes(rules("main", "agent:main:mail"), ["reader"])).toThrow(/allowlist/);
    expect(() => validateAgentRoutes(rules("reader", "agent:main:mail"), ["reader"])).toThrow(/allowlist/);
    expect(() => validateAgentRoutes({ decisions: { agent: "main" } }, ["reader"])).toThrow(/allowlist/);
  });
});
