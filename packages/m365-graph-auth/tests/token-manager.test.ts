import { describe, expect, it, vi } from "vitest";
import {
  DEFAULT_M365_FEATURES,
  GraphTokenManager,
  deriveAllowedGraphScopes,
  inferM365FeaturesFromScopes,
  isM365FeatureEnabled,
  parseM365Features,
  type HttpRequest,
  type HttpResponse,
} from "../src/index.js";

function response(body: Record<string, unknown>, status = 200): HttpResponse {
  return {
    status,
    headers: {},
    body: Buffer.from(JSON.stringify(body)),
  };
}

describe("Microsoft 365 feature policy", () => {
  it("strictly normalizes features and preserves backward-compatible defaults", () => {
    expect(parseM365Features(undefined)).toEqual(DEFAULT_M365_FEATURES);
    expect(parseM365Features(" MAIL-READ,mail-read,OneDrive-Write ")).toEqual([
      "mail-read",
      "onedrive-write",
    ]);
    expect(() => parseM365Features("mail-read,unknown")).toThrow(
      "Unknown Microsoft 365 feature(s): unknown",
    );
    expect(() => parseM365Features(["mail-read", 42])).toThrow(
      "Microsoft 365 features must be strings",
    );
  });

  it("applies write-to-read implications when deriving allowed scopes", () => {
    expect(isM365FeatureEnabled(["mail-write"], "mail-read")).toBe(true);
    expect(isM365FeatureEnabled(["mail-send"], "mail-read")).toBe(false);
    expect(deriveAllowedGraphScopes(["calendar-write", "onedrive-write"])).toEqual([
      "Calendars.ReadWrite",
      "Files.ReadWrite",
      "Calendars.Read",
      "Files.Read",
    ]);
  });

  it("infers only features represented by exact known scope names", () => {
    expect(inferM365FeaturesFromScopes([
      "mail.read",
      "Files.ReadWrite",
      "User.Read",
    ])).toEqual({
      features: ["mail-read", "onedrive-write"],
      unknownScopes: ["User.Read"],
    });
  });
});

describe("GraphTokenManager", () => {
  it("normalizes scopes and caches access tokens", async () => {
    const request = vi.fn<HttpRequest>().mockResolvedValue(response({
      access_token: "cached-token",
      expires_in: 3600,
    }));
    const manager = new GraphTokenManager({
      clientId: "client",
      refreshToken: "refresh",
      request,
    });

    expect(await manager.getAccessToken(["Mail.Read", "mail.read"])).toBe("cached-token");
    expect(await manager.getAccessToken(["mail.read"])).toBe("cached-token");
    expect(request).toHaveBeenCalledTimes(1);
    expect(String(request.mock.calls[0][0].body)).toContain("scope=mail.read");
    expect(String(request.mock.calls[0][0].body)).not.toContain("client_secret");
  });

  it("serializes refreshes across different scope sets", async () => {
    let active = 0;
    let maxActive = 0;
    const releases: Array<() => void> = [];
    const request: HttpRequest = async () => {
      active += 1;
      maxActive = Math.max(maxActive, active);
      await new Promise<void>((resolve) => releases.push(resolve));
      active -= 1;
      return response({ access_token: `token-${releases.length}`, expires_in: 3600 });
    };
    const manager = new GraphTokenManager({
      clientId: "client",
      refreshToken: "refresh",
      request,
    });

    const first = manager.getAccessToken(["Mail.Read"]);
    const second = manager.getAccessToken(["Files.Read"]);
    await vi.waitFor(() => expect(releases).toHaveLength(1));
    releases.shift()?.();
    await vi.waitFor(() => expect(releases).toHaveLength(1));
    releases.shift()?.();
    await Promise.all([first, second]);
    expect(maxActive).toBe(1);
  });

  it("persists rotated refresh tokens before using them", async () => {
    const persisted: string[] = [];
    const request = vi.fn<HttpRequest>()
      .mockResolvedValueOnce(response({
        access_token: "first",
        expires_in: 1,
        refresh_token: "rotated",
      }))
      .mockImplementationOnce(async (options) => {
        expect(String(options.body)).toContain("refresh_token=rotated");
        return response({ access_token: "second", expires_in: 3600 });
      });
    let now = 1_000;
    const manager = new GraphTokenManager({
      clientId: "client",
      refreshToken: "initial",
      expirySkewMs: 0,
      now: () => now,
      request,
      onRefreshToken: async (token) => {
        persisted.push(token);
      },
    });

    expect(await manager.getAccessToken(["Mail.Read"])).toBe("first");
    now += 1_001;
    expect(await manager.getAccessToken(["Mail.Read"])).toBe("second");
    expect(persisted).toEqual(["rotated"]);
  });

  it("authenticates broker requests and never falls back to direct refresh", async () => {
    const request = vi.fn<HttpRequest>().mockImplementation(async (options) => {
      expect(options.url).toBe("http://127.0.0.1:18790/token");
      expect(options.headers?.Authorization).toBe("Bearer broker-secret");
      expect(JSON.parse(String(options.body))).toEqual({ scopes: ["Mail.Read"] });
      throw new Error("broker unavailable");
    });
    const manager = new GraphTokenManager({
      clientId: "client",
      refreshToken: "direct-refresh",
      tokenBrokerUrl: "http://127.0.0.1:18790/token",
      tokenBrokerSecret: "broker-secret",
      request,
    });

    await expect(manager.getAccessToken(["Mail.Read"])).rejects.toThrow("broker unavailable");
    expect(request).toHaveBeenCalledTimes(1);
  });

  it("uses a still-valid cached token during a broker outage", async () => {
    let now = 1_000;
    const request = vi.fn<HttpRequest>()
      .mockResolvedValueOnce(response({ access_token: "local-cache", expires_in: 30 }))
      .mockRejectedValueOnce(new Error("broker unavailable"));
    const manager = new GraphTokenManager({
      tokenBrokerUrl: "http://127.0.0.1:18790/token",
      tokenBrokerSecret: "broker-secret",
      expirySkewMs: 60_000,
      now: () => now,
      request,
    });

    expect(await manager.getAccessToken(["Mail.Read"])).toBe("local-cache");
    now += 1_000;
    expect(await manager.getAccessToken(["Mail.Read"])).toBe("local-cache");
    expect(request).toHaveBeenCalledTimes(2);
  });

  it("surfaces OAuth errors without exposing request credentials", async () => {
    const request = vi.fn<HttpRequest>().mockResolvedValue(response({
      error: "invalid_grant",
      error_description: "Consent is missing for Files.Read",
    }, 400));
    const manager = new GraphTokenManager({
      clientId: "client",
      clientSecret: "client-secret",
      refreshToken: "refresh-secret",
      request,
    });

    await expect(manager.getAccessToken(["Files.Read"])).rejects.toThrow(
      "invalid_grant — Consent is missing for Files.Read",
    );
  });
});
