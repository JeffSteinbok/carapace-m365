import http from "node:http";
import https from "node:https";
import { M365_FEATURE_POLICY } from "./feature-policy.js";

export {
  DEFAULT_M365_FEATURES,
  M365_FEATURE_POLICY,
  M365_FEATURES,
  UnknownM365FeatureError,
  deriveAllowedGraphScopes,
  deriveFeatureGraphScopes,
  expandM365Features,
  getDisallowedGraphScopes,
  inferM365FeaturesFromScopes,
  isM365FeatureEnabled,
  normalizeGraphScopeNames,
  normalizeM365Features,
  parseM365Features,
} from "./feature-policy.js";

export type M365Feature = keyof typeof M365_FEATURE_POLICY;

export const DEFAULT_TOKEN_BROKER_URL = "http://127.0.0.1:18790/token";
const tokenUrlForTenant = (tenant: string): string =>
  `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`;

export interface HttpResponse {
  status: number;
  headers: http.IncomingHttpHeaders;
  body: Buffer;
}

export interface HttpRequestOptions {
  method: string;
  url: string;
  headers?: Record<string, string>;
  body?: string | Buffer;
  timeoutMs?: number;
}

export type HttpRequest = (options: HttpRequestOptions) => Promise<HttpResponse>;

export interface GraphTokenManagerOptions {
  clientId?: string;
  clientSecret?: string;
  refreshToken?: string;
  tenant?: string;
  tokenUrl?: string;
  tokenBrokerUrl?: string;
  tokenBrokerSecret?: string;
  expirySkewMs?: number;
  request?: HttpRequest;
  now?: () => number;
  onRefreshToken?: (refreshToken: string) => void | Promise<void>;
}

export interface AccessToken {
  accessToken: string;
  expiresAt: number;
}

interface TokenResponse {
  access_token?: string;
  expires_in?: number;
  refresh_token?: string;
  error?: string;
  error_description?: string;
}

export function normalizeScopes(scopes: readonly string[]): string[] {
  const unique = new Map<string, string>();
  for (const value of scopes) {
    const scope = value.trim();
    if (scope) unique.set(scope.toLowerCase(), scope);
  }
  return [...unique.entries()]
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([, scope]) => scope);
}

export const nodeHttpRequest: HttpRequest = ({
  method,
  url,
  headers = {},
  body,
  timeoutMs = 30_000,
}) => new Promise((resolve, reject) => {
  const parsed = new URL(url);
  const transport = parsed.protocol === "http:" ? http : https;
  const requestHeaders = { ...headers };
  if (body !== undefined && requestHeaders["Content-Length"] === undefined) {
    requestHeaders["Content-Length"] = String(Buffer.byteLength(body));
  }
  const req = transport.request(parsed, {
    method,
    headers: requestHeaders,
    timeout: timeoutMs,
  }, (res) => {
    const chunks: Buffer[] = [];
    res.on("data", (chunk: Buffer | string) => {
      chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    });
    res.on("end", () => {
      resolve({
        status: res.statusCode ?? 0,
        headers: res.headers,
        body: Buffer.concat(chunks),
      });
    });
  });
  req.on("error", reject);
  req.on("timeout", () => {
    req.destroy();
    reject(new Error(`HTTP request timed out after ${timeoutMs}ms`));
  });
  if (body !== undefined) req.write(body);
  req.end();
});

function parseTokenResponse(response: HttpResponse, source: string): TokenResponse {
  let parsed: TokenResponse;
  try {
    parsed = JSON.parse(response.body.toString("utf8")) as TokenResponse;
  } catch {
    throw new Error(`${source} returned HTTP ${response.status} with an invalid JSON response`);
  }
  if (response.status < 200 || response.status >= 300 || parsed.error) {
    const detail = parsed.error_description
      ? `${parsed.error ?? "request_failed"} — ${parsed.error_description}`
      : (parsed.error ?? "request_failed");
    throw new Error(`${source} failed (HTTP ${response.status}): ${detail}`);
  }
  if (!parsed.access_token) {
    throw new Error(`${source} returned HTTP ${response.status} without an access_token`);
  }
  return parsed;
}

function validateBrokerUrl(value: string): string {
  const url = new URL(value);
  if (url.protocol === "https:") return url.toString();
  const loopback = url.hostname === "127.0.0.1"
    || url.hostname === "::1"
    || url.hostname === "[::1]"
    || url.hostname === "localhost";
  if (url.protocol !== "http:" || !loopback) {
    throw new Error("Token broker URLs must use HTTPS, except for loopback HTTP");
  }
  return url.toString();
}

export class GraphTokenManager {
  private readonly options: GraphTokenManagerOptions;
  private readonly request: HttpRequest;
  private readonly now: () => number;
  private readonly expirySkewMs: number;
  private readonly cache = new Map<string, AccessToken>();
  private refreshToken?: string;
  private mutex: Promise<void> = Promise.resolve();

  constructor(options: GraphTokenManagerOptions) {
    this.options = options;
    this.request = options.request ?? nodeHttpRequest;
    this.now = options.now ?? Date.now;
    this.expirySkewMs = options.expirySkewMs ?? 60_000;
    this.refreshToken = options.refreshToken;
    if (options.tokenBrokerUrl) validateBrokerUrl(options.tokenBrokerUrl);
  }

  async getAccessToken(scopes: readonly string[]): Promise<string> {
    return (await this.getToken(scopes)).accessToken;
  }

  async getToken(scopes: readonly string[]): Promise<AccessToken> {
    const normalized = normalizeScopes(scopes);
    if (normalized.length === 0) {
      throw new Error("At least one Microsoft Graph scope is required");
    }
    const key = normalized.map((scope) => scope.toLowerCase()).join(" ");
    const cached = this.cache.get(key);
    if (cached && cached.expiresAt > this.now() + this.expirySkewMs) {
      return cached;
    }

    return this.runExclusive(async () => {
      const current = this.cache.get(key);
      if (current && current.expiresAt > this.now() + this.expirySkewMs) {
        return current;
      }
      try {
        const token = this.options.tokenBrokerUrl
          ? await this.requestBrokerToken(normalized)
          : await this.refreshDirectly(normalized);
        this.cache.set(key, token);
        return token;
      } catch (error) {
        const stillValid = this.cache.get(key);
        if (stillValid && stillValid.expiresAt > this.now()) {
          return stillValid;
        }
        throw error;
      }
    });
  }

  invalidate(scopes?: readonly string[]): void {
    if (!scopes) {
      this.cache.clear();
      return;
    }
    const key = normalizeScopes(scopes).map((scope) => scope.toLowerCase()).join(" ");
    this.cache.delete(key);
  }

  private async runExclusive<T>(work: () => Promise<T>): Promise<T> {
    const previous = this.mutex;
    let release: () => void = () => {};
    this.mutex = new Promise<void>((resolve) => {
      release = resolve;
    });
    await previous;
    try {
      return await work();
    } finally {
      release();
    }
  }

  private async requestBrokerToken(scopes: string[]): Promise<AccessToken> {
    const secret = this.options.tokenBrokerSecret;
    if (!secret) {
      throw new Error("A token broker secret is required when tokenBrokerUrl is configured");
    }
    const response = await this.request({
      method: "POST",
      url: validateBrokerUrl(this.options.tokenBrokerUrl ?? DEFAULT_TOKEN_BROKER_URL),
      headers: {
        Authorization: `Bearer ${secret}`,
        Accept: "application/json",
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ scopes }),
    });
    const parsed = parseTokenResponse(response, "Microsoft 365 token broker");
    return this.toCachedToken(parsed);
  }

  private async refreshDirectly(scopes: string[]): Promise<AccessToken> {
    if (!this.options.clientId) {
      throw new Error("M365_CLIENT_ID or OUTLOOK_CLIENT_ID must be set for direct token refresh");
    }
    if (!this.refreshToken) {
      throw new Error("M365_REFRESH_TOKEN or OUTLOOK_REFRESH_TOKEN must be set for direct token refresh");
    }
    const params: Record<string, string> = {
      client_id: this.options.clientId,
      refresh_token: this.refreshToken,
      grant_type: "refresh_token",
      scope: scopes.join(" "),
    };
    if (this.options.clientSecret) params.client_secret = this.options.clientSecret;
    const response = await this.request({
      method: "POST",
      url: this.options.tokenUrl ?? tokenUrlForTenant(this.options.tenant ?? "consumers"),
      headers: {
        Accept: "application/json",
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: new URLSearchParams(params).toString(),
    });
    const parsed = parseTokenResponse(response, "Microsoft OAuth token refresh");
    if (parsed.refresh_token && parsed.refresh_token !== this.refreshToken) {
      await this.options.onRefreshToken?.(parsed.refresh_token);
      this.refreshToken = parsed.refresh_token;
    }
    return this.toCachedToken(parsed);
  }

  private toCachedToken(parsed: TokenResponse): AccessToken {
    return {
      accessToken: parsed.access_token as string,
      expiresAt: this.now() + Math.max(1, parsed.expires_in ?? 3600) * 1000,
    };
  }
}
