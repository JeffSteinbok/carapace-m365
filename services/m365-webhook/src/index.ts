import { timingSafeEqual } from "node:crypto";
import { watch, type FSWatcher } from "node:fs";
import http, {
  type IncomingMessage,
  type ServerResponse,
} from "node:http";
import { pathToFileURL } from "node:url";
import {
  getDisallowedGraphScopes,
  GraphTokenManager,
  type AccessToken,
} from "@carapace/m365-graph-auth";
import {
  ActionRegistry,
  registerBuiltinActions,
  type ActionPlugin,
} from "carapace-mail-runtime";
import {
  buildPipelineRules,
  loadRuntimeConfig,
  loadServiceConfig,
  log,
  RENEWAL_CHECK_INTERVAL_MS,
  RENEWAL_THRESHOLD_MS,
  SUBSCRIPTION_TTL_MS,
  type ServiceConfig,
} from "./config.js";
import { GraphClient } from "./graph.js";
import {
  handleNotification,
  handleValidation,
  readBody,
} from "./handlers.js";
import { StateStore } from "./state.js";

export interface TokenSource {
  getToken(scopes: readonly string[]): Promise<AccessToken>;
}

function secretMatches(header: string | undefined, expected: string): boolean {
  if (!header?.startsWith("Bearer ")) return false;
  const provided = Buffer.from(header.slice(7), "utf8");
  const wanted = Buffer.from(expected, "utf8");
  return provided.length === wanted.length && timingSafeEqual(provided, wanted);
}

function json(
  res: ServerResponse,
  status: number,
  value: Record<string, unknown>,
): void {
  const body = JSON.stringify(value);
  res.writeHead(status, {
    "Content-Type": "application/json",
    "Content-Length": String(Buffer.byteLength(body)),
    "Cache-Control": "no-store",
  });
  res.end(body);
}

async function handleTokenRequest(
  req: IncomingMessage,
  res: ServerResponse,
  options: {
    tokenSource: TokenSource;
    secret: string;
    features: ServiceConfig["features"];
  },
): Promise<void> {
  if (req.method !== "POST") {
    res.writeHead(405, { Allow: "POST" });
    res.end("Method not allowed");
    return;
  }
  const authorization = Array.isArray(req.headers.authorization)
    ? req.headers.authorization[0]
    : req.headers.authorization;
  if (!secretMatches(authorization, options.secret)) {
    res.writeHead(401, {
      "WWW-Authenticate": "Bearer",
      "Cache-Control": "no-store",
    });
    res.end("Unauthorized");
    return;
  }

  try {
    const body = JSON.parse(await readBody(req, 32_768)) as { scopes?: unknown };
    if (!Array.isArray(body.scopes)
      || body.scopes.length < 1
      || body.scopes.length > 32
      || body.scopes.some((scope) =>
        typeof scope !== "string"
        || !/^[A-Za-z][A-Za-z0-9.]{1,99}$/.test(scope)
      )) {
      json(res, 400, { error: "scopes must be an array of Microsoft Graph delegated scope names" });
      return;
    }
    const scopes = body.scopes as string[];
    const disallowedScopes = getDisallowedGraphScopes(scopes, options.features);
    if (disallowedScopes.length) {
      json(res, 403, {
        error: "scope_not_allowed",
        error_description:
          `Requested Microsoft Graph scope(s) are not permitted by configured features: ${disallowedScopes.join(", ")}`,
      });
      return;
    }
    const token = await options.tokenSource.getToken(scopes);
    json(res, 200, {
      access_token: token.accessToken,
      token_type: "Bearer",
      expires_in: Math.max(1, Math.floor((token.expiresAt - Date.now()) / 1000)),
    });
  } catch (error) {
    json(res, 502, {
      error: "token_request_failed",
      error_description: (error as Error).message,
    });
  }
}

export function createServiceServer(options: {
  config: Pick<
    ServiceConfig,
    "tokenPath" | "tokenBrokerSecret" | "webhookPath" | "features"
  >;
  tokenSource: TokenSource;
  graph?: GraphClient;
  clientState?: string;
  getPipelineRules?: () => ReturnType<typeof buildPipelineRules>;
  registry?: ActionRegistry;
  runtimeConfig?: Record<string, unknown>;
  notifyChannel?: string;
  notifyTarget?: string;
  pipelineWorkspace?: string;
}): http.Server {
  return http.createServer(async (req, res) => {
    const url = new URL(req.url ?? "/", "http://localhost");
    if (url.pathname === options.config.tokenPath) {
      await handleTokenRequest(req, res, {
        tokenSource: options.tokenSource,
        secret: options.config.tokenBrokerSecret,
        features: options.config.features,
      });
      return;
    }
    if (url.pathname !== options.config.webhookPath && url.pathname !== "/") {
      res.writeHead(404);
      res.end("Not found");
      return;
    }
    if (handleValidation(req, res)) return;
    if (req.method !== "POST") {
      res.writeHead(405, { Allow: "POST" });
      res.end("Method not allowed");
      return;
    }
    if (!options.graph
      || !options.clientState
      || !options.getPipelineRules
      || !options.registry
      || !options.notifyChannel
      || !options.notifyTarget
      || !options.pipelineWorkspace) {
      res.writeHead(503);
      res.end("Webhook pipeline unavailable");
      return;
    }
    const bodyPromise = readBody(req);
    res.writeHead(202);
    res.end();
    try {
      await handleNotification(await bodyPromise, {
        graph: options.graph,
        clientState: options.clientState,
        pipelineRules: options.getPipelineRules(),
        registry: options.registry,
        runtimeConfig: options.runtimeConfig ?? {},
        notifyChannel: options.notifyChannel,
        notifyTarget: options.notifyTarget,
        pipelineWorkspace: options.pipelineWorkspace,
      });
    } catch (error) {
      log(`error handling webhook request: ${error}`);
    }
  });
}

export async function ensureSubscription(
  graph: GraphClient,
  store: StateStore,
  options: {
    webhookUrl: string;
    clientState: string;
  },
): Promise<void> {
  const state = store.load();
  const configurationMatches = state.notificationUrl === options.webhookUrl
    && state.clientState === options.clientState;

  if (state.subscriptionId && (!state.expirationDateTime || !configurationMatches)) {
    log("stored Graph subscription configuration changed; replacing subscription");
    try {
      await graph.deleteSubscription(state.subscriptionId);
      log("stale Graph subscription removed");
    } catch (error) {
      log(`could not remove stale Graph subscription; continuing with replacement: ${error}`);
    }
    store.update({
      subscriptionId: undefined,
      expirationDateTime: undefined,
      notificationUrl: undefined,
      clientState: undefined,
    });
  } else if (state.subscriptionId && state.expirationDateTime) {
    const remaining = new Date(state.expirationDateTime).getTime() - Date.now();
    if (remaining > RENEWAL_THRESHOLD_MS) {
      log(`existing Graph subscription remains valid for ${Math.round(remaining / 3_600_000)}h`);
      return;
    }
    if (remaining > 0) {
      try {
        const expiration = new Date(Date.now() + SUBSCRIPTION_TTL_MS).toISOString();
        const renewed = await graph.renewSubscription(state.subscriptionId, expiration);
        store.update({
          subscriptionId: renewed.id,
          expirationDateTime: renewed.expirationDateTime,
          notificationUrl: options.webhookUrl,
          clientState: options.clientState,
        });
        log(`Graph subscription renewed through ${renewed.expirationDateTime}`);
        return;
      } catch (error) {
        log(`Graph subscription renewal failed; creating a replacement: ${error}`);
      }
    }
  }

  const expiration = new Date(Date.now() + SUBSCRIPTION_TTL_MS).toISOString();
  const subscription = await graph.createSubscription({
    notificationUrl: options.webhookUrl,
    clientState: options.clientState,
    expirationDateTime: expiration,
  });
  store.update({
    subscriptionId: subscription.id,
    expirationDateTime: subscription.expirationDateTime,
    notificationUrl: options.webhookUrl,
    clientState: options.clientState,
  });
  log(`Graph subscription active through ${subscription.expirationDateTime}`);
}

export async function main(): Promise<void> {
  const config = loadServiceConfig();
  const store = new StateStore(config.statePath);
  const state = store.load();
  const refreshToken = state.refreshToken || config.initialRefreshToken;
  if (!refreshToken) {
    throw new Error(
      "No refresh token is available in state or M365_REFRESH_TOKEN/OUTLOOK_REFRESH_TOKEN",
    );
  }
  if (!state.refreshToken) store.update({ refreshToken });

  const tokens = new GraphTokenManager({
    clientId: config.clientId,
    clientSecret: config.clientSecret,
    refreshToken,
    tenant: config.tenant,
    onRefreshToken: async (rotatedToken) => {
      store.update({ refreshToken: rotatedToken });
    },
  });
  const graph = new GraphClient(tokens);
  const runtimeConfig = loadRuntimeConfig(config.runtimeConfigPath);
  const pipelineRules = buildPipelineRules(runtimeConfig);
  const registry = new ActionRegistry();
  registerBuiltinActions(registry, { mailboxPrefixResolver: () => "📬" });

  for (const pluginPath of runtimeConfig.action_plugins ?? []) {
    try {
      const plugin = await import(pluginPath) as ActionPlugin;
      if (typeof plugin.register !== "function") {
        log(`action plugin ${pluginPath} does not export register(); skipping`);
        continue;
      }
      await plugin.register(registry);
      log(`loaded action plugin: ${pluginPath}`);
    } catch (error) {
      log(`failed to load action plugin ${pluginPath}: ${error}`);
    }
  }

  let watcher: FSWatcher | undefined;
  let debounce: ReturnType<typeof setTimeout> | undefined;
  try {
    watcher = watch(config.runtimeConfigPath, () => {
      if (debounce) clearTimeout(debounce);
      debounce = setTimeout(() => {
        const updated = buildPipelineRules(loadRuntimeConfig(config.runtimeConfigPath));
        pipelineRules.splice(0, pipelineRules.length, ...updated);
        log(`runtime config reloaded: ${updated.length} mail rule(s)`);
      }, 500);
    });
  } catch {
    // The runtime config is optional and may not exist yet.
  }

  const server = createServiceServer({
    config,
    tokenSource: tokens,
    graph,
    clientState: config.webhookClientState,
    getPipelineRules: () => pipelineRules,
    registry,
    runtimeConfig,
    notifyChannel: config.notifyChannel,
    notifyTarget: config.notifyTarget,
    pipelineWorkspace: config.pipelineWorkspace,
  });
  await new Promise<void>((resolve) => {
    server.listen(config.port, config.bind, () => {
      log(`listening on ${config.bind}:${config.port}`);
      log(`token broker path: ${config.tokenPath}`);
      log(`Graph webhook URL: ${config.webhookUrl}`);
      resolve();
    });
  });

  try {
    await ensureSubscription(graph, store, {
      webhookUrl: config.webhookUrl,
      clientState: config.webhookClientState,
    });
  } catch (error) {
    log(`initial Graph subscription check failed: ${error}`);
  }
  const renewal = setInterval(() => {
    void ensureSubscription(graph, store, {
      webhookUrl: config.webhookUrl,
      clientState: config.webhookClientState,
    }).catch((error) => log(`Graph subscription check failed: ${error}`));
  }, RENEWAL_CHECK_INTERVAL_MS);

  const shutdown = (signal: string): void => {
    log(`${signal} received; preserving Graph subscription and shutting down`);
    clearInterval(renewal);
    if (debounce) clearTimeout(debounce);
    watcher?.close();
    server.close(() => {
      log("server closed");
      process.exit(0);
    });
    setTimeout(() => process.exit(0), 5_000).unref();
  };
  process.on("SIGTERM", () => shutdown("SIGTERM"));
  process.on("SIGINT", () => shutdown("SIGINT"));
}

const invokedPath = process.argv[1] ? pathToFileURL(process.argv[1]).href : "";
if (import.meta.url === invokedPath) {
  void main().catch((error) => {
    log(`fatal: ${(error as Error).message}`);
    process.exit(1);
  });
}
