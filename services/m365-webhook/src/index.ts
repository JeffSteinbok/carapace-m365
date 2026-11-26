import { timingSafeEqual } from "node:crypto";
import { watch, type FSWatcher } from "node:fs";
import http, {
  type IncomingMessage,
  type ServerResponse,
} from "node:http";
import { basename, dirname } from "node:path";
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
  normalizeAction,
} from "carapace-mail-runtime";
import {
  buildPipelineRules,
  loadRuntimeConfig,
  buildTrustedForwarders,
  type TrustedForwarder,
  type RuntimeConfig,
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
  parseNotifications,
  clientStateMatches,
} from "./handlers.js";
import { MailStore } from "@carapace/m365-mail-store";
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
    forbiddenStatus?: number;
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
    res.writeHead(options.forbiddenStatus ?? 401, {
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

interface ServerOptions {
  config: Pick<
    ServiceConfig,
    "tokenPath" | "tokenBrokerSecret" | "webhookPath" | "features"
  >;
  tokenSource: TokenSource;
  graph?: GraphClient;
  clientState?: string;
  getPipelineRules?: () => ReturnType<typeof buildPipelineRules>;
  getTrustedForwarders?: () => TrustedForwarder[];
  registry?: ActionRegistry;
  runtimeConfig?: Record<string, unknown>;
  notifyChannel?: string;
  notifyTarget?: string;
  pipelineWorkspace?: string;
  notificationStore?: StateStore;
  mailStore?: MailStore;
  accountId?: string;
  allowedAgents?: string[];
  logger?: typeof log;
}

export function createServiceServer(options: ServerOptions, accounts?: readonly AccountContext[]): http.Server {
  const routes = new Map<string, AccountContext>();
  for (const account of accounts ?? []) {
    const base = options.config.tokenPath.replace(/\/$/, "");
    routes.set(`${base}/${account.accountId}`, account);
    if (account.legacyTokenPath) routes.set(options.config.tokenPath, account);
  }
  return http.createServer(async (req, res) => {
    const url = new URL(req.url ?? "/", "http://localhost");
    const broker = accounts ? routes.get(url.pathname) : url.pathname === options.config.tokenPath ? options : undefined;
    if (broker) {
      await handleTokenRequest(req, res, {
        tokenSource: broker.tokenSource,
        secret: broker.config.tokenBrokerSecret,
        features: broker.config.features,
        forbiddenStatus: accounts && url.pathname !== options.config.tokenPath ? 403 : undefined,
      });
      return;
    }
    if (accounts && (url.pathname === options.config.tokenPath ||
        url.pathname.startsWith(options.config.tokenPath.replace(/\/$/, "") + "/"))) {
      json(res, 403, { error: "Forbidden" });
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
    if (!accounts && !options.graph) {
      res.writeHead(503);
      res.end("Webhook pipeline unavailable");
      return;
    }
    if (!accounts) {
      const bodyPromise = readBody(req);
      res.writeHead(202).end();
      try {
        await runNotification(await bodyPromise, options);
      } catch {
        log("error handling webhook request");
      }
      return;
    }

    let body: string;
    try {
      body = await readBody(req);
    } catch {
      res.writeHead(400).end();
      return;
    }
    const subscriptions = new Map<string, AccountContext | undefined>();
    for (const account of accounts) {
      try {
        const id = account.notificationStore?.load().subscriptionId;
        if (id) subscriptions.set(id, subscriptions.has(id) ? undefined : account);
      } catch {
        account.logger?.("could not read current Graph subscription state");
      }
    }
    const notifications = parseNotifications(body);
    const deliveries = [];
    let retry = false;
    for (const notification of notifications) {
      const id = notification.subscriptionId;
      const account = typeof id === "string" ? subscriptions.get(id) : undefined;
      if (!account) {
        const authenticated = accounts.some(({ clientState }) =>
          clientState !== undefined && clientStateMatches(notification.clientState, clientState));
        if (authenticated) {
          retry = true;
          log("authenticated Graph notification has an unknown or ambiguous subscription; requesting retry");
        } else {
          log("ignored unauthenticated Graph notification with an unknown or ambiguous subscription");
        }
        continue;
      }
      deliveries.push({ notification, account });
    }
    // Graph retries authenticated batches while local subscription state catches
    // up. Known deliveries proceed independently and use per-account deduplication.
    res.writeHead(retry ? 503 : notifications.length ? 202 : 400).end();
    await Promise.all(deliveries.map(async ({ notification, account }) => {
      try {
        await runNotification(JSON.stringify({ value: [notification] }), account);
      } catch {
        account.logger?.("failed to handle Graph notification");
      }
    }));
  });
}

async function runNotification(body: string, options: ServerOptions): Promise<void> {
  if (!options.graph || !options.clientState || !options.getPipelineRules || !options.registry
      || !options.notifyChannel || !options.notifyTarget || !options.pipelineWorkspace) return;
  await handleNotification(body, {
    graph: options.graph,
    clientState: options.clientState,
    pipelineRules: options.getPipelineRules(),
    registry: options.registry,
    runtimeConfig: options.runtimeConfig ?? {},
    notifyChannel: options.notifyChannel,
    notifyTarget: options.notifyTarget,
    pipelineWorkspace: options.pipelineWorkspace,
    notificationStore: options.notificationStore,
    mailStore: options.mailStore,
    trustedForwarders: options.getTrustedForwarders?.() ?? [],
    accountId: options.accountId,
    allowedAgents: options.allowedAgents,
    logger: options.logger,
  });
}

export async function ensureSubscription(
  graph: GraphClient,
  store: StateStore,
  options: {
    webhookUrl: string;
    clientState: string;
    logger?: typeof log;
  },
): Promise<void> {
  const logger = options.logger ?? log;
  const state = store.load();
  const configurationMatches = state.notificationUrl === options.webhookUrl
    && state.clientState === options.clientState;
  const clearSubscription = (): void => {
    store.update({
      subscriptionId: undefined,
      expirationDateTime: undefined,
      notificationUrl: undefined,
      clientState: undefined,
    });
  };

  if (state.subscriptionId && (!state.expirationDateTime || !configurationMatches)) {
    logger("stored Graph subscription configuration changed; replacing subscription");
    try {
      await graph.deleteSubscription(state.subscriptionId);
      logger("stale Graph subscription removed");
    } catch (error) {
      throw new Error(`could not remove stale Graph subscription; refusing replacement: ${error}`);
    }
    clearSubscription();
  } else if (state.subscriptionId && state.expirationDateTime) {
    const remaining = new Date(state.expirationDateTime).getTime() - Date.now();
    if (remaining > RENEWAL_THRESHOLD_MS) {
      logger(`existing Graph subscription remains valid for ${Math.round(remaining / 3_600_000)}h`);
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
        logger(`Graph subscription renewed through ${renewed.expirationDateTime}`);
        return;
      } catch (error) {
        logger(`Graph subscription renewal failed; removing old subscription before replacement: ${error}`);
        await graph.deleteSubscription(state.subscriptionId);
        clearSubscription();
      }
    } else {
      await graph.deleteSubscription(state.subscriptionId);
      clearSubscription();
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
  logger(`Graph subscription active through ${subscription.expirationDateTime}`);
}

export interface AccountContext extends ServerOptions {
  legacyTokenPath?: boolean;
  checkSubscription(): Promise<void>;
  stopWatching(): void;
}

export function validateAgentRoutes(runtime: RuntimeConfig, allowlist: readonly string[]): RuntimeConfig {
  const check = (agent: unknown, session?: unknown): void => {
    if (typeof agent !== "string" || !allowlist.includes(agent) ||
        (session !== undefined && (typeof session !== "string" || !session.startsWith(`agent:${agent}:`)))) {
      throw new Error("Mail rules point outside the account agent allowlist");
    }
  };
  for (const rule of runtime.mail_rules ?? []) {
    for (const action of rule.actions) {
      const [name, params] = normalizeAction(action);
      if (name === "prompt_email") check(params.agent ?? "main", params.session);
    }
  }
  const decisions = runtime.decisions as { agent?: string } | undefined;
  if (decisions) check(decisions.agent ?? "main");
  return runtime;
}

async function createAccountContext(config: ServiceConfig): Promise<AccountContext> {
  const logger = config.logger ?? log;
  const store = new StateStore(config.statePath);
  const state = store.load();
  const refreshToken = state.refreshToken || config.initialRefreshToken;
  if (!refreshToken) {
    throw new Error(
      "No refresh token is available in state or M365_REFRESH_TOKEN",
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
  const readRules = () => {
    const runtime = loadRuntimeConfig(config.runtimeConfigPath, logger);
    return config.accountId ? validateAgentRoutes(runtime, config.allowedAgents ?? []) : runtime;
  };
  let runtimeConfig: RuntimeConfig;
  try {
    runtimeConfig = readRules();
  } catch {
    logger("Refused runtime configuration outside the account allowlist; starting with empty rules");
    runtimeConfig = {};
  }
  const pipelineRules = buildPipelineRules(runtimeConfig);
  const trustedForwarders = buildTrustedForwarders(runtimeConfig, logger);
  const registry = new ActionRegistry();
  registerBuiltinActions(registry, { mailboxPrefixResolver: () => "📬" });

  for (const pluginPath of runtimeConfig.action_plugins ?? []) {
    try {
      const plugin = await import(pluginPath) as ActionPlugin;
      if (typeof plugin.register !== "function") {
        logger(`action plugin ${pluginPath} does not export register(); skipping`);
        continue;
      }
      await plugin.register(registry);
      logger(`loaded action plugin: ${pluginPath}`);
    } catch (error) {
      logger(`failed to load action plugin ${pluginPath}: ${error}`);
    }
  }

  let watcher: FSWatcher | undefined;
  let debounce: ReturnType<typeof setTimeout> | undefined;
  try {
    watcher = watch(dirname(config.runtimeConfigPath), (_event, filename) => {
      if (filename && filename !== basename(config.runtimeConfigPath)) return;
      if (debounce) clearTimeout(debounce);
      debounce = setTimeout(() => {
        let reloaded: RuntimeConfig;
        try {
          reloaded = readRules();
        } catch {
          logger("Refused runtime configuration outside the account allowlist");
          return;
        }
        const updated = buildPipelineRules(reloaded);
        pipelineRules.splice(0, pipelineRules.length, ...updated);
        trustedForwarders.splice(0, trustedForwarders.length, ...buildTrustedForwarders(reloaded, logger));
        logger(
          `runtime config reloaded: ${updated.length} mail rule(s), `
            + `${trustedForwarders.length} trusted forwarder(s)`,
        );
      }, 500);
    });
  } catch {
    // The runtime config is optional and may not exist yet.
  }

  const mailStore = new MailStore(config.mailStorePath);
  if (mailStore.enabled) {
    logger(`mail store enabled at ${config.mailStorePath} (${mailStore.count()} messages)`);
  }

  let checking: Promise<void> | undefined;
  return Object.freeze({
    accountId: config.accountId,
    legacyTokenPath: config.legacyTokenPath,
    logger,
    allowedAgents: config.accountId ? config.allowedAgents ?? [] : undefined,
    config,
    tokenSource: tokens,
    graph,
    clientState: config.webhookClientState,
    getPipelineRules: () => pipelineRules,
    getTrustedForwarders: () => trustedForwarders,
    registry,
    runtimeConfig,
    notifyChannel: config.notifyChannel,
    notifyTarget: config.notifyTarget,
    pipelineWorkspace: config.pipelineWorkspace,
    notificationStore: store,
    mailStore,
    checkSubscription() {
      checking ??= ensureSubscription(graph, store, {
        webhookUrl: config.webhookUrl,
        clientState: config.webhookClientState,
        logger,
      }).catch((error) => logger(`Graph subscription check failed: ${error}`))
        .finally(() => { checking = undefined; });
      return checking;
    },
    stopWatching() {
      if (debounce) clearTimeout(debounce);
      watcher?.close();
    },
  });
}

export async function main(): Promise<void> {
  const { accounts: configs } = loadServiceConfig();
  const config = configs[0]!;
  const accounts = await Promise.all(configs.map(createAccountContext));
  const primary = accounts[0]!;
  const server = createServiceServer(primary, config.accountId ? accounts : undefined);
  await new Promise<void>((resolve, reject) => {
    server.once("error", reject);
    server.listen(config.port, config.bind, () => {
      server.off("error", reject);
      log(`listening on ${config.bind}:${config.port}`);
      resolve();
    });
  });

  const check = () => Promise.all(accounts.map((account) => account.checkSubscription()));
  const renewal = setInterval(() => void check(), RENEWAL_CHECK_INTERVAL_MS);
  const shutdown = (signal: string): void => {
    log(`${signal} received; preserving Graph subscriptions and shutting down`);
    clearInterval(renewal);
    for (const account of accounts) account.stopWatching();
    server.close(() => {
      for (const account of accounts) account.mailStore?.close();
      process.exit(0);
    });
    setTimeout(() => process.exit(0), 5_000).unref();
  };
  process.on("SIGTERM", () => shutdown("SIGTERM"));
  process.on("SIGINT", () => shutdown("SIGINT"));
  await check();
}

const invokedPath = process.argv[1] ? pathToFileURL(process.argv[1]).href : "";
if (import.meta.url === invokedPath) {
  void main().catch((error) => {
    log(`fatal: ${(error as Error).message}`);
    process.exit(1);
  });
}
