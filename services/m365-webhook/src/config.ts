import { appendFileSync, existsSync, readFileSync, realpathSync, statSync } from "node:fs";
import { homedir } from "node:os";
import { basename, dirname, join, resolve } from "node:path";
import {
  parseM365Features,
  type M365Feature,
} from "@carapace/m365-graph-auth";

export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
export const SUBSCRIPTION_TTL_MS = 72 * 60 * 60 * 1000;
export const RENEWAL_THRESHOLD_MS = 12 * 60 * 60 * 1000;
export const RENEWAL_CHECK_INTERVAL_MS = 30 * 60 * 1000;

export interface RuntimeConfig {
  mail_rules?: MailRule[];
  action_plugins?: string[];
  trusted_forwarders?: TrustedForwarder[];
  [key: string]: unknown;
}

/**
 * A server that forwards mail into this mailbox and records its own
 * authentication verdict before doing so. See forwardedAuth in handlers.ts
 * for how the verdict is used.
 */
export interface TrustedForwarder {
  /** Envelope sender domain the forwarder rewrites to (SRS), e.g. "srs.example.net". */
  srs_domain: string;
  /** Domain of the authserv-id the forwarder stamps, e.g. "example.net" for "mx1.example.net". */
  authserv_id: string;
}

export interface MailRule {
  id: string;
  accounts?: string[];
  match?: Record<string, unknown>;
  actions: Array<string | Record<string, unknown>>;
  continue?: boolean;
  enabled?: boolean;
  [key: string]: unknown;
}

export interface ServiceConfig {
  accountId?: string;
  legacyTokenPath?: boolean;
  allowedAgents?: string[];
  logger?: typeof log;
  bind: string;
  port: number;
  tokenPath: string;
  tokenBrokerSecret: string;
  webhookPath: string;
  webhookUrl: string;
  webhookClientState: string;
  statePath: string;
  runtimeConfigPath: string;
  pipelineWorkspace: string;
  clientId: string;
  clientSecret?: string;
  tenant: string;
  features: M365Feature[];
  initialRefreshToken?: string;
  notifyTarget: string;
  notifyChannel: string;
  /** Path to SQLite mail store DB. Undefined = feature disabled. */
  mailStorePath?: string;
}

export function log(message: string): void {
  console.log(`[m365-webhook] ${message}`);
}

function env(name: string): string {
  return process.env[name]?.trim() || "";
}

function required(value: string, names: string): string {
  if (!value) throw new Error(`Required environment variable is missing: ${names}`);
  return value;
}

function normalizePath(value: string, fallback: string): string {
  const path = value.trim() || fallback;
  return path.startsWith("/") ? path : `/${path}`;
}

export function loadServiceConfig(): { accounts: ServiceConfig[] } {
  const accountsPath = env("M365_ACCOUNTS_PATH");
  if (!accountsPath) return { accounts: [baseConfig()] };
  const { accounts } = JSON.parse(readFileSync(accountsPath, "utf8")) as {
    accounts: Record<string, {
      statePath: string;
      rulesPath: string;
      pipelineWorkspace: string;
      logPath: string;
      mailStorePath?: string;
      brokerSecretEnv: string;
      clientStateEnv: string;
      features: string | string[];
      notify: { channel: string; target: string };
      legacyTokenPath?: boolean;
      allowedAgents: string[];
    }>;
  };
  const paths = new Set<string>();
  const secrets = new Set<string>();
  const files = new Set<string>();
  let legacy = false;
  function accountPath(value: string): string {
    const path = resolve(required(value?.trim(), "account path").replace(/^~\//, homedir() + "/"));
    function canonical(candidate: string): string {
      if (existsSync(candidate)) return realpathSync(candidate);
      return join(canonical(dirname(candidate)), basename(candidate));
    }
    const key = canonical(path);
    const stat = existsSync(path) ? statSync(path) : undefined;
    const inode = stat && `${stat.dev}:${stat.ino}`;
    if (paths.has(key) || (inode && files.has(inode))) throw new Error("Shared Microsoft 365 account path");
    paths.add(key);
    if (inode) files.add(inode);
    return path;
  }
  const entries = Object.entries(accounts);
  if (!entries.length) throw new Error("At least one Microsoft 365 account is required");
  return { accounts: entries.map(([accountId, entry]) => {
    if (!/^[a-z0-9][a-z0-9_-]*$/.test(accountId)) throw new Error("Invalid account ID");
    if (entry.legacyTokenPath && legacy) throw new Error("Multiple legacy token routes");
    legacy ||= entry.legacyTokenPath === true;
    const tokenBrokerSecret = required(env(entry.brokerSecretEnv), "account broker secret");
    const webhookClientState = required(env(entry.clientStateEnv), "account client state");
    for (const secret of [tokenBrokerSecret, webhookClientState]) {
      if (secrets.has(secret)) throw new Error("Shared Microsoft 365 account credential");
      secrets.add(secret);
    }
    if (entry.features === undefined) throw new Error("Account features must be specified");
    if (!Array.isArray(entry.allowedAgents) ||
        entry.allowedAgents.some((agent) => typeof agent !== "string" || !/^[a-z0-9_-]+$/.test(agent))) {
      throw new Error("Invalid account agent allowlist");
    }
    const mailStorePath = entry.mailStorePath || (entry.legacyTokenPath ? env("M365_MAIL_STORE_PATH") : "");
    const logPath = accountPath(entry.logPath);
    const logger = (message: string): void => {
      try {
        appendFileSync(logPath, `[${new Date().toISOString()}] [${accountId}] ${message}\n`, { mode: 0o600 });
      } catch {
        log("Failed to write an account log");
      }
    };
    return baseConfig({
      accountId,
      legacyTokenPath: entry.legacyTokenPath,
      allowedAgents: entry.allowedAgents,
      logger,
      tokenBrokerSecret,
      webhookClientState,
      statePath: accountPath(entry.statePath),
      runtimeConfigPath: accountPath(entry.rulesPath),
      pipelineWorkspace: accountPath(entry.pipelineWorkspace),
      features: parseM365Features(entry.features),
      notifyTarget: required(entry.notify?.target, "account notify target"),
      notifyChannel: required(entry.notify?.channel, "account notify channel"),
      mailStorePath: mailStorePath ? accountPath(mailStorePath) : undefined,
    });
  }) };
}

function baseConfig(account?: Partial<ServiceConfig>): ServiceConfig {
  const home = homedir();
  const statePath = env("M365_WEBHOOK_STATE_PATH")
    || join(home, ".openclaw/state/m365-webhook.json");
  const runtimeConfigPath = env("M365_WEBHOOK_CONFIG_PATH")
    || join(home, ".openclaw/services/m365-webhook-config.json");
  const portValue = env("M365_TOKEN_BROKER_PORT")
    || env("M365_WEBHOOK_PORT")
    || "18790";
  const port = Number.parseInt(portValue, 10);
  if (!Number.isInteger(port) || port < 1 || port > 65535) {
    throw new Error(`Invalid webhook port: ${portValue}`);
  }

  return {
    bind: env("M365_TOKEN_BROKER_BIND") || "127.0.0.1",
    port,
    tokenPath: normalizePath(
      env("M365_TOKEN_BROKER_PATH"),
      "/token",
    ),
    tokenBrokerSecret: account?.tokenBrokerSecret ?? required(
      env("M365_TOKEN_BROKER_SECRET"),
      "M365_TOKEN_BROKER_SECRET",
    ),
    webhookPath: normalizePath(
      env("M365_WEBHOOK_PATH"),
      "/m365/webhook",
    ),
    webhookUrl: required(
      env("M365_WEBHOOK_URL"),
      "M365_WEBHOOK_URL",
    ),
    webhookClientState: account?.webhookClientState ?? required(
      env("M365_WEBHOOK_CLIENT_STATE"),
      "M365_WEBHOOK_CLIENT_STATE",
    ),
    statePath,
    runtimeConfigPath,
    pipelineWorkspace: env("M365_PIPELINE_WORKSPACE")
      || join(home, ".openclaw/services/mail-runtime"),
    clientId: required(
      env("M365_CLIENT_ID"),
      "M365_CLIENT_ID",
    ),
    clientSecret: env("M365_CLIENT_SECRET") || undefined,
    tenant: env("M365_TENANT") || "consumers",
    features: parseM365Features(
      env("M365_FEATURES") || undefined,
    ),
    initialRefreshToken: account ? undefined : env("M365_REFRESH_TOKEN") || undefined,
    notifyTarget: account?.notifyTarget ?? required(process.env.NOTIFY_TARGET?.trim() || "", "NOTIFY_TARGET"),
    notifyChannel: process.env.NOTIFY_CHANNEL?.trim() || "discord",
    mailStorePath: env("M365_MAIL_STORE_PATH") || undefined,
    ...account,
  };
}

export function loadRuntimeConfig(path: string, logger = log): RuntimeConfig {
  if (!existsSync(path)) {
    logger(`No runtime config found at ${path}; starting with empty rules`);
    return {};
  }
  try {
    return JSON.parse(readFileSync(path, "utf8")) as RuntimeConfig;
  } catch (error) {
    logger(`Invalid runtime config at ${path}: ${error}; using empty rules`);
    return {};
  }
}

export function buildPipelineRules(config: RuntimeConfig): MailRule[] {
  return [...(config.mail_rules ?? [])];
}

export function buildTrustedForwarders(config: RuntimeConfig, logger = log): TrustedForwarder[] {
  const forwarders: TrustedForwarder[] = [];
  for (const entry of config.trusted_forwarders ?? []) {
    const srs = typeof entry?.srs_domain === "string" ? entry.srs_domain.trim().toLowerCase() : "";
    const authserv = typeof entry?.authserv_id === "string" ? entry.authserv_id.trim().toLowerCase() : "";
    if (!srs || !authserv) {
      logger(`ignoring trusted_forwarders entry without srs_domain and authserv_id: ${JSON.stringify(entry)}`);
      continue;
    }
    forwarders.push({ srs_domain: srs, authserv_id: authserv });
  }
  return forwarders;
}
