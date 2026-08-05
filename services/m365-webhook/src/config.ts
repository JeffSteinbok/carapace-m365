import { existsSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
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
  [key: string]: unknown;
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
}

export function log(message: string): void {
  console.log(`[m365-webhook] ${message}`);
}

function env(m365Name: string, outlookName?: string): string {
  return process.env[m365Name]?.trim()
    || (outlookName ? process.env[outlookName]?.trim() : "")
    || "";
}

function required(value: string, names: string): string {
  if (!value) throw new Error(`Required environment variable is missing: ${names}`);
  return value;
}

function normalizePath(value: string, fallback: string): string {
  const path = value.trim() || fallback;
  return path.startsWith("/") ? path : `/${path}`;
}

export function loadServiceConfig(): ServiceConfig {
  const home = homedir();
  const newStatePath = join(home, ".openclaw/state/m365-webhook.json");
  const oldStatePath = join(home, ".openclaw/state/outlook-webhook.json");
  const configuredState = env("M365_WEBHOOK_STATE_PATH", "OUTLOOK_WEBHOOK_STATE_PATH");
  const statePath = configuredState
    || (existsSync(newStatePath) || !existsSync(oldStatePath) ? newStatePath : oldStatePath);

  const newRuntimePath = join(home, ".openclaw/services/m365-webhook-config.json");
  const oldRuntimePath = join(home, ".openclaw/services/outlook-webhook-config.json");
  const configuredRuntime = env("M365_WEBHOOK_CONFIG_PATH", "OUTLOOK_WEBHOOK_CONFIG_PATH");
  const runtimeConfigPath = configuredRuntime
    || (existsSync(newRuntimePath) || !existsSync(oldRuntimePath) ? newRuntimePath : oldRuntimePath);

  const portValue = env("M365_TOKEN_BROKER_PORT", "OUTLOOK_TOKEN_BROKER_PORT")
    || env("M365_WEBHOOK_PORT", "OUTLOOK_WEBHOOK_PORT")
    || "18790";
  const port = Number.parseInt(portValue, 10);
  if (!Number.isInteger(port) || port < 1 || port > 65535) {
    throw new Error(`Invalid webhook port: ${portValue}`);
  }

  return {
    bind: env("M365_TOKEN_BROKER_BIND", "OUTLOOK_TOKEN_BROKER_BIND") || "127.0.0.1",
    port,
    tokenPath: normalizePath(
      env("M365_TOKEN_BROKER_PATH", "OUTLOOK_TOKEN_BROKER_PATH"),
      "/token",
    ),
    tokenBrokerSecret: required(
      env("M365_TOKEN_BROKER_SECRET", "OUTLOOK_TOKEN_BROKER_SECRET"),
      "M365_TOKEN_BROKER_SECRET",
    ),
    webhookPath: normalizePath(
      env("M365_WEBHOOK_PATH", "OUTLOOK_WEBHOOK_PATH"),
      "/outlook/webhook",
    ),
    webhookUrl: required(
      env("M365_WEBHOOK_URL", "OUTLOOK_WEBHOOK_URL"),
      "M365_WEBHOOK_URL or OUTLOOK_WEBHOOK_URL",
    ),
    webhookClientState: required(
      env("M365_WEBHOOK_CLIENT_STATE", "OUTLOOK_WEBHOOK_CLIENT_STATE"),
      "M365_WEBHOOK_CLIENT_STATE or OUTLOOK_WEBHOOK_CLIENT_STATE",
    ),
    statePath,
    runtimeConfigPath,
    pipelineWorkspace: env("M365_PIPELINE_WORKSPACE")
      || join(home, ".openclaw/services/mail-runtime"),
    clientId: required(
      env("M365_CLIENT_ID", "OUTLOOK_CLIENT_ID"),
      "M365_CLIENT_ID or OUTLOOK_CLIENT_ID",
    ),
    clientSecret: env("M365_CLIENT_SECRET", "OUTLOOK_CLIENT_SECRET") || undefined,
    tenant: env("M365_TENANT", "OUTLOOK_TENANT") || "consumers",
    features: parseM365Features(
      env("M365_FEATURES", "OUTLOOK_FEATURES") || undefined,
    ),
    initialRefreshToken: env("M365_REFRESH_TOKEN", "OUTLOOK_REFRESH_TOKEN") || undefined,
    notifyTarget: required(process.env.NOTIFY_TARGET?.trim() || "", "NOTIFY_TARGET"),
    notifyChannel: process.env.NOTIFY_CHANNEL?.trim() || "discord",
  };
}

export function loadRuntimeConfig(path: string): RuntimeConfig {
  if (!existsSync(path)) {
    log(`No runtime config found at ${path}; starting with empty rules`);
    return {};
  }
  try {
    return JSON.parse(readFileSync(path, "utf8")) as RuntimeConfig;
  } catch (error) {
    log(`Invalid runtime config at ${path}: ${error}; using empty rules`);
    return {};
  }
}

export function buildPipelineRules(config: RuntimeConfig): MailRule[] {
  return [...(config.mail_rules ?? [])];
}
