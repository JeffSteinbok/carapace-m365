#!/usr/bin/env node

import http from "node:http";
import crypto from "node:crypto";
import { spawn } from "node:child_process";
import {
  DEFAULT_M365_FEATURES,
  M365_FEATURE_POLICY,
  M365_FEATURES,
  deriveFeatureGraphScopes,
  inferM365FeaturesFromScopes,
  normalizeGraphScopeNames,
  parseM365Features,
} from "../packages/m365-graph-auth/src/feature-policy.js";

const DEFAULT_CLIENT_ID = "0c3df71b-4dc2-49a7-b6e7-e5c3c48bf501";

function arg(name, fallback) {
  const index = process.argv.indexOf(`--${name}`);
  return index !== -1 && process.argv[index + 1]
    ? process.argv[index + 1]
    : fallback;
}

function hasArg(name) {
  return process.argv.includes(`--${name}`);
}

function printHelp() {
  console.log(`Microsoft 365 PKCE login

Usage:
  npm run login
  npm run login -- --features mail-read,mail-send
  npm run login -- --scopes "Mail.Read Files.Read"

Options:
  --features <list>  Comma-separated feature names
  --scopes <list>    Explicit comma- or space-separated delegated scopes
  --client-id <id>   Azure application client ID
  --tenant <tenant>  consumers (default), common, organizations, or tenant ID
  --port <port>      Loopback callback port (default 53682)
  --list-features    Print available feature names
  --help             Show this help

The default enables the core Microsoft 365 consent:
  ${DEFAULT_M365_FEATURES.join(", ")}

offline_access and openid are added only to this interactive authorization
request so Microsoft can issue a refresh token.`);
}

if (hasArg("help")) {
  printHelp();
  process.exit(0);
}
if (hasArg("list-features")) {
  for (const feature of M365_FEATURES) {
    console.log(
      `${feature.padEnd(16)} ${M365_FEATURE_POLICY[feature].scopes.join(" ")}`,
    );
  }
  process.exit(0);
}

const clientId = arg(
  "client-id",
  process.env.M365_CLIENT_ID || DEFAULT_CLIENT_ID,
);
const tenant = arg(
  "tenant",
  process.env.M365_TENANT || "consumers",
);
const port = Number(arg("port", "53682"));
if (!Number.isInteger(port) || port < 1 || port > 65535) {
  console.error("Error: --port must be an integer between 1 and 65535.");
  process.exit(1);
}

const explicitScopes = arg("scopes", "");
let requestedFeatures;
try {
  requestedFeatures = parseM365Features(
    arg("features", DEFAULT_M365_FEATURES.join(",")),
  );
} catch (error) {
  console.error(`Error: ${error.message}`);
  console.error("Run with --list-features to see valid names.");
  process.exit(1);
}

const delegatedScopes = explicitScopes
  ? normalizeGraphScopeNames(explicitScopes.split(/[\s,]+/))
  : deriveFeatureGraphScopes(requestedFeatures);
const inferred = explicitScopes
  ? inferM365FeaturesFromScopes(delegatedScopes)
  : { features: requestedFeatures, unknownScopes: [] };
const scopes = normalizeGraphScopeNames([
  ...delegatedScopes,
  "offline_access",
  "openid",
]);
if (delegatedScopes.length === 0) {
  console.error("Error: select at least one feature or delegated scope.");
  process.exit(1);
}
if (inferred.unknownScopes.length) {
  console.warn(
    `Warning: these scopes are not represented by supported Microsoft 365 features: ${inferred.unknownScopes.join(", ")}`,
  );
  console.warn(
    "The broker will reject them until its shared feature policy and configured allowlist are updated manually.",
  );
}

const redirectUri = `http://localhost:${port}`;
const authUrlBase =
  `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/authorize`;
const tokenUrl =
  `https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`;

function base64url(buffer) {
  return buffer.toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function openBrowser(url) {
  const command = process.platform === "darwin"
    ? "open"
    : process.platform === "win32"
      ? "cmd"
      : "xdg-open";
  const args = process.platform === "win32"
    ? ["/c", "start", "", url]
    : [url];
  try {
    spawn(command, args, { stdio: "ignore", detached: true }).unref();
  } catch {
    // The authorization URL is also printed for manual use.
  }
}

const verifier = base64url(crypto.randomBytes(32));
const challenge = base64url(crypto.createHash("sha256").update(verifier).digest());
const state = base64url(crypto.randomBytes(16));
const scope = scopes.join(" ");

async function exchange(code) {
  const response = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: clientId,
      grant_type: "authorization_code",
      code,
      redirect_uri: redirectUri,
      code_verifier: verifier,
      scope,
    }),
  });
  const body = await response.json();
  if (!response.ok && !body.error) {
    body.error = `HTTP ${response.status}`;
  }
  return body;
}

const authorizationUrl = `${authUrlBase}?${new URLSearchParams({
  client_id: clientId,
  response_type: "code",
  redirect_uri: redirectUri,
  response_mode: "query",
  scope,
  state,
  code_challenge: challenge,
  code_challenge_method: "S256",
})}`;

console.log(clientId === DEFAULT_CLIENT_ID
  ? `Using the published app registration (${clientId}).`
  : `Using client ID ${clientId}.`);
console.log(`Tenant: ${tenant}`);
console.log(`Delegated Graph scopes: ${delegatedScopes.join(" ")}`);

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, redirectUri);
  if (url.pathname !== "/") {
    res.writeHead(404).end();
    return;
  }
  const authError = url.searchParams.get("error");
  if (authError) {
    const description = url.searchParams.get("error_description") ?? "";
    res.writeHead(400, { "Content-Type": "text/html; charset=utf-8" });
    res.end(`<h1>Login failed</h1><p>${escapeHtml(authError)}: ${escapeHtml(description)}</p>`);
    console.error(`Authorization failed: ${authError} — ${description}`);
    server.close();
    process.exitCode = 1;
    return;
  }
  const code = url.searchParams.get("code");
  if (!code) {
    res.writeHead(400).end("Missing code");
    return;
  }
  if (url.searchParams.get("state") !== state) {
    res.writeHead(400).end("State mismatch");
    server.close();
    process.exitCode = 1;
    return;
  }

  try {
    const tokens = await exchange(code);
    if (tokens.error || !tokens.refresh_token) {
      const description = tokens.error_description ?? "No refresh_token was returned.";
      res.writeHead(400, { "Content-Type": "text/html; charset=utf-8" });
      res.end(`<h1>Token exchange failed</h1><p>${escapeHtml(tokens.error ?? "error")}: ${escapeHtml(description)}</p>`);
      console.error(`Token exchange failed: ${tokens.error ?? "error"} — ${description}`);
      process.exitCode = 1;
      return;
    }
    res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
    res.end("<h1>Signed in</h1><p>You can close this tab and return to the terminal.</p>");
    console.log("\nRefresh token acquired. Store it in the webhook service state or environment:\n");
    console.log(`M365_REFRESH_TOKEN=${tokens.refresh_token}`);
    console.log(`M365_REFRESH_TOKEN=${tokens.refresh_token}`);
    console.log(`M365_FEATURES=${inferred.features.join(",")}`);
    console.log("\nUse the same features in the plugin config:\n");
    console.log(JSON.stringify({ features: inferred.features }, null, 2));
    if (inferred.unknownScopes.length) {
      console.warn(
        `\nWarning: M365_FEATURES covers only recognized scopes; not covered: ${inferred.unknownScopes.join(", ")}`,
      );
      console.warn(
        "Do not treat this as a complete broker allowlist. Update the shared feature policy and set the matching broker/plugin features manually.",
      );
    }
    console.log();
  } catch (error) {
    res.writeHead(500, { "Content-Type": "text/html; charset=utf-8" });
    res.end("<h1>Token exchange failed</h1>");
    console.error(`Token exchange failed: ${error}`);
    process.exitCode = 1;
  } finally {
    server.close();
  }
});

server.listen(port, "127.0.0.1", () => {
  console.log(`Listening on ${redirectUri} for the OAuth redirect.`);
  console.log("Opening your browser. If it does not open, paste this URL:\n");
  console.log(`${authorizationUrl}\n`);
  openBrowser(authorizationUrl);
});
