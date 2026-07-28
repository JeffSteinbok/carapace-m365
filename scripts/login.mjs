#!/usr/bin/env node
/**
 * carapace-outlook — public-client PKCE login.
 *
 * Mints an OAuth refresh token for a personal Microsoft account WITHOUT a
 * client secret, using the Authorization Code + PKCE flow against a loopback
 * redirect. Requires only the app's client ID (a public-client registration
 * with "Allow public client flows" enabled and http://localhost registered as
 * a Mobile/desktop redirect URI).
 *
 * Usage:
 *   OUTLOOK_CLIENT_ID=<app-id> node scripts/login.mjs
 *   node scripts/login.mjs --client-id <app-id> [--port 53682]
 *
 * On success it prints the refresh token; set it as OUTLOOK_REFRESH_TOKEN.
 * No external dependencies — Node stdlib only.
 */

import http from "node:http";
import crypto from "node:crypto";
import { spawn } from "node:child_process";
import { URL } from "node:url";

const TENANT = "consumers"; // personal Microsoft accounts
const AUTH_URL = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/authorize`;
const TOKEN_URL = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`;
const SCOPE = "Calendars.ReadWrite Mail.ReadWrite Mail.Send Tasks.ReadWrite offline_access openid";

function arg(name, fallback) {
  const i = process.argv.indexOf(`--${name}`);
  if (i !== -1 && process.argv[i + 1]) return process.argv[i + 1];
  return fallback;
}

const clientId = arg("client-id", process.env.OUTLOOK_CLIENT_ID);
const port = Number(arg("port", "53682"));
const redirectUri = `http://localhost:${port}`;

if (!clientId) {
  console.error("Error: provide the app client ID via --client-id or OUTLOOK_CLIENT_ID.");
  process.exit(1);
}

function base64url(buf) {
  return buf.toString("base64").replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

const verifier = base64url(crypto.randomBytes(32));
const challenge = base64url(crypto.createHash("sha256").update(verifier).digest());
const state = base64url(crypto.randomBytes(16));

function openBrowser(url) {
  const cmd = process.platform === "darwin" ? "open"
    : process.platform === "win32" ? "cmd" : "xdg-open";
  const args = process.platform === "win32" ? ["/c", "start", "", url] : [url];
  try { spawn(cmd, args, { stdio: "ignore", detached: true }).unref(); } catch { /* ignore */ }
}

async function exchange(code) {
  const body = new URLSearchParams({
    client_id: clientId,
    grant_type: "authorization_code",
    code,
    redirect_uri: redirectUri,
    code_verifier: verifier,
    scope: SCOPE,
  }).toString();
  const res = await fetch(TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  return res.json();
}

const authUrl = `${AUTH_URL}?` + new URLSearchParams({
  client_id: clientId,
  response_type: "code",
  redirect_uri: redirectUri,
  response_mode: "query",
  scope: SCOPE,
  state,
  code_challenge: challenge,
  code_challenge_method: "S256",
}).toString();

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, redirectUri);
  if (url.pathname !== "/") { res.writeHead(404).end(); return; }

  const err = url.searchParams.get("error");
  if (err) {
    res.writeHead(400, { "Content-Type": "text/html" });
    res.end(`<h1>Login failed</h1><p>${err}: ${url.searchParams.get("error_description") ?? ""}</p>`);
    console.error(`Auth error: ${err} — ${url.searchParams.get("error_description") ?? ""}`);
    server.close(); process.exit(1);
  }

  const code = url.searchParams.get("code");
  if (!code) { res.writeHead(400).end("Missing code"); return; }
  if (url.searchParams.get("state") !== state) {
    res.writeHead(400).end("State mismatch"); server.close(); process.exit(1);
  }

  const tokens = await exchange(code);
  if (tokens.error) {
    res.writeHead(400, { "Content-Type": "text/html" });
    res.end(`<h1>Token exchange failed</h1><pre>${tokens.error}: ${tokens.error_description ?? ""}</pre>`);
    console.error(`Token error: ${tokens.error} — ${tokens.error_description ?? ""}`);
    server.close(); process.exit(1);
  }

  res.writeHead(200, { "Content-Type": "text/html" });
  res.end("<h1>Signed in ✅</h1><p>You can close this tab and return to the terminal.</p>");

  console.log("\n✅ Refresh token acquired. Set this in your environment:\n");
  console.log(`OUTLOOK_REFRESH_TOKEN=${tokens.refresh_token}\n`);
  server.close();
  process.exit(0);
});

server.listen(port, () => {
  console.log(`Listening on ${redirectUri} for the OAuth redirect...`);
  console.log("Opening your browser to sign in. If it doesn't open, paste this URL:\n");
  console.log(authUrl + "\n");
  openBrowser(authUrl);
});
