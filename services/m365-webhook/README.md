# Microsoft 365 webhook and token broker

This standalone process owns Microsoft Graph webhook delivery and token
brokering. It:

- receives Microsoft Graph inbox change notifications;
- runs messages through `carapace-mail-runtime`;
- owns and rotates the Microsoft OAuth refresh token;
- exposes authenticated `POST /token` to the root plugin;
- persists refresh-token and subscription state atomically;
- reuses/renews subscriptions across restarts.

Normal shutdown deliberately does **not** delete a valid Graph subscription.

## Build and run

From the `carapace-m365` repository root:

```bash
npm install
npm run build
node services/m365-webhook/dist/index.js
```

## Environment

| Variable | Default / fallback | Purpose |
|---|---|---|
| `M365_CLIENT_ID` | — | App registration client ID |
| `M365_CLIENT_SECRET` | — | Optional confidential-client secret |
| `M365_TENANT` | `consumers` | OAuth tenant such as `common` or a tenant ID |
| `M365_REFRESH_TOKEN` | — | Initial token, used when state has no token |
| `M365_FEATURES` | `calendar-write,mail-write,mail-send,tasks-write` | Authoritative comma-separated broker scope allowlist |
| `M365_TOKEN_BROKER_SECRET` | — | Required bearer secret for `/token` |
| `M365_TOKEN_BROKER_BIND` | `127.0.0.1` | HTTP bind address |
| `M365_TOKEN_BROKER_PORT` | `18790` | Shared broker/webhook port |
| `M365_TOKEN_BROKER_PATH` | `/token` | Token endpoint path |
| `M365_WEBHOOK_PORT` | `18790` | Shared broker/webhook port fallback |
| `M365_WEBHOOK_PATH` | `/m365/webhook` | Local Graph webhook path |
| `M365_WEBHOOK_URL` | — | Required public HTTPS notification URL |
| `M365_WEBHOOK_CLIENT_STATE` | — | Required Graph notification validation secret |
| `M365_WEBHOOK_STATE_PATH` | `~/.openclaw/state/m365-webhook.json` | Token/subscription state |
| `M365_WEBHOOK_CONFIG_PATH` | `~/.openclaw/services/m365-webhook-config.json` | Mail-runtime rules |
| `M365_PIPELINE_WORKSPACE` | `~/.openclaw/services/mail-runtime` | Mail-runtime workspace |
| `NOTIFY_TARGET` | none | Required OpenClaw notification target |
| `NOTIFY_CHANNEL` | `discord` | OpenClaw notification channel |

Use long independently generated values for `M365_TOKEN_BROKER_SECRET` and
`M365_WEBHOOK_CLIENT_STATE`. Neither is logged.

Set `M365_FEATURES` to the exact same normalized feature list used by the
plugin. Supported values are `mail-read`, `mail-write`, `mail-send`,
`calendar-read`, `calendar-write`, `tasks-read`, `tasks-write`,
`onedrive-read`, and `onedrive-write`. Write features imply their matching read
feature. `mail-send` is independent. OneDrive scopes are denied unless
explicitly enabled.

## Token endpoint

The plugin calls:

```http
POST /token
Authorization: Bearer <M365_TOKEN_BROKER_SECRET>
Content-Type: application/json

{"scopes":["Mail.Read"]}
```

Successful responses contain an access token and expiry. The endpoint accepts
only authenticated requests and delegated scope-name syntax, then checks every
scope case-insensitively against its own `M365_FEATURES` allowlist. Disallowed
or unknown scopes return JSON HTTP 403 without contacting Microsoft. Loopback
HTTP is the default; use HTTPS for any non-loopback broker URL.

All refreshes are serialized across scope sets. If Microsoft returns a rotated
`refresh_token`, it is written atomically before it is used for later refreshes.

This is one of three enforcement layers: plugin feature configuration controls
tool registration, this service independently controls broker scopes, and
Microsoft consent controls what the refresh token can actually obtain. None of
the layers expands either of the others.

## State

State JSON contains:

```json
{
  "refreshToken": "...",
  "subscriptionId": "...",
  "expirationDateTime": "...",
  "notificationUrl": "https://example.test/m365/webhook",
  "clientState": "..."
}
```

Writes use a same-directory temporary file and atomic rename. The service
requests mode `0600` for files and `0700` for directories where supported.
The service never logs `clientState`. A changed notification URL or client
state causes the old subscription to be removed best-effort and replaced.

Once state contains `refreshToken`, it is authoritative over environment
variables. For incremental consent, stop the service and replace that field
while preserving subscription metadata, then restart.

## Proxy routing

Keep the public Graph webhook path unauthenticated at the proxy because Graph
validates through `validationToken` and `clientState`. Do **not** expose `/token`
through that unauthenticated route.

Example routing:

```json
{
  "path": "/m365/webhook",
  "target": "http://127.0.0.1:18790",
  "auth": "none"
}
```

The root plugin should call `http://127.0.0.1:18790/token` directly.

## Migration from openclaw-hub

Old executable:

```text
/home/openclaw/git/carapace-m365/services/m365-webhook/dist/index.js
```

New executable:

```text
/home/openclaw/git/carapace-m365/services/m365-webhook/dist/index.js
```

No change to `openclaw-hub` is required in this repository task. Update the
deployed systemd unit separately after this repository is installed and built.

Example user unit:

```ini
[Unit]
Description=OpenClaw Microsoft 365 Webhook and Token Broker
After=network.target

[Service]
ExecStart=/usr/bin/node /home/openclaw/git/carapace-m365/services/m365-webhook/dist/index.js
Restart=on-failure
RestartSec=10
EnvironmentFile=%h/.openclaw/.env

[Install]
WantedBy=default.target
```

Then, outside this repository change:

```bash
systemctl --user daemon-reload
systemctl --user enable --now m365-webhook
systemctl --user status m365-webhook
```

Do not run multiple services with the same refresh token; competing rotation can
invalidate one process.
