# Microsoft 365 webhook and token broker

This standalone process replaces
`openclaw-hub/services/outlook-webhook`. It:

- receives Microsoft Graph inbox change notifications;
- runs messages through `carapace-mail-runtime`;
- owns and rotates the Microsoft OAuth refresh token;
- exposes authenticated `POST /token` to the root plugin;
- persists refresh-token and subscription state atomically;
- reuses/renews subscriptions across restarts.

Normal shutdown deliberately does **not** delete a valid Graph subscription.

## Build and run

From the `carapace-outlook` repository root:

```bash
npm install
npm run build
node services/m365-webhook/dist/index.js
```

## Environment

| Variable | Default / fallback | Purpose |
|---|---|---|
| `M365_CLIENT_ID` | `OUTLOOK_CLIENT_ID` | App registration client ID |
| `M365_CLIENT_SECRET` | `OUTLOOK_CLIENT_SECRET` | Optional confidential-client secret |
| `M365_TENANT` | `OUTLOOK_TENANT`, then `consumers` | OAuth tenant such as `common` or a tenant ID |
| `M365_REFRESH_TOKEN` | `OUTLOOK_REFRESH_TOKEN` | Initial token, used when state has no token |
| `M365_TOKEN_BROKER_SECRET` | `OUTLOOK_TOKEN_BROKER_SECRET` | Required bearer secret for `/token` |
| `M365_TOKEN_BROKER_BIND` | `OUTLOOK_TOKEN_BROKER_BIND`, then `127.0.0.1` | HTTP bind address |
| `M365_TOKEN_BROKER_PORT` | `OUTLOOK_TOKEN_BROKER_PORT`, `M365_WEBHOOK_PORT`, `OUTLOOK_WEBHOOK_PORT`, then `18790` | Shared broker/webhook port |
| `M365_TOKEN_BROKER_PATH` | `OUTLOOK_TOKEN_BROKER_PATH`, then `/token` | Token endpoint path |
| `M365_WEBHOOK_PORT` | `OUTLOOK_WEBHOOK_PORT`, then `18790` | Backward-compatible shared port setting |
| `M365_WEBHOOK_PATH` | `OUTLOOK_WEBHOOK_PATH`, then `/outlook/webhook` | Local Graph webhook path |
| `M365_WEBHOOK_URL` | `OUTLOOK_WEBHOOK_URL` | Required public HTTPS notification URL |
| `M365_WEBHOOK_CLIENT_STATE` | `OUTLOOK_WEBHOOK_CLIENT_STATE` | Required Graph notification validation secret |
| `M365_WEBHOOK_STATE_PATH` | `OUTLOOK_WEBHOOK_STATE_PATH`, then `~/.openclaw/state/m365-webhook.json` | Token/subscription state |
| `M365_WEBHOOK_CONFIG_PATH` | `OUTLOOK_WEBHOOK_CONFIG_PATH`, then `~/.openclaw/services/m365-webhook-config.json` | Mail-runtime rules |
| `M365_PIPELINE_WORKSPACE` | `~/.openclaw/services/mail-runtime` | Mail-runtime workspace |
| `NOTIFY_TARGET` | none | Required OpenClaw notification target |
| `NOTIFY_CHANNEL` | `discord` | OpenClaw notification channel |

If the new default state/config file is absent and the old Outlook default
exists, the service reuses the old path automatically.

Use long independently generated values for `M365_TOKEN_BROKER_SECRET` and
`M365_WEBHOOK_CLIENT_STATE`. Neither is logged.

## Token endpoint

The plugin calls:

```http
POST /token
Authorization: Bearer <M365_TOKEN_BROKER_SECRET>
Content-Type: application/json

{"scopes":["Mail.Read"]}
```

Successful responses contain an access token and expiry. The endpoint accepts
only authenticated requests and delegated scope-name syntax. Loopback HTTP is
the default; use HTTPS for any non-loopback broker URL.

All refreshes are serialized across scope sets. If Microsoft returns a rotated
`refresh_token`, it is written atomically before it is used for later refreshes.

## State

State JSON contains:

```json
{
  "refreshToken": "...",
  "subscriptionId": "...",
  "expirationDateTime": "...",
  "notificationUrl": "https://example.test/outlook/webhook",
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
  "path": "/outlook/webhook",
  "target": "http://127.0.0.1:18790",
  "auth": "none"
}
```

The root plugin should call `http://127.0.0.1:18790/token` directly.

## Migration from openclaw-hub

Old executable:

```text
/home/openclaw/git/openclaw-hub/services/outlook-webhook/dist/index.js
```

New executable:

```text
/home/openclaw/git/carapace-outlook/services/m365-webhook/dist/index.js
```

No change to `openclaw-hub` is required in this repository task. Update the
deployed systemd unit separately after this repository is installed and built.

Example user unit:

```ini
[Unit]
Description=OpenClaw Microsoft 365 Webhook and Token Broker
After=network.target

[Service]
ExecStart=/usr/bin/node /home/openclaw/git/carapace-outlook/services/m365-webhook/dist/index.js
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

If retaining the old unit name temporarily, only its `ExecStart` needs to point
to the new path. Do not run both services with the same refresh token; competing
rotation can invalidate one process.
