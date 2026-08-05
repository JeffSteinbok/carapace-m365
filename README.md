# 📧 📅 ✅ 📁 Carapace Microsoft 365

Microsoft 365 tools for Outlook mail, calendars, Microsoft To Do, and OneDrive
through Microsoft Graph.

The package is named `carapace-m365`. Outlook and OneDrive tool names remain
brand-specific within the Microsoft 365 plugin.

## Architecture

The recommended deployment uses the standalone token broker included in this
repository:

```text
Carapace/OpenClaw M365 plugin ── authenticated /token ──► m365-webhook
                                                     ├─ owns refresh token
                                                     ├─ persists rotation
                                                     ├─ renews Graph subscription
                                                     └─ dispatches inbox notifications
```

The plugin caches scoped access tokens locally. When a broker is configured it
is authoritative: the plugin never falls back to refreshing independently after
a broker failure. It may continue only with a cached access token that has not
actually expired.

Direct refresh-token mode remains available for migration and simple
single-process use.

## Quick start

### 1. Consent only to the features you want

The default grants the core capabilities:
`calendar-write,mail-write,mail-send,tasks-write`. It does **not** add OneDrive,
and OneDrive tools are not registered unless `onedrive-read` or
`onedrive-write` is explicitly enabled.

```powershell
npm run login
```

Use `--features` for least-privilege or incremental consent:

```powershell
# Read-only OneDrive
npm run login -- --features onedrive-read

# Full OneDrive (read and write)
npm run login -- --features onedrive-write

# Existing mail capabilities only
npm run login -- --features mail-write,mail-send

# Calendar only
npm run login -- --features calendar-write

# Existing Outlook capabilities plus full OneDrive
npm run login -- --features calendar-write,mail-write,mail-send,tasks-write,onedrive-write

# Exact delegated scopes are also supported
npm run login -- --scopes "Mail.Read Files.Read"
```

Run `npm run login -- --list-features` for every feature name. The login command
adds `offline_access` and `openid` because interactive authorization needs them
to issue a refresh token; Graph access-token requests use only the scopes needed
for each operation. After consent it prints `M365_FEATURES` and a matching
plugin `features` snippet; use that same normalized list in both places.

### 2. Run the token broker/webhook service

Build everything:

```powershell
npm install
npm run build
```

Set at least:

```text
M365_CLIENT_ID=your-app-client-id
M365_REFRESH_TOKEN=the-token-from-login
M365_FEATURES=calendar-write,mail-write,mail-send,tasks-write
M365_TOKEN_BROKER_SECRET=a-long-random-secret
M365_WEBHOOK_URL=https://your-public-host/m365/webhook
M365_WEBHOOK_CLIENT_STATE=another-long-random-secret
NOTIFY_TARGET=your-openclaw-notification-target
```

Then start:

```powershell
node services/m365-webhook/dist/index.js
```

The service binds to `127.0.0.1:18790` by default, exposes authenticated
`POST /token`, and accepts Graph notifications at `/m365/webhook`. See
[services/m365-webhook/README.md](services/m365-webhook/README.md) for proxy,
state, deployment, and systemd details.

### 3. Configure the plugin

```json
{
  "plugins": {
    "entries": {
      "m365": {
        "enabled": true,
        "config": {
          "features": ["calendar-write", "mail-write", "mail-send", "tasks-write"],
          "tokenBrokerUrl": "http://127.0.0.1:18790/token",
          "tokenBrokerSecret": "${M365_TOKEN_BROKER_SECRET}"
        }
      }
    }
  }
}
```

If only `tokenBrokerSecret` is configured, the URL defaults to
`http://127.0.0.1:18790/token`. Non-loopback broker URLs must use HTTPS.

### Direct mode

For migration or a single process without the webhook broker:

```json
{
  "plugins": {
    "entries": {
      "m365": {
        "enabled": true,
        "config": {
          "features": ["calendar-write", "mail-write", "mail-send", "tasks-write"],
          "clientId": "${M365_CLIENT_ID}",
          "refreshToken": "${M365_REFRESH_TOKEN}"
        }
      }
    }
  }
}
```

`clientSecret` is optional. Public PKCE clients do not have one.

Direct mode is strictly a single-process fallback. It persists the authoritative
refresh token at `~/.openclaw/state/m365-direct-token.json` by default, using an
atomic replacement and restrictive permissions where supported. Override the
path with `directTokenStatePath`, `M365_DIRECT_TOKEN_STATE_PATH`, or
Once present, the state token takes precedence over configuration and
environment values. Multiple plugin
processes must use broker mode so only one process owns refresh-token rotation.

## Least-privilege permissions

Choose the narrowest delegated permission that supports the desired operation.

| Feature | Desired capability | Minimum Microsoft Graph delegated permission |
|---|---|---|
| `mail-read` | List, search, read mail; download attachments; inbox webhook | `Mail.Read` |
| `mail-write` | Move or flag mail; also enables mail reads | `Mail.ReadWrite` |
| `mail-send` | Send, reply, or forward mail | `Mail.Send` |
| `calendar-read` | Read calendars/events | `Calendars.Read` |
| `calendar-write` | Create, update, delete events or meetings; also enables reads | `Calendars.ReadWrite` |
| `tasks-read` | Read Microsoft To Do lists/tasks | `Tasks.Read` |
| `tasks-write` | Create, update, complete, or delete tasks; also enables reads | `Tasks.ReadWrite` |
| `onedrive-read` | List, search, inspect, or download OneDrive content | `Files.Read` |
| `onedrive-write` | Upload, create, move, rename, or delete OneDrive content; also enables reads | `Files.ReadWrite` |
| Login-only | Receive a reusable refresh token during login | `offline_access` |

`*.ReadWrite` permissions include the corresponding reads, so selecting both
read and write for the same feature is unnecessary.

Permissions are enforced in three independent layers:

1. the plugin registers only tools enabled by its configured `features`;
2. the broker rejects scopes outside its own `M365_FEATURES` allowlist before
   contacting Microsoft;
3. Microsoft issues tokens only for delegated scopes actually consented by the
   signed-in account.

Configure the same feature list in the plugin and broker. The broker does not
trust the plugin's list, and neither configuration can expand Microsoft consent.
Unknown features and unrecognized broker scopes fail closed.

> **Azure permission configuration is not consent.** Adding a delegated
> permission in Azure does not grant the account access by itself. Re-run
> `npm run login` with the new feature/scopes, complete the consent prompt, and
> replace the refresh token owned by the broker. Stop the service and update the
> `refreshToken` field in its state file while preserving subscription metadata,
> then restart it. Merely changing plugin configuration does not expand the
> scopes of an existing refresh token.

Personal Microsoft accounts normally consent during login. Work/school tenants
can apply policies that require an administrator to approve some or all
delegated permissions. If the consent screen says approval is required, a tenant
administrator must grant consent before login can mint a usable refresh token.

For app-registration steps, account audiences, public PKCE clients, and
confidential clients, see
[docs/custom-app-registration.md](docs/custom-app-registration.md).

## Configuration

| Plugin field | Environment | Purpose |
|---|---|---|
| `clientId` | `M365_CLIENT_ID` | Azure application/client ID |
| `clientSecret` | `M365_CLIENT_SECRET` | Optional confidential-client secret |
| `refreshToken` | `M365_REFRESH_TOKEN` | Direct-mode refresh token |
| `directTokenStatePath` | `M365_DIRECT_TOKEN_STATE_PATH`, then `~/.openclaw/state/m365-direct-token.json` | Direct-mode durable token state |
| `tenant` | `M365_TENANT`, then `consumers` | OAuth tenant such as `consumers`, `common`, or a tenant ID |
| `tokenBrokerUrl` | `M365_TOKEN_BROKER_URL` | Broker `/token` URL |
| `tokenBrokerSecret` | `M365_TOKEN_BROKER_SECRET` | Broker bearer secret |
| `features` | `M365_FEATURES`, then `calendar-write,mail-write,mail-send,tasks-write` | Enabled tool capabilities; use the same list in the broker |
| `personalCalendarNames` | `M365_PERSONAL_CALENDAR_NAMES` | Extra comma-separated personal calendar names |
| `familyCalendarNames` | `M365_FAMILY_CALENDAR_NAMES` | Extra comma-separated family calendar names |

## Tools

### Outlook mail, calendar, and task tools

All existing names remain unchanged:

- Mail: `outlook_inbox`, `outlook_search`, `outlook_read`,
  `outlook_save_attachments`, `outlook_send`, `outlook_reply`,
  `outlook_forward`, `outlook_move`, `outlook_flag`
- Calendar: `outlook_calendar_fetch`, `outlook_create_event`,
  `outlook_update_event`, `outlook_delete_event`, `outlook_meeting`,
  `outlook_query_events`
- Tasks: `outlook_task_lists`, `outlook_tasks`, `outlook_create_task`,
  `outlook_update_task`, `outlook_complete_task`, `outlook_delete_task`

### OneDrive tools

These tools are hidden by default. Enable `onedrive-read` for the read tools or
`onedrive-write` for both read and write tools, in both plugin and broker config.

OneDrive accepts item IDs or drive-root-relative paths. Parameters that offer
both forms reject ambiguous requests containing both.

| Tool | Main parameters | Description |
|---|---|---|
| `onedrive_list` | `item_id` or `path`, `limit` | List root or folder children |
| `onedrive_search` | `query`, `limit` | Search files and folders |
| `onedrive_metadata` | `item_id` or `path` | Get item/root metadata |
| `onedrive_download` | `item_id` or `path`, `output_path`, `overwrite` | Download locally; existing files are preserved by default |
| `onedrive_upload` | `local_path`, `name`, `parent_id` or `parent_path`, `overwrite` | Simple upload up to 4 MiB |
| `onedrive_create_folder` | `name`, `parent_id` or `parent_path` | Create a folder with conflict behavior `fail` |
| `onedrive_move` | source `item_id` or `path`, `new_name`, destination `parent_id` or `parent_path` | Move and/or rename |
| `onedrive_delete` | `item_id` or `path` | Delete an item |

Remote paths reject empty, `.` and `..` segments. Local downloads never
overwrite unless `overwrite=true`. Uploads also refuse an existing remote name
unless explicitly allowed.

## Token and state behavior

- Access tokens are cached by normalized delegated-scope set until shortly
  before expiration.
- Token refresh is serialized, including across different scope sets, so
  rotating refresh tokens cannot be reused concurrently.
- The webhook service atomically writes rotated refresh tokens and subscription
  metadata to one state file with restrictive permissions where supported.
- Direct mode atomically writes its rotated refresh token to its own state file
  and is supported only when exactly one plugin process uses that state.
- A normal shutdown preserves the Graph subscription. Restarted services reuse
  or renew it, avoiding notification gaps caused by delete-and-recreate.
- Tokens and broker secrets are never written to logs.

## Development

```powershell
npm install
npm test
npm run build
```

`npm test` covers the plugin, shared token manager, OneDrive tools, and token
broker. `npm run build` builds the shared auth package, plugin/CLI/adapter, and
standalone webhook service.
