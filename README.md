# 🐚 📧 📅 ✅ 📁 Carapace Microsoft 365

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
| `onedrive-read` | List, search, inspect, read text, or download OneDrive content | `Files.Read` |
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
| `access` | — | Optional per-calendar read/write access control (see below) |

### Calendar access control

Add an `access` block to restrict which agents and verified senders can read
or write each calendar. The calendar ids are `personal` (the default calendar)
and `family`. The rules come from
[carapace-plugin-sdk](https://github.com/JeffSteinbok/carapace-plugin-sdk#call-context-and-access-control-v31):

- **A grant** is `"*"` or `{ "agent": ..., "sender": ... }`. Every field in a grant must match, and any one matching grant allows the call.
- **`agent`** is an OpenClaw agent id.
- **`sender`** names an OpenClaw root-config `accessGroups` entry (`type: "message.senders"`), or the special name `"owner"`.
- **Turns with no sender** (cron, `openclaw agent`, heartbeat) match only grants with no `sender` field.
- **No `access` block** allows everything, the same as before.
- **Default-deny** applies once `access.calendars` exists. An unlisted calendar, or an op with no grant list, is denied.

```json
{
  "accessGroups": {
    "jeff": { "type": "message.senders", "members": { "imessage": ["+14255550100", "jeff@example.com"], "discord": ["123456789012345678"] } }
  },
  "plugins": {
    "entries": {
      "m365": {
        "config": {
          "access": {
            "calendars": {
              "family":   { "read": ["*"], "write": ["*"] },
              "personal": { "read": [{ "agent": "main" }, { "sender": "jeff" }], "write": [{ "agent": "main" }] }
            }
          }
        }
      }
    }
  }
}
```

How each tool is checked:

| Tool | Check |
|---|---|
| `outlook_calendar_fetch` | read on `calendar`. With `all` (the default), only readable calendars are returned, and `omitted_calendars` lists the rest. It errors only when none are readable |
| `outlook_query_events` | read on `calendar` (`personal` when omitted), in every mode: date range, text, `uid`. An explicit `calendar` now scopes `uid` and text queries to that calendar |
| `outlook_create_event` | write on `calendar` (default `personal`) |
| `outlook_meeting` | write on `personal` |
| `outlook_update_event`, `outlook_delete_event` | The event's calendar is looked up first (`GET /me/events/{id}/calendar`), then write is required on it. Events on any other calendar are checked as `other:<name>` and are denied unless granted. The lookup only runs when `access.calendars` is set |

A denial is returned as a tool error, for example
`Access denied: agent "family" / sender +14255550100 may not read calendar "personal"`.
It carries no event data and is logged once through the plugin logger. The
standalone CLI reads config from the environment, so it never has an `access`
block and is not restricted.

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
| `onedrive_list` | `item_id` or `path`, `limit`, `cursor` | List root or folder children |
| `onedrive_search` | `query`, `limit`, `cursor` | Search files and folders |
| `onedrive_metadata` | `item_id` or `path` | Get item/root metadata |
| `onedrive_read` | `item_id` or `path`, `offset`, `max_chars` | Read a bounded window of UTF-8 text (file limit 1 MiB) |
| `onedrive_download` | `item_id` or `path`, `output_path`, `overwrite` | Download locally; existing files are preserved by default |
| `onedrive_upload` | `local_path`, `name`, `parent_id` or `parent_path`, `overwrite` | Simple upload up to 4 MiB |
| `onedrive_create_folder` | `name`, `parent_id` or `parent_path` | Create a folder with conflict behavior `fail` |
| `onedrive_move` | source `item_id` or `path`, `new_name`, destination `parent_id` or `parent_path` | Move and/or rename |
| `onedrive_delete` | `item_id` or `path` | Delete an item |

Remote paths reject empty, `.` and `..` segments. Local downloads never
overwrite unless `overwrite=true`. Uploads also refuse an existing remote name
unless explicitly allowed.

### Personal OneDrive configuration and exclusions

Use `tenant: "consumers"` (the default) for personal Microsoft accounts. For
example, the plugin config portion of your OpenClaw configuration can be:

```json
{
  "tenant": "consumers",
  "features": ["onedrive-read"],
  "oneDriveDeniedPaths": ["Private", "Documents/Financial", "Notes/passwords.txt"]
}
```

Paths are relative to your own drive root. Each exclusion blocks the exact file
or directory and every descendant; it is not a glob or substring match.
Matching is case-insensitive, NFC-normalized, treats backslashes as separators,
and conservatively decodes percent escapes (including nested escapes). Unsafe
or ambiguous paths fail closed. `/` or an empty exclusion blocks the whole drive.
Omit the property or use `[]` for no exclusions; malformed configuration fails
closed. Configure this in the plugin, not only in the token broker.

With exclusions active, all read entry points resolve fresh metadata and walk
parent IDs to the authenticated user's root. Search and list omit denied or
unverifiable entries, including remote/shared shortcuts; direct access reports
an error. Missing ancestry, cycles, and ancestry deeper than 128 items are
refused. Item IDs and misleading search parent paths do not bypass the checks.
Reads/downloads use the verified item ID and recheck ancestry before releasing
bytes. All OneDrive write handlers refuse operations while exclusions are
configured, including when `onedrive-write` is enabled, to prevent move/rename
routes around the policy. Use `onedrive-read` to keep write tools hidden.

These are **application-level exclusions**, not narrower Microsoft Graph OAuth
permissions. The OAuth token still has the consented `Files.Read` or
`Files.ReadWrite` scope; another client or direct Graph access is outside this
safeguard. Exclusions do not remove previously downloaded files, cached output,
or other copies. Graph has no atomic path-policy/content transaction: concurrent
external moves/renames cannot be completely eliminated as a race. Changes to
exclusions take effect on subsequent calls using the updated plugin config.
No OAuth grants or deployed configuration are changed by configuring this option.

### Reading and enumerating

`onedrive_read` accepts `.txt`, `.md`, `.markdown`, `.csv`, `.tsv`, `.json`,
`.jsonl`, `.xml`, `.yaml`, `.yml`, `.log`, `.html`, `.css`, `.js`, `.ts`, `.py`,
`.sh`, `.sql`, `.ini`, and `.toml`. It returns raw text, never renders HTML or
executes code, and rejects invalid UTF-8 and binary control characters. Office,
PDF, archives, images, UTF-16 files, and other binary formats are not extracted;
use `onedrive_download` instead (100 MiB maximum). Size is checked before transfer
and actual response bytes are bounded, even if metadata understates the size.
No document-parser dependencies are added.

Text output defaults to 16,000 UTF-16 code units, with `max_chars` from 1 to
100,000. Pass the returned `next_offset` as `offset` to continue; `null` means
finished. Each call fetches the whole file (up to 1 MiB), so changing files are
not snapshot-consistent and a window may split a surrogate pair.

List/search return `{ count, items, next_cursor }`. Repeat the same arguments
(including query, folder and limit) with `cursor: next_cursor` until it is `null`.
`limit` is a Graph page size from 1 to 200 (default 50 for list, 25 for search).
Filtering can yield an empty page with a non-null cursor: continue that cursor.
List enumerates immediate children; call it on returned folders to recurse.
Search uses Graph's indexed name/content matching, not a complete filesystem scan.

Cursors are authenticated, bound to the request and configuration, expire after
one hour, and stop working after a plugin process restart or config change.
Restart enumeration in that case. A cursor is not a raw URL; arbitrary Graph
routes and cross-host continuations are rejected. Policy checks are repeated on
every page, at the cost of extra Graph requests and possible throttling for large
trees. Graph JSON responses are bounded to 4 MiB.

Graph API references: [children and paging](https://learn.microsoft.com/en-us/graph/api/driveitem-list-children?view=graph-rest-1.0),
[search](https://learn.microsoft.com/en-us/graph/api/driveitem-search?view=graph-rest-1.0),
and [file content](https://learn.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0).

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

### Deploying a local build to OpenClaw

Install your local build as a packed, managed plugin instead of pointing
`plugins.load.paths` at this checkout (which drags the dev `node_modules`
into OpenClaw's update snapshots):

```bash
npm run deploy              # build → pack to out/m365.tgz → print install command
npm run deploy -- --install # same, then run the install command
```

The install command is `openclaw plugins install npm-pack:<repo>/out/m365.tgz --force`.
OpenClaw installs the pack into its managed plugin root along with the
package's production `dependencies`. What goes in the pack is controlled by
`"files"` in `package.json`; the deploy script lists the packed files and
refuses to deploy if `src/`, `tests/` or `node_modules/` end up in it.

The pack contains only the plugin (`dist/` with the shared auth package
bundled in). The token broker/webhook service (`services/m365-webhook`) and
`npm run login` are separate and keep running from this checkout; the
plugin reaches the broker over HTTP via `tokenBrokerUrl`.
