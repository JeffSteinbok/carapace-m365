# 🐚📧 carapace-outlook

Unified mail, calendar, and task tools for Outlook / Microsoft 365 via Microsoft Graph API. Replaces the former `outlook-mail` and `outlook-calendar` plugins.

Built with [🦞🐚 Carapace](https://github.com/JeffSteinbok/carapace-plugin-sdk).

## Setup

Get running in under a minute using the **published app registration** — no
Azure setup required. (The client ID of a public/PKCE app is not a secret, so
carapace-outlook ships one as the default.)

### 1. Sign in and mint a refresh token

From the plugin directory:

```bash
npm run login
```

This opens your browser, you consent with your **personal Microsoft account**,
and it prints an `OUTLOOK_REFRESH_TOKEN`. That refresh token is the only
credential you need — there is no client secret.

### 2. Set the environment variable

```bash
OUTLOOK_REFRESH_TOKEN=your-refresh-token
```

### 3. Configure the plugin

```json
{
  "plugins": {
    "entries": {
      "outlook": {
        "enabled": true,
        "config": {
          "refreshToken": "${OUTLOOK_REFRESH_TOKEN}"
        }
      }
    }
  }
}
```

`clientId` defaults to the published app registration and `clientSecret` is not
used. Credentials can be passed directly in config or via environment variables
— env vars are the recommended approach.

> **Prefer your own Azure app?** If you want a different account-type audience,
> your own name on the consent screen, a registration you control, or the older
> confidential-client (client-secret) flow, see
> **[docs/custom-app-registration.md](docs/custom-app-registration.md)**. It
> covers registering your own public-client or confidential-client app, a
> pros/cons comparison of the two auth mechanisms, and the manual OAuth flow.

## Configuration Reference

| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `clientId` | string | No | Azure app client ID. Defaults to the published carapace-outlook app registration; override to use your own |
| `clientSecret` | string | No | Azure app client secret. Omit for public-client (PKCE) apps; provide only for confidential-client registrations |
| `refreshToken` | string | Yes | OAuth refresh token with mail + calendar + task scopes |
| `personalCalendarNames` | string[] | No | Additional calendar names to treat as personal (default: `["Calendar", "calendar"]`) |
| `familyCalendarNames` | string[] | No | Additional calendar names to treat as family (default: `["Your Family", "Family"]`) |

---

## Tools

### Mail

| Tool | Description |
|------|-------------|
| [`outlook_inbox`](#outlook_inbox) | List recent messages from inbox or another folder |
| [`outlook_search`](#outlook_search) | Search messages by text, sender, subject, or date range |
| [`outlook_read`](#outlook_read) | Read a specific message including full body |
| [`outlook_save_attachments`](#outlook_save_attachments) | Save attachments from a message to a local directory |
| [`outlook_send`](#outlook_send) | Send a plain-text email with optional attachments |
| [`outlook_reply`](#outlook_reply) | Reply to a message with proper threading |
| [`outlook_forward`](#outlook_forward) | Forward a message to new recipients |
| [`outlook_move`](#outlook_move) | Move a message to a different folder |
| [`outlook_flag`](#outlook_flag) | Flag, complete, or unflag a message |

### Calendar

| Tool | Description |
|------|-------------|
| [`outlook_calendar_fetch`](#outlook_calendar_fetch) | Fetch upcoming events from personal, family, or all calendars |
| [`outlook_create_event`](#outlook_create_event) | Create a new calendar event |
| [`outlook_update_event`](#outlook_update_event) | Update an existing event by ID |
| [`outlook_delete_event`](#outlook_delete_event) | Delete a calendar event by ID |
| [`outlook_meeting`](#outlook_meeting) | Create a meeting and send invites to attendees |
| [`outlook_query_events`](#outlook_query_events) | Query events by date range, text, attendee, or UID |

### Tasks

| Tool | Description |
|------|-------------|
| [`outlook_task_lists`](#outlook_task_lists) | List Microsoft To Do lists available on the account |
| [`outlook_tasks`](#outlook_tasks) | List tasks from a Microsoft To Do list |
| [`outlook_create_task`](#outlook_create_task) | Create a Microsoft To Do task |
| [`outlook_update_task`](#outlook_update_task) | Update an existing task |
| [`outlook_complete_task`](#outlook_complete_task) | Mark a task complete |
| [`outlook_delete_task`](#outlook_delete_task) | Delete a task |

---

## Tool Reference

<a id="outlook_inbox"></a>

### `outlook_inbox`

| Param | Type | Description |
|-------|------|-------------|
| `folder` | string | Folder name (default: `inbox`). Well-known names: `inbox`, `junkemail`, `deleteditems`, `sentitems`, `drafts`, `outbox`, `archive` |
| `limit` | number | Max messages to return (default: `10`) |
| `unread` | boolean | Only return unread messages |

<a id="outlook_search"></a>

### `outlook_search`

| Param | Type | Description |
|-------|------|-------------|
| `query` | string | Full-text search across subject and body |
| `from` | string | Filter by sender email or domain |
| `subject` | string | Subject substring filter |
| `since` | string | Start date `YYYY-MM-DD` |
| `before` | string | End date `YYYY-MM-DD` |
| `limit` | number | Max results (default: `10`) |

<a id="outlook_read"></a>

### `outlook_read`

| Param | Type | Description |
|-------|------|-------------|
| `message_id` | string | Microsoft Graph message ID |

<a id="outlook_save_attachments"></a>

### `outlook_save_attachments`

| Param | Type | Description |
|-------|------|-------------|
| `message_id` | string | Microsoft Graph message ID |
| `output_dir` | string | Local directory to save attachments into |
| `content_types` | string[] | Content-type filters (default: `["image/*"]`) |

<a id="outlook_send"></a>

### `outlook_send`

| Param | Type | Description |
|-------|------|-------------|
| `to` | string \| string[] | Recipient email address(es) |
| `subject` | string | Email subject |
| `body` | string | Plain-text body |
| `cc` | string[] | CC recipients |
| `attachment` | string[] | Local file path(s) to attach |
| `in_reply_to` | string | Message-ID for threading (include angle brackets) |
| `references` | string | Space-separated Message-IDs for full thread References header |
| `signature` | string | Signature block appended after body |

<a id="outlook_reply"></a>

### `outlook_reply`

| Param | Type | Description |
|-------|------|-------------|
| `message_id` | string | Graph message ID to reply to |
| `body` | string | Reply body |
| `reply_all` | boolean | Reply to all recipients (default: `false`) |
| `signature` | string | Signature block |

<a id="outlook_forward"></a>

### `outlook_forward`

| Param | Type | Description |
|-------|------|-------------|
| `message_id` | string | Graph message ID to forward |
| `to` | string \| string[] | Recipient(s) to forward to |
| `comment` | string | Optional note to prepend |

<a id="outlook_move"></a>

### `outlook_move`

| Param | Type | Description |
|-------|------|-------------|
| `message_id` | string | Graph message ID to move |
| `destination_folder` | string | Target folder name or well-known name |

<a id="outlook_flag"></a>

### `outlook_flag`

| Param | Type | Description |
|-------|------|-------------|
| `message_id` | string | Graph message ID |
| `flag_status` | string | `flagged`, `complete`, or `notFlagged` |

<a id="outlook_calendar_fetch"></a>

### `outlook_calendar_fetch`

| Param | Type | Description |
|-------|------|-------------|
| `calendar` | string | `personal`, `family`, or `all` (default: `all`) |
| `days` | number | Days ahead to fetch (default: `7`) |

<a id="outlook_create_event"></a>

### `outlook_create_event`

| Param | Type | Description |
|-------|------|-------------|
| `subject` | string | Event title |
| `start` | string | Start datetime ISO (e.g. `2026-03-15T14:00`) |
| `duration` | string | Duration string e.g. `1h`, `30m` (default: `1h`; ignored if `end` supplied) |
| `end` | string | End datetime ISO (overrides duration) |
| `timezone` | string | IANA timezone (default: `America/Los_Angeles`) |
| `location` | string | Event location |
| `description` | string | Event description/body |
| `attendees` | string[] | Attendee email addresses |
| `calendar` | string | `personal` or `family` (default: `personal`) |

<a id="outlook_update_event"></a>

### `outlook_update_event`

| Param | Type | Description |
|-------|------|-------------|
| `event_id` | string | Graph event ID (from `outlook_calendar_fetch`) |
| `subject` | string | New title |
| `start` | string | New start datetime ISO |
| `end` | string | New end datetime ISO |
| `duration` | string | New duration (if no `end`) |
| `timezone` | string | IANA timezone for `start`/`end` |
| `location` | string | New location |
| `description` | string | New description |
| `add_attendees` | string[] | Emails to add as attendees |
| `remove_attendees` | string[] | Emails to remove from attendees |
| `status` | string | `confirmed`, `tentative`, or `cancelled` |

<a id="outlook_delete_event"></a>

### `outlook_delete_event`

| Param | Type | Description |
|-------|------|-------------|
| `event_id` | string | Graph event ID to delete |

<a id="outlook_meeting"></a>

### `outlook_meeting`

| Param | Type | Description |
|-------|------|-------------|
| `to` | string \| string[] | Required attendee email(s) |
| `cc` | string[] | Optional attendees (marked optional/informational) |
| `subject` | string | Meeting title |
| `start` | string | Start datetime ISO |
| `duration` | string | Duration string (default: `1h`) |
| `end` | string | End datetime ISO (overrides duration) |
| `timezone` | string | IANA timezone (default: `America/Los_Angeles`) |
| `location` | string | Meeting location |
| `description` | string | Agenda / meeting notes |
| `signature` | string | Signature block for the invite |

<a id="outlook_query_events"></a>

### `outlook_query_events`

| Param | Type | Description |
|-------|------|-------------|
| `after` | string | Events at or after this date (ISO, e.g. `2026-03-01`) |
| `before` | string | Events before this date (ISO, e.g. `2026-04-01`) |
| `text` | string | Filter by title/description text |
| `attendee` | string | Filter to events including this attendee email |
| `uid` | string | Return the single event with this exact iCalUId |

---

## Building

```bash
cd plugins/outlook
npm install && npm run build
node dist/bin/outlook.js --help
```

Built with [Carapace Plugin SDK](https://github.com/JeffSteinbok/carapace-plugin-sdk).
