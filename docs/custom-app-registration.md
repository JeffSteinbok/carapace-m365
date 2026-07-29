# Using your own Azure app registration

By default, carapace-outlook uses a **published public-client app registration**
(client ID `0c3df71b-4dc2-49a7-b6e7-e5c3c48bf501`). Because the client ID of a
public (PKCE) client is **not a secret**, you can use it directly: run
`npm run login`, consent with your own personal Microsoft account, and you're
done — no Azure setup required. See the main [README](../README.md) for that
quick start.

You only need your own app registration if you want to:

- use a **different account-type audience** (e.g. work/school accounts),
- show **your own name** on the consent screen instead of the published app's,
- keep your usage under a registration **you control**, or
- use the older **confidential-client (client secret)** flow.

This document covers both ways to register and authenticate your own app.

---

## Auth mechanism comparison

carapace-outlook authenticates to Microsoft Graph with a refresh token. There are
two ways to obtain that token; the plugin supports **both**.

| | **Public client + PKCE** (recommended) | **Confidential client + secret** |
|---|---|---|
| Client secret | **None** | Required (`OUTLOOK_CLIENT_SECRET`) |
| Azure setup | "Allow public client flows" = Yes; `http://localhost` as a *Mobile & desktop* redirect | Client secret created under *Certificates & secrets* |
| How you get the token | `npm run login` (browser PKCE flow) | Manual auth-code + `curl` exchange |
| Best for | Local / desktop / CLI installs (openclaw on your own machine) | Server-side apps that can genuinely keep a secret hidden |
| **Pros** | Nothing sensitive stored on the box; matches Microsoft's guidance for native/desktop/CLI apps; a leaked config can't be replayed without the user re-consenting; no secret to rotate or expire | Familiar, widely-documented flow; the secret + refresh token together are a stable long-lived credential |
| **Cons** | The MSA refresh token still rotates/expires (~90 days idle) — re-run `npm run login` to refresh; requires the public-client toggle on the app | A client secret on a local machine adds **no real security** (a public client can't hide it) yet is another credential to store, leak, and rotate; secrets have hard expiry dates and silently break the plugin when they lapse |

> **TL;DR:** For an openclaw instance running on your own machine, use **public
> client + PKCE**. A client secret only makes sense when the plugin runs somewhere
> that can actually keep the secret confidential (a locked-down server), which is
> rarely the case for a personal install.

---

## Option A — Public client + PKCE (recommended, no secret)

### 1. Register (or configure) the Azure app

1. Go to the [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) and open (or create) your app.
2. **Supported account types:** *Personal Microsoft accounts only* (or *any org directory + personal* if you also need work accounts).
3. **Authentication → Add a platform → Mobile and desktop applications**, add redirect URI `http://localhost`.
4. **Authentication → Advanced settings → Allow public client flows → Yes.**
5. Copy the **Application (client) ID** — this is your `OUTLOOK_CLIENT_ID`. You do **not** need a client secret.

### 2. Configure API permissions

**API permissions → Add a permission → Microsoft Graph → Delegated permissions**, then add:

- `Calendars.ReadWrite` — read and write calendar events
- `Mail.ReadWrite` — read, move, flag, and manage mail
- `Mail.Send` — send email
- `Tasks.ReadWrite` — create and manage Microsoft To Do tasks
- `offline_access` — required to get a refresh token

Personal accounts consent at sign-in time; no admin consent needed.

### 3. Mint a refresh token

From the plugin directory, pass your own client ID:

```bash
OUTLOOK_CLIENT_ID=your-app-client-id npm run login
# or: npm run login -- --client-id your-app-client-id
```

This opens your browser, runs the Authorization Code + PKCE flow against a
loopback redirect (`http://localhost:53682` by default — override with
`--port`), and prints the `OUTLOOK_REFRESH_TOKEN` to set. No secret is involved
at any point.

### 4. Set environment variables

```bash
OUTLOOK_CLIENT_ID=your-app-client-id
OUTLOOK_REFRESH_TOKEN=your-refresh-token
# no OUTLOOK_CLIENT_SECRET
```

---

## Option B — Confidential client + secret

### 1. Register the Azure app

1. In [App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade), click **New registration**.
2. **Name:** anything (e.g. `carapace-outlook`).
3. **Supported account types:** *Personal Microsoft accounts only* (or *any org directory + personal*).
4. **Redirect URI:** *Web* → `http://localhost`.
5. Copy the **Application (client) ID** — `OUTLOOK_CLIENT_ID`.

### 2. Create a client secret

**Certificates & secrets → Client secrets → New client secret**, then copy the
**Value** immediately (shown once) — this is your `OUTLOOK_CLIENT_SECRET`.

### 3. Configure API permissions

Same delegated permissions as Option A (`Calendars.ReadWrite`, `Mail.ReadWrite`,
`Mail.Send`, `Tasks.ReadWrite`, `offline_access`).

### 4. Get a refresh token (manual OAuth flow)

Generate the authorization URL (replace `YOUR_CLIENT_ID`):

```
https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize
  ?client_id=YOUR_CLIENT_ID
  &response_type=code
  &redirect_uri=http%3A%2F%2Flocalhost
  &scope=Calendars.ReadWrite+Mail.ReadWrite+Mail.Send+Tasks.ReadWrite+offline_access
  &response_mode=query
```

> **Note:** Use `/consumers/` for personal Microsoft accounts. Use `/common/` only if your app registration is set to **All** audience.

1. Open the URL, sign in, and accept the permissions.
2. You'll be redirected to `http://localhost/?code=...` — copy the `code`.
3. Exchange it for tokens:

```bash
curl -X POST https://login.microsoftonline.com/consumers/oauth2/v2.0/token \
  -H "Content-Type: application/x-www-form-urlencoded" \
  -d "client_id=YOUR_CLIENT_ID" \
  -d "client_secret=YOUR_CLIENT_SECRET" \
  -d "code=YOUR_CODE" \
  -d "redirect_uri=http://localhost" \
  -d "grant_type=authorization_code"
```

4. Copy the `refresh_token` — this is your `OUTLOOK_REFRESH_TOKEN`.

### 5. Set environment variables

```bash
OUTLOOK_CLIENT_ID=your-app-client-id
OUTLOOK_CLIENT_SECRET=your-client-secret
OUTLOOK_REFRESH_TOKEN=your-refresh-token
```

---

## Configure the plugin

**Public client (Option A)** — client ID + refresh token:

```json
{
  "plugins": {
    "entries": {
      "outlook": {
        "enabled": true,
        "config": {
          "clientId": "${OUTLOOK_CLIENT_ID}",
          "refreshToken": "${OUTLOOK_REFRESH_TOKEN}"
        }
      }
    }
  }
}
```

**Confidential client (Option B)** — add `clientSecret`:

```json
{
  "plugins": {
    "entries": {
      "outlook": {
        "enabled": true,
        "config": {
          "clientId": "${OUTLOOK_CLIENT_ID}",
          "clientSecret": "${OUTLOOK_CLIENT_SECRET}",
          "refreshToken": "${OUTLOOK_REFRESH_TOKEN}"
        }
      }
    }
  }
}
```

When `clientSecret` is omitted, the plugin uses the public-client token flow
automatically.

> **Important (both options):** The refresh token is long-lived but scoped. If
> you add permissions later, you must repeat the flow — updating only the token
> without re-consenting will not grant new scopes.
