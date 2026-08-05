# Registering your own Microsoft 365 app

The current published client ID remains available, and the repository/app can be
renamed later without changing the plugin ID. Use your own Microsoft Entra app
registration when you need work/school accounts, a consent screen you control,
or a confidential-client deployment.

## Public PKCE or confidential client

| | Public client + PKCE (recommended for local CLI use) | Confidential client |
|---|---|---|
| Client secret | None | Required |
| Redirect platform | Mobile and desktop application | Web |
| Redirect URI | `http://localhost` | `http://localhost` |
| Login | `npm run login -- --client-id ...` | Authorization-code flow with secret |
| Best fit | Desktop, CLI, personal server | Locked-down server able to protect a secret |

A public client ID is not a secret. A refresh token and token-broker secret are
sensitive and must be protected.

## Create the app registration

1. Open **Microsoft Entra ID → App registrations → New registration**.
2. Choose the supported account audience:
   - personal accounts only: login tenant `consumers`;
   - work/school and personal: login tenant `common`;
   - one organization: use that tenant ID.
3. Copy the **Application (client) ID**.
4. For public PKCE:
   - **Authentication → Add a platform → Mobile and desktop applications**;
   - add `http://localhost`;
   - enable **Allow public client flows**.
5. For a confidential client:
   - add `http://localhost` as a Web redirect;
   - create a secret under **Certificates & secrets** and protect its value.

## Add only the delegated permissions you want

Under **API permissions → Add a permission → Microsoft Graph → Delegated
permissions**, select only the capabilities you intend to use:

| Feature | Capability | Delegated permission |
|---|---|---|
| `mail-read` | Read mail and receive inbox webhooks | `Mail.Read` |
| `mail-write` | Move/flag/manage mail; also enables reads | `Mail.ReadWrite` |
| `mail-send` | Send/reply/forward | `Mail.Send` |
| `calendar-read` | Read calendars | `Calendars.Read` |
| `calendar-write` | Modify calendars; also enables reads | `Calendars.ReadWrite` |
| `tasks-read` | Read To Do tasks | `Tasks.Read` |
| `tasks-write` | Modify To Do tasks; also enables reads | `Tasks.ReadWrite` |
| `onedrive-read` | Read/download OneDrive | `Files.Read` |
| `onedrive-write` | Modify OneDrive; also enables reads | `Files.ReadWrite` |
| Login-only | Issue refresh tokens during login | `offline_access` |

Do not add `.All` application permissions. This integration uses delegated
permissions as the signed-in user.

Adding permissions in Azure only makes them available for consent. It does
**not** grant access to an existing refresh token. After adding a permission,
run login again with the matching feature/scopes, complete consent, and replace
the broker-owned refresh token in its state file.

Work/school tenant policy may require administrator consent. Personal Microsoft
accounts usually allow user consent for these delegated permissions.

The default feature list is
`calendar-write,mail-write,mail-send,tasks-write`. It preserves pre-OneDrive
behavior; OneDrive tools and broker scopes remain disabled until explicitly
configured.

## Public-client login examples

```powershell
$env:M365_CLIENT_ID = "your-client-id"

# Mail only
npm run login -- --tenant common --features mail-write,mail-send

# Calendar read-only
npm run login -- --tenant common --features calendar-read

# OneDrive read-only
npm run login -- --tenant common --features onedrive-read

# Full OneDrive
npm run login -- --tenant common --features onedrive-write

# Existing Outlook features plus OneDrive
npm run login -- --tenant common --features calendar-write,mail-write,mail-send,tasks-write,onedrive-write
```

For personal-account-only apps, omit `--tenant common` or use
`--tenant consumers`. Exact delegated scopes can be requested with `--scopes`.

The resulting value should normally be owned by the webhook/token broker:

```text
M365_CLIENT_ID=your-client-id
M365_REFRESH_TOKEN=returned-refresh-token
M365_FEATURES=calendar-write,mail-write,mail-send,tasks-write
M365_TOKEN_BROKER_SECRET=long-random-bearer-secret
```

Use the same list in the plugin:

```json
{
  "features": ["calendar-write", "mail-write", "mail-send", "tasks-write"]
}
```

`OUTLOOK_CLIENT_ID`, `OUTLOOK_REFRESH_TOKEN`, and
`OUTLOOK_TOKEN_BROKER_SECRET` remain accepted aliases. `OUTLOOK_FEATURES` is
also accepted as the broker/plugin environment fallback.

The list participates in three independent enforcement layers: plugin tool
registration, the broker's authoritative scope allowlist, and Microsoft
consent. Configure the same features in the first two layers, then consent to
their scopes. Unknown features and scopes fail closed.

## Confidential-client authorization

Use the same feature-specific scope list, plus `offline_access` and `openid`, in
the authorization request:

```text
https://login.microsoftonline.com/common/oauth2/v2.0/authorize
  ?client_id=YOUR_CLIENT_ID
  &response_type=code
  &redirect_uri=http%3A%2F%2Flocalhost
  &response_mode=query
  &scope=Mail.ReadWrite+Mail.Send+offline_access+openid
```

Exchange the code at:

```text
POST https://login.microsoftonline.com/common/oauth2/v2.0/token
Content-Type: application/x-www-form-urlencoded

client_id=YOUR_CLIENT_ID
client_secret=YOUR_CLIENT_SECRET
code=THE_AUTHORIZATION_CODE
redirect_uri=http://localhost
grant_type=authorization_code
scope=Mail.ReadWrite Mail.Send offline_access openid
```

Configure `M365_CLIENT_SECRET` (or `OUTLOOK_CLIENT_SECRET`) in the broker. The
shared token module omits `client_secret` entirely for public clients.

## Incremental consent and broker state

When expanding permissions:

1. Add the delegated permission in Azure.
2. Re-run `npm run login` with all capabilities the replacement token should
   retain.
3. Stop `m365-webhook`.
4. Replace only the `refreshToken` value in the configured state JSON, preserving
   all subscription metadata, including `subscriptionId`,
   `expirationDateTime`, `notificationUrl`, and `clientState`.
5. Ensure the state remains readable only by the service account where the
   operating system supports file modes.
6. Restart the service.

The plugin must not keep a separate refresh token when broker mode is enabled.
That prevents competing refresh-token rotations.

Direct mode is only supported for a single plugin process. It owns a separate
durable token state file (by default
`~/.openclaw/state/m365-direct-token.json`). Deployments with multiple plugin
processes must use the broker so refresh-token rotation has one owner.
