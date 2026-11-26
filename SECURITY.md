# Security

The app uses Microsoft Graph delegated permissions for the signed-in account.
The default features request `Mail.ReadWrite`, `Mail.Send`,
`Calendars.ReadWrite`, and `Tasks.ReadWrite`. Read-only features request
`Mail.Read`, `Calendars.Read`, `Tasks.Read`, or `Files.Read`. OneDrive write
access requires `Files.ReadWrite`. Interactive login adds `offline_access` and
`openid`. Select only the features needed for each account, and mirror those
features in the broker and plugin configuration.

Use the login script's `--token-file` argument to save the initial refresh token
to a new private file. Browser login without that option prints the token to
the terminal. Keep token files, terminal output, and service credentials out of
version control. The broker persists rotated refresh tokens in its state file,
one file per configured account. Direct mode keeps its token in
`~/.openclaw/state/m365-direct-token.json` by default and supports a single
process. These state files are protected by filesystem permissions where
supported. Access tokens are cached in process memory.

Keep broker endpoints private. Expose only the Graph webhook route through your
public HTTPS reverse proxy. The broker checks a bearer secret and enforces
its configured scope allowlist before returning access tokens. The webhook
matches the subscription and checks its client state before dispatch. Configure
separate secrets, client state, and storage for each account.

OpenClaw agent context selects the configured broker account. Assign agents
explicitly with `accounts.<id>.agents` and `defaultAgents`, and configure each
account's rule routes with `allowedAgents`. Unassigned agents are refused once
assignments are configured. Keep broker credentials, state files, and rules
outside agent-accessible filesystems and shells. The gateway and broker
operator retain access to the credentials for the accounts they manage.
