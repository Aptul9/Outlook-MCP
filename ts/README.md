# Communication MCP (TypeScript)

One remote MCP endpoint for mail, calendars and Teams chats. Mail and calendars: every mailbox OE Copilot is signed in to, over Outlook REST v2.0. Teams: the tools of a TeamsRelay `/mcp` endpoint, under `teams_` names. Clients hold one bearer token; every backend credential stays on the server.

The server signs nothing in and stores no token on disk. OE Copilot owns sign-in and renewal for every mailbox (Outlook Mobile client, or the Outlook web session for tenants that block Outlook Mobile); TeamsRelay owns the Teams sessions.

```text
MCP client --HTTPS, Bearer MCP_HTTP_TOKEN--> TeamsRelay Caddy --> communication-mcp:8190/mcp
  communication-mcp --> agent:8180 (OE Copilot: token broker or read gateway) --> Outlook REST
  communication-mcp --> webapp:8090/mcp (TeamsRelay, its MCP_TOKEN) --> Teams
```

## Remote server

`src/http.ts`, Streamable HTTP (stateless, one server instance per request). `npm start` runs it.

| Path | Answer |
|---|---|
| `POST/GET/DELETE /mcp` | MCP. Needs `Authorization: Bearer <MCP_HTTP_TOKEN>` (constant-time compare), a `Host` listed in `MCP_HTTP_HOSTS` (421 otherwise: DNS rebinding), no `Origin` header (403: no browser callers), body up to 4 MiB (413) |
| `GET /healthz` | `{"ok":true}`, no auth, no backend call |

| Variable | Default | Meaning |
|---|---|---|
| `MCP_HTTP_TOKEN` | | Client bearer token, at least 32 characters |
| `MCP_HTTP_HOSTS` | | Comma-separated `Host` values accepted on `/mcp` |
| `PORT` | `8190` | Listen port |
| `BIND_HOST` | `127.0.0.1` | Listen address (`0.0.0.0` in the container) |

Plus the backend variables below. Teams tools are discovered once at start and shared by every request; a relay that is down at start, or a call that loses the connection, triggers a rediscovery at most once a minute, and the list is refreshed every 10 minutes. Mail and calendar tools never wait for Teams.

Differences from local mode, because the caller's disk is not the server's: `send_email` and `forward_email` take `files` (`name`, `content_base64`, `content_type`) instead of local paths, and `get_attachment` returns the bytes in the result (images inline, other files as embedded resources) instead of saving them.

### Deployment

`compose.yml`, on the host of OE Copilot and TeamsRelay. The container joins `oe-copilot_default` and `teamsrelay_default`, reaches `agent:8180` and `webapp:8090` there, and listens on `127.0.0.1:8190` only. Backend addresses may be plain http to loopback or to a single-label container name; anything else must be https.

```bash
tar -czf cmcp.tgz --exclude=node_modules --exclude=test --exclude=.env -C ts .
scp cmcp.tgz host:/tmp/
# first time: cp .env.example .env, chmod 600, fill it; later runs keep .env
ssh host 'cd /opt/communication-mcp && find . -mindepth 1 -maxdepth 1 ! -name .env -exec rm -rf {} + \
  && tar --no-same-owner --no-same-permissions -xzf /tmp/cmcp.tgz && rm /tmp/cmcp.tgz \
  && chmod -R go-w . && chmod 750 . && docker compose up -d --build'
```

On `oe-vm` (2026-11-08): `/opt/communication-mcp`, broker mode, `OUTLOOK_MCP_BROKER_KEY` = the agent's `API_KEY` (an OE Copilot Action secret: a rotation there needs the same value here), `OUTLOOK_MCP_TEAMS_TOKEN` = TeamsRelay `MCP_TOKEN`.

Public name: a site in the TeamsRelay Caddy, which already receives every TLS name that sslh does not route elsewhere.

```caddyfile
mcp.example.com {
	reverse_proxy communication-mcp:8190
}
```

Client registration, O?penCode (`opencode.jsonc`):

```jsonc
"mcp": {
  "communication": {
    "type": "remote",
    "url": "https://mcp.example.com/mcp",
    "headers": { "Authorization": "Bearer {env:COMMUNICATION_MCP_TOKEN}" }
  }
}
```

Claude Code: `claude mcp add --transport http --scope user communication https://mcp.example.com/mcp --header "Authorization: Bearer <token>"`.

## Local stdio (compatibility)

Node 24 or later (type stripping), no build step. Same tools, local file paths for attachments.

```powershell
npm ci
npm run typecheck
npm test
npm run start:stdio
```

## Token broker

| Variable | Default | Meaning |
|---|---|---|
| `OUTLOOK_MCP_SSH_HOST` | `oecopilot-vm` | SSH host of OE Copilot; the request runs inside the agent container with its own `API_KEY`, which never leaves the host |
| `OUTLOOK_MCP_AGENT_CONTAINER` | `oe-copilot-agent` | Agent container |
| `OUTLOOK_MCP_AGENT_PORT` | `8180` | Agent port inside the container |
| `OUTLOOK_MCP_BROKER_URL` | | Direct HTTP instead of SSH: https, or http on localhost or a container name |
| `OUTLOOK_MCP_BROKER_KEY` | | Agent `API_KEY`, required with `OUTLOOK_MCP_BROKER_URL` |

A broker row is used only when its token is for the Outlook resource, carries a mail permission, names the same mailbox as the row, and has more than two minutes left. The API address is fixed in the code, never taken from the broker answer. A 401 from Outlook reloads tokens once.

## Differences from the Python server

- No local token sources (New Outlook IdentityCache, phone over SSH): OE Copilot is the only source.
- `account` fragments that match several mailboxes are refused instead of taking the first.
- Plain text containing `< b` is no longer mistaken for HTML (the Python check allowed a space after `<`).
- `get_attachment` never overwrites an existing file; it saves as `name (2).ext`.
- `list_calendar_events` requires start and end with a zone, takes `time_zone` for the returned times, and reports `has_more`.

## Consolidation boundary

The package exports `outlook-mcp/core`: mail/calendar clients and token-service adapters, without any MCP imports. The MCP layer only declares tools and formats results. No mailbox or body cache is created.

OE-Copilot remains the credential owner. Its optional communication API returns account metadata and executes constrained mail/calendar reads. Configure `OUTLOOK_MCP_SERVICE_URL` plus `OUTLOOK_MCP_BROKER_KEY`, or `OUTLOOK_MCP_SERVICE_SSH`. In this mode Microsoft bearer tokens do not leave OE-Copilot. Write tools are not registered in service mode; legacy broker mode retains the existing write tools.

Production services and their deployment pipelines remain separate. Teams forwarding is optional and preserves TeamsRelay's tool behavior, including read-state effects on refresh tools. No calendar source or production account is migrated by installing this package.

### Unified Teams tools

Set `OUTLOOK_MCP_TEAMS_URL` to the existing TeamsRelay `/mcp` endpoint and `OUTLOOK_MCP_TEAMS_TOKEN` to its bearer credential. Teams tools are discovered and exposed under `teams_` names. Mail/calendar tools keep their current names for compatibility. Schemas, annotations and content blocks remain those of TeamsRelay. An unavailable Teams endpoint does not disable mail/calendar. Existing services are not merged or redeployed.

### Unread completeness

`list_unread` searches all folders of all mailboxes, or selected accounts. Requests run with bounded parallelism. Each mailbox reports its own result or error, `has_more`, `incomplete`, page count and stop reason. Limits are 1000 summaries per mailbox and 50 pages scanned; reaching a safety bound is explicit. No read state changes and no local mailbox cache.

List/search follows same-collection Outlook pagination links. Unknown folder names now fail instead of quietly searching the whole mailbox. Folder lists recurse into child folders and follow pages, with explicit safety errors rather than silent truncation.

The cross-mailbox MCP unread tool defaults to 20 summaries per mailbox to bound tool output; an explicit limit can request up to 1000. Summary output is compact and full bodies remain on demand. The core API can be used independently of MCP through the package `./core` export. Dependency pins are SDK client/server 2.4.0 and zod 4.6.7; audit after upgrade reports no known vulnerabilities.
