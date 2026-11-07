# Outlook MCP (TypeScript)

Mail and calendars of every mailbox OE Copilot is signed in to, over Outlook REST v2.0. Same 12 tools as the Python server (`list_accounts`, `list_folders`, `list_emails`, `search_emails`, `read_email`, `read_thread`, `list_calendars`, `list_calendar_events`, `send_email`, `forward_email`, `get_attachment`, `manage_email`).

The server signs nothing in and stores no token on disk. OE Copilot owns sign-in and renewal for every mailbox (Outlook Mobile client, or the Outlook web session for tenants that block Outlook Mobile); this server reads current access tokens from its `GET /api/token-posta` and keeps them in memory.

## Run

Node 24 or later (type stripping), no build step.

```powershell
npm ci
npm run typecheck
npm test
node src/main.ts
```

Register in Claude Code:

```powershell
claude mcp add --scope user outlook -- node "C:\Users\Antonio\Downloads\personal projects\outlook-remote\ts\src\main.ts"
```

## Token broker

| Variable | Default | Meaning |
|---|---|---|
| `OUTLOOK_MCP_SSH_HOST` | `oecopilot-vm` | SSH host of OE Copilot; the request runs inside the agent container with its own `API_KEY`, which never leaves the host |
| `OUTLOOK_MCP_AGENT_CONTAINER` | `oe-copilot-agent` | Agent container |
| `OUTLOOK_MCP_AGENT_PORT` | `8180` | Agent port inside the container |
| `OUTLOOK_MCP_BROKER_URL` | | Direct HTTP instead of SSH: https, or http on localhost only |
| `OUTLOOK_MCP_BROKER_KEY` | | Agent `API_KEY`, required with `OUTLOOK_MCP_BROKER_URL` |

A broker row is used only when its token is for the Outlook resource, carries a mail permission, names the same mailbox as the row, and has more than two minutes left. The API address is fixed in the code, never taken from the broker answer. A 401 from Outlook reloads tokens once.

## Differences from the Python server

- No local token sources (New Outlook IdentityCache, phone over SSH): OE Copilot is the only source.
- `account` fragments that match several mailboxes are refused instead of taking the first.
- Plain text containing `< b` is no longer mistaken for HTML (the Python check allowed a space after `<`).
- `get_attachment` never overwrites an existing file; it saves as `name (2).ext`.
- `list_calendar_events` requires start and end with a zone, takes `time_zone` for the returned times, and reports `has_more`.
