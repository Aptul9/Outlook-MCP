# Outlook MCP

Full-featured MCP (Model Context Protocol) server for **New Outlook** (`olk.exe` / `Microsoft.OutlookForWindows`).

Reads emails, navigates entire conversation threads, searches messages, lists folders/accounts, and sends emails or thread replies. Works transparently across **all signed-in Microsoft 365 / Exchange accounts** on your PC.

**Zero third-party Python dependencies** (standard library only: `ctypes`, `urllib`, `sqlite3`, `json`).
**Zero app restarts**: harvests fresh OAuth2 bearer tokens directly from the Windows DPAPI-protected identity token cache (`%LOCALAPPDATA%\Microsoft\IdentityCache`), without touching or closing New Outlook.

---

## Tools

| Tool | Description | Key Arguments |
|---|---|---|
| `list_accounts` | Lists all detected Outlook accounts and token status | None |
| `list_folders` | Lists mail folders (Inbox, Sent, Archive, etc.) with total & unread counts | `account` |
| `list_emails` | Lists messages newest first with IDs, thread ID, sender, subject, preview | `folder`, `query`, `limit`, `unread_only`, `account` |
| `read_email` | Reads one email completely: headers, clean Markdown body, attachments list | `message_id`, `account` |
| `read_thread` | **Reads an entire conversation thread** in chronological order (oldest to newest) with full Markdown bodies and attachments list | `thread_id` (or any `message_id`), `limit`, `account` |
| `search_emails` | Full-text search across all folders | `query`, `limit`, `account` |
| `send_email` | Send a new email or post a reply to a thread/message | `to`, `subject`, `body`, `cc`, `bcc`, `reply_to_id`, `reply_all`, `attachments`, `account` |

---

## Requirements

- Windows 10 or 11
- New Microsoft Outlook (`olk.exe`), signed in to at least one account.
- Python 3.8 or later.

---

## Install

Run the installer via PowerShell:

```powershell
cd C:\Users\Samuele\Downloads\outlook-mcp
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
```

The script will:
1. Copy the server files to `%LOCALAPPDATA%\Programs\OutlookMCP\`
2. Verify token harvesting across your accounts
3. Register the server in Claude Code CLI as `outlook`: `claude mcp add --scope user outlook -- <python> <server>`

To also register in Claude Desktop:
```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -ClaudeDesktop
```

---

## Configuration for Other MCP Clients

```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": [
        "C:\\Users\\<YourUser>\\AppData\\Local\\Programs\\OutlookMCP\\outlook_mcp.py"
      ]
    }
  }
}
```

---

## Verification & Manual Testing

Run any tool directly from the terminal without starting an MCP client:

```powershell
python outlook_mcp.py --call list_accounts
python outlook_mcp.py --call list_emails "{'limit': 5}"
python outlook_mcp.py --call read_thread "{'thread_id': '<thread_or_message_id>'}"
python outlook_mcp.py --call search_emails "{'query': 'deploy'}"
```

---

## Sending Emails with Rich Formatting & Attachments

`send_email` supports full HTML formatting and arbitrary binary attachments.

### Example JSON invocation
```json
{
  "account": "prismaprogetti",
  "to": "d.fortunio@prismaprogetti.it",
  "subject": "[Test] Verifica Outlook MCP - Formattazione HTML ed Allegati",
  "is_html": true,
  "body": "<p>Ciao <b>Daniele</b>,</p><p>Badge: <span style=\"background-color:#36b37e;color:white;padding:2px 8px;border-radius:10px;\">TEST SUPERATO</span></p><p><mark style=\"background-color:#fff066;\"><b>Evidenziatore giallo</b></mark></p>",
  "attachments": [
    "C:\\Users\\Samuele\\AppData\\Local\\Temp\\opencode\\test_attachment.png"
  ]
}
```

### Supported Styles
- **Grassetti, corsivi, sottolineati**: `<b>`, `<i>`, `<u>`
- **Evidenziatori**: `<mark style="background-color: #fff066;">...</mark>` o `<span>` con `background-color`
- **Badge e Box**: `<div>` o `<span>` con stili CSS inline (bordi, padding, border-radius, colori di sfondo)
- **Liste e Tabelle**: tag HTML standard `<ul>`, `<li>`, `<table>`
- **Allegati**: lista di percorsi locali in `attachments` codificati in Base64 via `#Microsoft.OutlookServices.FileAttachment`.

---

## Replying with Cc, Bcc & Attachments

With `reply_to_id` set, `to`, `cc`, `bcc` and `attachments` all apply to the reply:

```json
{
  "account": "prismaprogetti",
  "reply_to_id": "<message_id>",
  "cc": ["first@example.com", "second@example.com"],
  "body": "Reply text"
}
```

- A reply carrying any of those fields goes through `createreply` / `createreplyall`, a `PATCH` on the draft, then `send`. The draft is deleted if a step fails. A plain reply uses the direct `reply` / `replyall` action.
- `cc` and `bcc` are added to the recipients of the reply, duplicates skipped.
- `to` replaces the default recipient. It is required to reply to your own sent message: without it, the reply goes back to you.
- The tool result lists the final To, Cc and Bcc.
- `is_html` has no effect on replies: `body` is sent as the reply comment.

After changing the code, run `install.ps1` and restart the MCP client: a running server keeps the code it started with.

