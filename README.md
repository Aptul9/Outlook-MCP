# Outlook MCP

Full-featured MCP (Model Context Protocol) server for **OE Copilot** (virtual Android phone on remote server) and **New Outlook** (`olk.exe` / `Microsoft.OutlookForWindows`).

Reads emails, navigates entire conversation threads, searches messages, downloads and expands attachments (including embedded Outlook emails), manages folders/flags/read status, and sends, replies to, or forwards emails with rich HTML formatting, selective recipients, multiple CC/BCC, and binary attachments.

- **Zero third-party Python dependencies on Windows** (standard library only: `ctypes`, `urllib`, `sqlite3`, `subprocess`, `json`).
- **Dual token harvester (`outlook_tokens.py`)**:
  1. **OE Copilot (`oecopilot-vm`)**: harvests and decrypts `HxCoreSecureStringAES256Gcm` tokens from `acompliAcct.db` and Android Keystore `persistent.sqlite` on the virtual phone over SSH (`OUTLOOK_MCP_SSH_HOST`, default `oecopilot-vm`), plus phone sync status from `stato/accessi.json`.
  2. **Local New Outlook (`IdentityCache`)**: harvests tokens from `%LOCALAPPDATA%\Microsoft\IdentityCache` via Windows DPAPI.
  - Tokens are cached locally in `%LOCALAPPDATA%\OutlookMCP\tokens.bin` protected by Windows DPAPI.

---

## Tools

| Tool | Description | Key Arguments |
|---|---|---|
| `list_accounts` | Lists all detected Outlook accounts (OE Copilot + local), token expiry, and phone sync status | None |
| `list_folders` | Lists mail folders (including nested child folders) with total and unread counts | `account`, `include_child_folders` |
| `list_emails` | Lists messages newest first with IDs, thread ID, From, To, Cc, subject, preview | `folder`, `query`, `from_address`, `unread_only`, `has_attachments`, `flagged_only`, `start_date`, `end_date`, `limit`, `account` |
| `read_email` | Reads one email completely: headers, clean Markdown body (including tables), attachments list | `message_id`, `account` |
| `read_thread` | **Reads an entire conversation thread** in chronological order (oldest to newest) with participants list, full Markdown bodies, and attachments | `thread_id` (or any `message_id`), `limit`, `account` |
| `search_emails` | Full-text search across all folders (or a specific folder) with sender, date, unread, and attachment filters | `query`, `from_address`, `folder`, `unread_only`, `has_attachments`, `start_date`, `end_date`, `limit`, `account` |
| `send_email` | Send a new email, reply/reply-all/selective-reply to a thread, forward an email, or save as draft | `to`, `subject`, `body`, `cc`, `bcc`, `is_html`, `reply_to_id`, `reply_all`, `replace_cc`, `exclude_recipients`, `forward_to_id`, `attachments`, `importance`, `save_as_draft`, `account` |
| `forward_email` | Forward an existing email preserving original server-side attachments and quoted history | `message_id`, `to`, `body`, `cc`, `bcc`, `subject`, `is_html`, `attachments`, `save_as_draft`, `account` |
| `get_attachment` | Download attachment(s) to disk, return images inline, and expand embedded Outlook emails (`ItemAttachment`) into Markdown | `message_id`, `attachment_id`, `index`, `save_dir`, `account` |
| `manage_email` | Mark read/unread, flag/unflag, archive, delete, or move message(s) to a folder | `message_id` (string or list), `action`, `destination_folder`, `account` |

---

## Install

```powershell
cd C:\Users\Samuele\Downloads\outlook-mcp
powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -NoClaudeCode
```

---

## Replying with Rich HTML, Multiple CC, and Selective Recipients

`send_email` supports full rich HTML (`is_html: true` or auto-detected HTML tags) on **both new emails and replies/forwards**, as well as complete recipient control on replies:

```json
{
  "account": "prisma",
  "reply_to_id": "<message_id>",
  "reply_all": true,
  "to": ["m.svaicari@prismaprogetti.it"],
  "cc": ["c.ragno@prismaprogetti.it", "d.fortunio@prismaprogetti.it"],
  "exclude_recipients": ["someone@example.com"],
  "is_html": true,
  "body": "<p>Ciao <b>Marco</b>,</p><ul><li>Punto 1</li><li>Punto 2</li></ul>",
  "attachments": ["C:\\path\\to\\file.pdf"]
}
```

- **Selective reply to specific people**: pass `to` (single address or list) to override `ToRecipients`. With `reply_all: false` (default), only the specified `to` and `cc` recipients receive the reply while keeping the thread (`ConversationId`, `In-Reply-To`, and quoted history) intact.
- **Reply-all + extra CC**: pass `reply_all: true` and `cc` to keep existing thread CCs and merge additional CC addresses (or pass `replace_cc: true` to replace the CC list).
- **Exclude specific people on reply-all**: pass `exclude_recipients` to strip specific addresses from `To` and `Cc`.
- **Rich HTML or plain text**: when `is_html: true` (or when `body` contains HTML tags), rich HTML is inserted directly above the quoted thread. When `is_html: false` (plain text), line breaks (`\n`) are automatically converted to `<br>` so paragraphs are never collapsed.
- **Draft mode**: pass `save_as_draft: true` to create the message/reply/forward in `Drafts` without sending immediately.
