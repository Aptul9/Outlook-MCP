"""Outlook MCP: Model Context Protocol server over New Outlook messages.

Standard library only, zero external Python packages.
Harvests decrypted tokens straight from IdentityCache / OneAuth via Windows DPAPI.
No elevation required, no app restarts, runs transparently.
"""
import datetime as dt
import html
import json
import os
import re
import sys
import threading
import urllib.parse

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import outlook_api as api
import outlook_tokens as tokens

PROTOCOL_VERSION = '2025-06-18'
SERVER_VERSION = '1.0.0'

# ---------------------------------------------------------------- Formatting helpers

def fmt_time(iso_str):
    if not iso_str:
        return '-'
    try:
        # e.g. 2026-10-29T08:57:00Z
        clean = iso_str.replace('Z', '+00:00')
        t = dt.datetime.fromisoformat(clean)
        loc = t.astimezone()
        return loc.strftime('%Y-%m-%d %H:%M')
    except Exception:
        return iso_str[:16].replace('T', ' ')


def recipient_str(rec_obj):
    if isinstance(rec_obj, dict):
        addr_info = rec_obj.get('EmailAddress') or rec_obj
        name = addr_info.get('Name')
        addr = addr_info.get('Address')
        if name and addr and name != addr:
            return f"{name} <{addr}>"
        return addr or name or '?'
    return str(rec_obj)


def recipients_list_str(recs):
    if not recs:
        return '-'
    return ', '.join(recipient_str(r) for r in recs)


def format_message_summary(m):
    mid = m.get('Id')
    cid = m.get('ConversationId')
    dt_str = fmt_time(m.get('ReceivedDateTime') or m.get('CreatedDateTime'))
    sender = recipient_str(m.get('From') or m.get('Sender'))
    subj = m.get('Subject') or '(No Subject)'
    prev = m.get('BodyPreview') or ''
    unread = ' [UNREAD]' if m.get('IsRead') is False else ''
    att = ' [📎]' if m.get('HasAttachments') else ''
    
    lines = [
        f"[{dt_str}]{unread}{att} {sender}",
        f"Subject: {subj}",
        f"Preview: {prev[:120]}..." if len(prev) > 120 else f"Preview: {prev}",
        f"Id: {mid}",
        f"ThreadId: {cid}"
    ]
    return '\n'.join(lines)


def format_message_full(m):
    mid = m.get('Id')
    cid = m.get('ConversationId')
    dt_str = fmt_time(m.get('ReceivedDateTime') or m.get('CreatedDateTime'))
    sender = recipient_str(m.get('From') or m.get('Sender'))
    to_str = recipients_list_str(m.get('ToRecipients'))
    cc_str = recipients_list_str(m.get('CcRecipients'))
    subj = m.get('Subject') or '(No Subject)'
    
    raw_body = (m.get('Body') or {}).get('Content') or m.get('BodyPreview') or ''
    body_md = api.html_to_markdown(raw_body)
    
    header = [
        f"==================================================",
        f"From: {sender}",
        f"To: {to_str}",
    ]
    if m.get('CcRecipients'):
        header.append(f"Cc: {cc_str}")
    header.append(f"Date: {dt_str}")
    header.append(f"Subject: {subj}")
    header.append(f"Message ID: {mid}")
    header.append(f"Thread ID: {cid}")
    
    atts = m.get('AttachmentsList', [])
    if atts:
        header.append("Attachments:")
        for a in atts:
            size_kb = (a.get('Size') or 0) // 1024
            header.append(f"  📎 {a.get('Name')} ({size_kb} KB, {a.get('ContentType')})")
    header.append("--------------------------------------------------")
    
    return '\n'.join(header) + '\n\n' + body_md


# ---------------------------------------------------------------- Tool implementations

def tool_list_accounts(args):
    accts = api.list_accounts()
    if not accts:
        return "No Outlook accounts found. Ensure New Outlook is opened and signed in."
    lines = [f"Found {len(accts)} active Outlook account(s):", ""]
    for a in accts:
        lines.append(f"- {a['account']} ({a['endpoint_type']}, token expires in {a['expires_in_minutes']} min)")
    return '\n'.join(lines)


def tool_list_folders(args):
    account = args.get('account')
    folders = api.list_folders(account)
    info = tokens.get_token_for_account(account)
    lines = [f"Folders for {info['upn']}:", ""]
    for f in folders:
        unread = f" ({f['unread_count']} unread)" if f['unread_count'] > 0 else ""
        lines.append(f"- {f['name']} [{f['total_count']} items{unread}] (id: {f['id']})")
    return '\n'.join(lines)


def tool_list_emails(args):
    account = args.get('account')
    folder = args.get('folder') or 'Inbox'
    query = args.get('query')
    limit = int(args.get('limit') or 20)
    unread_only = bool(args.get('unread_only', False))
    
    info = tokens.get_token_for_account(account)
    emails = api.list_emails(account=account, folder=folder, query=query, limit=limit, unread_only=unread_only)
    
    if not emails:
        return f"No emails found in '{folder}' for {info['upn']} matching criteria."
    
    out = [
        f"Account: {info['upn']} | Folder: {folder} | Showing {len(emails)} messages (newest first)",
        ""
    ]
    out.append('\n\n'.join(format_message_summary(m) for m in emails))
    return '\n'.join(out)


def tool_read_email(args):
    account = args.get('account')
    message_id = args.get('message_id')
    if not message_id:
        return "message_id is required"
    msg = api.get_message(account=account, message_id=message_id)
    return format_message_full(msg)


def tool_read_thread(args):
    account = args.get('account')
    thread_id = args.get('thread_id')
    if not thread_id:
        return "thread_id (or message_id) is required"
    limit = int(args.get('limit') or 50)
    
    msgs = api.read_thread(account=account, thread_id=thread_id, limit=limit)
    if not msgs:
        return f"No messages found for thread '{thread_id}'"
    
    info = tokens.get_token_for_account(account)
    subj = msgs[0].get('Subject') or '(No Subject)'
    cid = msgs[0].get('ConversationId') or thread_id
    
    out = [
        f"Thread: {subj} ({len(msgs)} message(s), oldest to newest)",
        f"Account: {info['upn']} | ConversationId: {cid}",
        ""
    ]
    
    for i, m in enumerate(msgs, 1):
        out.append(f"--- [Message {i} of {len(msgs)}] ---")
        out.append(format_message_full(m))
        out.append("")
        
    return '\n'.join(out)


def tool_search_emails(args):
    account = args.get('account')
    query = args.get('query')
    if not query:
        return "query is required"
    limit = int(args.get('limit') or 20)
    
    info = tokens.get_token_for_account(account)
    msgs = api.search_emails(account=account, query=query, limit=limit)
    if not msgs:
        return f"No emails matching query '{query}' for {info['upn']}."
    
    out = [
        f"Search results for '{query}' ({len(msgs)} hits, account: {info['upn']}):",
        ""
    ]
    out.append('\n\n'.join(format_message_summary(m) for m in msgs))
    return '\n'.join(out)


def tool_send_email(args):
    account = args.get('account')
    to = args.get('to')
    subject = args.get('subject') or ''
    body = args.get('body') or ''
    cc = args.get('cc')
    bcc = args.get('bcc')
    is_html = bool(args.get('is_html', False))
    reply_to_id = args.get('reply_to_id')
    reply_all = bool(args.get('reply_all', False))
    attachments = args.get('attachments')
    
    if reply_to_id:
        if not body.strip():
            return "body comment is required when replying"
        res = api.send_email(account=account, to=to, body=body, cc=cc, bcc=bcc, reply_to_id=reply_to_id,
                             reply_all=reply_all, attachments=attachments)
        sent = f" To: {res['to']}. Cc: {res['cc']}. Bcc: {res['bcc']}." if 'to' in res else ''
        return f"Reply successfully sent to message {reply_to_id} ({'reply all' if reply_all else 'single reply'}).{sent}"
    
    if not to:
        return "'to' recipient is required for sending new email"
    if not body.strip() and not subject.strip():
        return "Cannot send empty email (both subject and body are empty)"
    
    res = api.send_email(account=account, to=to, subject=subject, body=body,
                         cc=cc, bcc=bcc, is_html=is_html, attachments=attachments)
    to_str = ', '.join(res.get('to', []))
    info = tokens.get_token_for_account(account)
    return f"Email '{subject}' successfully sent from {info['upn']} to [{to_str}]."


# ---------------------------------------------------------------- Tool Registry

ACCOUNT_ARG = {'type': 'string', 'description': 'Filter or select by account email (e.g. part of address). If omitted, uses default account.'}

TOOLS = [
    {
        'name': 'list_accounts',
        'description': 'List all signed-in Outlook accounts detected on the system and their token status.',
        'inputSchema': {'type': 'object', 'properties': {}},
        'fn': tool_list_accounts
    },
    {
        'name': 'list_folders',
        'description': 'List mail folders (Inbox, Sent, Archive, custom folders, etc.) with unread and total message counts.',
        'inputSchema': {'type': 'object', 'properties': {
            'account': ACCOUNT_ARG
        }},
        'fn': tool_list_folders
    },
    {
        'name': 'list_emails',
        'description': 'List emails in a folder (default Inbox) sorted newest first. Returns summaries with message ID, thread ID, sender, subject, and preview.',
        'inputSchema': {'type': 'object', 'properties': {
            'account': ACCOUNT_ARG,
            'folder': {'type': 'string', 'description': 'Folder name (e.g. Inbox, SentItems, Archive, Drafts) or folder ID. Default: Inbox.'},
            'query': {'type': 'string', 'description': 'Optional search keyword to filter emails.'},
            'limit': {'type': 'integer', 'description': 'Max emails to return (default 20).'},
            'unread_only': {'type': 'boolean', 'description': 'If true, only return unread emails.'}
        }},
        'fn': tool_list_emails
    },
    {
        'name': 'read_email',
        'description': 'Read full content of a specific email by its message ID, including full body text (clean Markdown) and attachments list.',
        'inputSchema': {'type': 'object', 'required': ['message_id'], 'properties': {
            'message_id': {'type': 'string', 'description': 'The message ID from list_emails or search_emails.'},
            'account': ACCOUNT_ARG
        }},
        'fn': tool_read_email
    },
    {
        'name': 'read_thread',
        'description': 'Read an entire email conversation thread by its thread_id (ConversationId) or any message_id within the thread. Messages are returned in chronological order (oldest to newest) with full body content and attachments.',
        'inputSchema': {'type': 'object', 'required': ['thread_id'], 'properties': {
            'thread_id': {'type': 'string', 'description': 'ConversationId (from list_emails/search_emails) or any message ID from the thread.'},
            'account': ACCOUNT_ARG,
            'limit': {'type': 'integer', 'description': 'Max messages in thread to return (default 50).'}
        }},
        'fn': tool_read_thread
    },
    {
        'name': 'search_emails',
        'description': 'Search emails across all folders by keyword, sender, or subject.',
        'inputSchema': {'type': 'object', 'required': ['query'], 'properties': {
            'query': {'type': 'string', 'description': 'Search query string.'},
            'account': ACCOUNT_ARG,
            'limit': {'type': 'integer', 'description': 'Max results (default 20).'}
        }},
        'fn': tool_search_emails
    },
    {
        'name': 'send_email',
        'description': 'Send a new email or reply to an existing email / conversation thread.',
        'inputSchema': {'type': 'object', 'properties': {
            'account': ACCOUNT_ARG,
            'to': {'description': 'Recipient email address or list of addresses (required if not replying; when replying, replaces the default reply recipient).', 'type': ['string', 'array']},
            'subject': {'type': 'string', 'description': 'Subject of the email.'},
            'body': {'type': 'string', 'description': 'Text or HTML body content.'},
            'cc': {'description': 'CC recipient(s).', 'type': ['string', 'array']},
            'bcc': {'description': 'BCC recipient(s).', 'type': ['string', 'array']},
            'is_html': {'type': 'boolean', 'description': 'True if body is HTML, false for plain text (default false).'},
            'reply_to_id': {'type': 'string', 'description': 'Message ID to reply to. When specified, creates and sends a proper reply in that conversation thread.'},
            'reply_all': {'type': 'boolean', 'description': 'If replying, whether to reply-all to all recipients (default false).'},
            'attachments': {'type': 'array', 'items': {'type': 'string'}, 'description': 'List of local file paths to attach to the email.'}
        }},
        'fn': tool_send_email
    }
]

# ---------------------------------------------------------------- MCP JSON-RPC Server

def reply(msg_id, result=None, error=None):
    out = {'jsonrpc': '2.0', 'id': msg_id}
    if error:
        out['error'] = error
    else:
        out['result'] = result
    sys.stdout.write(json.dumps(out, ensure_ascii=False) + '\n')
    sys.stdout.flush()


def handle(req):
    method = req.get('method')
    msg_id = req.get('id')
    if msg_id is None:
        return  # notification
    if method == 'initialize':
        version = (req.get('params') or {}).get('protocolVersion') or PROTOCOL_VERSION
        reply(msg_id, {
            'protocolVersion': version,
            'capabilities': {'tools': {}},
            'serverInfo': {'name': 'outlook', 'version': SERVER_VERSION},
            'instructions': 'Access to New Outlook emails and threads. Read inbox, folders, navigate conversation threads, search, and send/reply to emails.'
        })
    elif method == 'ping':
        reply(msg_id, {})
    elif method == 'tools/list':
        reply(msg_id, {'tools': [{k: v for k, v in t.items() if k != 'fn'} for t in TOOLS]})
    elif method == 'tools/call':
        p = req.get('params') or {}
        tool_name = p.get('name')
        tool = next((t for t in TOOLS if t['name'] == tool_name), None)
        if not tool:
            reply(msg_id, error={'code': -32602, 'message': f'unknown tool {tool_name}'})
            return
        try:
            res_text = tool['fn'](p.get('arguments') or {})
            reply(msg_id, {'content': [{'type': 'text', 'text': res_text}]})
        except Exception as e:
            reply(msg_id, {'content': [{'type': 'text', 'text': f'Error: {type(e).__name__}: {e}'}], 'isError': True})
    else:
        reply(msg_id, error={'code': -32601, 'message': f'method not found: {method}'})


def main():
    sys.stdin.reconfigure(encoding='utf-8')
    sys.stdout.reconfigure(encoding='utf-8', newline='\n')
    
    if len(sys.argv) > 2 and sys.argv[1] == '--call':
        tool_name = sys.argv[2]
        tool = next((t for t in TOOLS if t['name'] == tool_name), None)
        if not tool:
            print(f"Tool {tool_name} not found")
            return
        raw_arg = ' '.join(sys.argv[3:]) if len(sys.argv) > 3 else '{}'
        try:
            args = json.loads(raw_arg)
        except Exception:
            # Fallback if quotes got stripped by Windows command line
            try:
                import ast
                args = ast.literal_eval(raw_arg)
            except Exception:
                args = {}
        print(tool['fn'](args))
        return
        
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            req = json.loads(line)
        except ValueError:
            continue
        handle(req)


if __name__ == '__main__':
    main()
