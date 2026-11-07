"""Outlook MCP: Model Context Protocol server over OE Copilot & New Outlook messages.

Standard library only, zero external Python packages on Windows.
Harvests decrypted tokens transparently from:
- OE Copilot virtual Android phone over SSH (oecopilot-vm)
- Local New Outlook IdentityCache via Windows DPAPI
"""
import base64
import datetime as dt
import html
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import outlook_api as api
import outlook_tokens as tokens

PROTOCOL_VERSION = '2025-06-18'
SERVER_VERSION = '2.0.0'

DOWNLOAD_DIR = os.path.join(os.environ.get('LOCALAPPDATA', '.'), 'OutlookMCP', 'downloads')

# ---------------------------------------------------------------- Formatting helpers

def fmt_time(iso_str):
    if not iso_str:
        return '-'
    try:
        clean = iso_str.replace('Z', '+00:00')
        t = dt.datetime.fromisoformat(clean)
        loc = t.astimezone()
        return loc.strftime('%Y-%m-%d %H:%M')
    except Exception:
        return iso_str[:16].replace('T', ' ')


def _field(obj, name):
    """Read a field case-insensitively (Outlook REST is PascalCase, Graph camelCase)."""
    if not isinstance(obj, dict):
        return None
    if name in obj:
        return obj[name]
    cam = name[0].lower() + name[1:]
    if cam in obj:
        return obj[cam]
    lower = name.lower()
    for k, v in obj.items():
        if k.lower() == lower:
            return v
    return None


def recipient_str(rec_obj):
    if isinstance(rec_obj, dict):
        addr_info = _field(rec_obj, 'EmailAddress') or rec_obj
        name = _field(addr_info, 'Name')
        addr = _field(addr_info, 'Address')
        if name and addr and name != addr:
            return f"{name} <{addr}>"
        return addr or name or '?'
    return str(rec_obj)


def recipients_list_str(recs):
    if not recs:
        return '-'
    return ', '.join(recipient_str(r) for r in recs)


def format_message_summary(m):
    mid = _field(m, 'Id')
    cid = _field(m, 'ConversationId')
    dt_str = fmt_time(_field(m, 'ReceivedDateTime') or _field(m, 'CreatedDateTime'))
    sender = recipient_str(_field(m, 'From') or _field(m, 'Sender'))
    to_recs = _field(m, 'ToRecipients') or []
    cc_recs = _field(m, 'CcRecipients') or []
    subj = _field(m, 'Subject') or '(No Subject)'
    prev = (_field(m, 'BodyPreview') or '').replace('\r', ' ').replace('\n', ' ').strip()
    unread = ' [UNREAD]' if _field(m, 'IsRead') is False else ''
    draft = ' [DRAFT]' if _field(m, 'IsDraft') is True else ''
    att = ' [📎]' if _field(m, 'HasAttachments') else ''
    flag_obj = _field(m, 'Flag') or {}
    flagged = ' [🚩]' if (_field(flag_obj, 'FlagStatus') or '').lower() == 'flagged' else ''
    imp = _field(m, 'Importance') or ''
    high_imp = ' [❗HIGH]' if imp.lower() == 'high' else ''

    lines = [
        f"[{dt_str}]{unread}{draft}{att}{flagged}{high_imp} From: {sender}",
    ]
    if to_recs:
        lines.append(f"To: {recipients_list_str(to_recs)}")
    if cc_recs:
        lines.append(f"Cc: {recipients_list_str(cc_recs)}")
    lines.extend([
        f"Subject: {subj}",
        f"Preview: {prev[:140]}..." if len(prev) > 140 else f"Preview: {prev}",
        f"Id: {mid}",
        f"ThreadId: {cid}"
    ])
    return '\n'.join(lines)


def format_message_full(m):
    mid = _field(m, 'Id')
    cid = _field(m, 'ConversationId')
    dt_str = fmt_time(_field(m, 'ReceivedDateTime') or _field(m, 'SentDateTime') or _field(m, 'CreatedDateTime'))
    sender = recipient_str(_field(m, 'From') or _field(m, 'Sender'))
    to_str = recipients_list_str(_field(m, 'ToRecipients'))
    cc_str = recipients_list_str(_field(m, 'CcRecipients'))
    bcc_str = recipients_list_str(_field(m, 'BccRecipients'))
    reply_to = _field(m, 'ReplyTo')
    subj = _field(m, 'Subject') or '(No Subject)'
    is_draft = _field(m, 'IsDraft') is True
    is_read = _field(m, 'IsRead')
    imp = _field(m, 'Importance') or 'Normal'

    body_obj = _field(m, 'Body') or {}
    raw_body = _field(body_obj, 'Content') or _field(m, 'BodyPreview') or ''
    body_type = (_field(body_obj, 'ContentType') or 'HTML').upper()
    body_md = api.html_to_markdown(raw_body) if body_type == 'HTML' else raw_body.strip()

    status_tags = []
    if is_draft:
        status_tags.append('DRAFT')
    if is_read is False:
        status_tags.append('UNREAD')
    if imp.lower() != 'normal':
        status_tags.append(f'Importance: {imp}')

    header = [
        "==================================================",
        f"From: {sender}",
        f"To: {to_str}",
    ]
    if _field(m, 'CcRecipients'):
        header.append(f"Cc: {cc_str}")
    if _field(m, 'BccRecipients'):
        header.append(f"Bcc: {bcc_str}")
    if reply_to:
        header.append(f"Reply-To: {recipients_list_str(reply_to)}")
    header.append(f"Date: {dt_str}")
    if status_tags:
        header.append(f"Status: {', '.join(status_tags)}")
    header.append(f"Subject: {subj}")
    header.append(f"Message ID: {mid}")
    header.append(f"Thread ID: {cid}")

    atts = m.get('AttachmentsList') or []
    if atts:
        header.append("Attachments:")
        for idx, a in enumerate(atts, 1):
            size_kb = max(1, (_field(a, 'Size') or 0) // 1024) if (_field(a, 'Size') or 0) else 0
            inline_tag = ', inline' if _field(a, 'IsInline') else ''
            aid = _field(a, 'Id') or ''
            header.append(
                f"  #{idx} 📎 {_field(a, 'Name')} ({size_kb} KB, {_field(a, 'ContentType')}{inline_tag}) [attachment_id: {aid}]"
            )
    header.append("--------------------------------------------------")

    return '\n'.join(header) + '\n\n' + body_md


# ---------------------------------------------------------------- Tool implementations

def tool_list_accounts(args):
    accts = api.list_accounts()
    if not accts:
        return "No Outlook accounts found (checked OE Copilot virtual phone and local IdentityCache)."
    lines = [f"Found {len(accts)} active Outlook account(s):", ""]
    for a in accts:
        src = a.get('source', 'local')
        pstat = f", phone sync: {a['phone_status']}" if a.get('phone_status') else ""
        pmsg = f" ({a['phone_msg']})" if a.get('phone_msg') else ""
        lines.append(
            f"- {a['account']} [{src}, {a['endpoint_type']}, token expires in {a['expires_in_minutes']} min{pstat}]{pmsg}"
        )
    return '\n'.join(lines)


def tool_list_folders(args):
    account = args.get('account')
    include_children = bool(args.get('include_child_folders', True))
    folders = api.list_folders(account, include_child_folders=include_children)
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
    from_address = args.get('from_address')
    has_attachments = args.get('has_attachments')
    if has_attachments is not None:
        has_attachments = bool(has_attachments)
    start_date = args.get('start_date')
    end_date = args.get('end_date')
    flagged_only = bool(args.get('flagged_only', False))

    info = tokens.get_token_for_account(account)
    emails = api.list_emails(
        account=account,
        folder=folder,
        query=query,
        limit=limit,
        unread_only=unread_only,
        from_address=from_address,
        has_attachments=has_attachments,
        start_date=start_date,
        end_date=end_date,
        flagged_only=flagged_only
    )

    if not emails:
        return f"No emails found in '{folder}' for {info['upn']} matching criteria."

    out = [
        f"Account: {info['upn']} | Folder: {folder} | Showing {len(emails)} message(s) (newest first)",
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
    subj = _field(msgs[0], 'Subject') or '(No Subject)'
    cid = _field(msgs[0], 'ConversationId') or thread_id

    # Collect unique participants across the thread
    participants = []
    seen_p = set()
    for m in msgs:
        for r in ([_field(m, 'From') or _field(m, 'Sender')]
                  + (_field(m, 'ToRecipients') or [])
                  + (_field(m, 'CcRecipients') or [])):
            if not r:
                continue
            s = recipient_str(r)
            k = s.casefold()
            if s != '?' and k not in seen_p:
                seen_p.add(k)
                participants.append(s)

    out = [
        f"Thread: {subj} ({len(msgs)} message(s), oldest to newest)",
        f"Account: {info['upn']} | ConversationId: {cid}",
        f"Participants: {', '.join(participants)}",
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
    from_address = args.get('from_address')
    if not query and not from_address:
        return "query or from_address is required"
    folder = args.get('folder')
    unread_only = bool(args.get('unread_only', False))
    has_attachments = args.get('has_attachments')
    if has_attachments is not None:
        has_attachments = bool(has_attachments)
    start_date = args.get('start_date')
    end_date = args.get('end_date')
    limit = int(args.get('limit') or 20)

    info = tokens.get_token_for_account(account)
    msgs = api.search_emails(
        account=account,
        query=query,
        folder=folder,
        from_address=from_address,
        unread_only=unread_only,
        has_attachments=has_attachments,
        start_date=start_date,
        end_date=end_date,
        limit=limit
    )
    desc = query or f"from:{from_address}"
    if not msgs:
        return f"No emails matching '{desc}' for {info['upn']}."

    out = [
        f"Search results for '{desc}' ({len(msgs)} hit(s), account: {info['upn']}):",
        ""
    ]
    out.append('\n\n'.join(format_message_summary(m) for m in msgs))
    return '\n'.join(out)


def _format_send_result(info, res, mode_label):
    status = res.get('status', 'sent')
    to_s = ', '.join(res.get('to') or []) or '-'
    cc_s = ', '.join(res.get('cc') or [])
    bcc_s = ', '.join(res.get('bcc') or [])
    subj = res.get('subject') or ''
    parts = [
        f"{mode_label} {'saved as draft' if status == 'draft' else 'successfully sent'} from {info['upn']}."
    ]
    if subj:
        parts.append(f"Subject: {subj}")
    parts.append(f"To: [{to_s}]")
    if cc_s:
        parts.append(f"Cc: [{cc_s}]")
    if bcc_s:
        parts.append(f"Bcc: [{bcc_s}]")
    if res.get('draft_id'):
        parts.append(f"Draft ID: {res['draft_id']}")
    return ' | '.join(parts)


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
    forward_to_id = args.get('forward_to_id')
    replace_cc = bool(args.get('replace_cc', False))
    exclude_recipients = args.get('exclude_recipients')
    importance = args.get('importance')
    attachments = args.get('attachments')
    save_as_draft = bool(args.get('save_as_draft', False))

    info = tokens.get_token_for_account(account)

    if reply_to_id:
        if not body.strip() and not save_as_draft:
            return "body comment is required when replying"
        res = api.send_email(
            account=account,
            to=to,
            subject=subject,
            body=body,
            cc=cc,
            bcc=bcc,
            is_html=is_html,
            reply_to_id=reply_to_id,
            reply_all=reply_all,
            replace_cc=replace_cc,
            exclude_recipients=exclude_recipients,
            importance=importance,
            attachments=attachments,
            save_as_draft=save_as_draft
        )
        label = f"Reply ({'reply-all' if reply_all else 'selective/single reply'} to {reply_to_id[:24]}...)"
        return _format_send_result(info, res, label)

    if forward_to_id:
        if not to and not save_as_draft:
            return "'to' recipient is required when forwarding an email"
        res = api.send_email(
            account=account,
            to=to,
            subject=subject,
            body=body,
            cc=cc,
            bcc=bcc,
            is_html=is_html,
            forward_to_id=forward_to_id,
            exclude_recipients=exclude_recipients,
            importance=importance,
            attachments=attachments,
            save_as_draft=save_as_draft
        )
        label = f"Forward (of {forward_to_id[:24]}...)"
        return _format_send_result(info, res, label)

    if not to and not save_as_draft:
        return "'to' recipient is required for sending new email"
    if not body.strip() and not subject.strip():
        return "Cannot send empty email (both subject and body are empty)"

    res = api.send_email(
        account=account,
        to=to,
        subject=subject,
        body=body,
        cc=cc,
        bcc=bcc,
        is_html=is_html,
        importance=importance,
        attachments=attachments,
        save_as_draft=save_as_draft
    )
    return _format_send_result(info, res, "Email")


def tool_forward_email(args):
    forward_args = dict(args)
    forward_args['forward_to_id'] = args.get('message_id') or args.get('forward_to_id')
    if not forward_args.get('forward_to_id'):
        return "message_id is required to forward an email"
    return tool_send_email(forward_args)


def _safe_filename(name, fallback):
    name = (name or fallback).strip()
    name = re.sub(r'[\\/:*?"<>|\r\n\t]', '_', name)
    name = name.strip('. ') or fallback
    return name[:150]


def tool_get_attachment(args):
    account = args.get('account')
    message_id = args.get('message_id')
    if not message_id:
        return 'message_id is required'
    atts = api.list_attachments(account=account, message_id=message_id)
    if not atts:
        return f'No attachments on message {message_id}.'
    att_id = args.get('attachment_id')
    idx = args.get('index')
    if att_id:
        picked = [(i, a) for i, a in enumerate(atts, 1) if _field(a, 'Id') == att_id]
        if not picked:
            return f'attachment_id not found on message {message_id}; it has {len(atts)} attachment(s).'
    elif idx is not None:
        i = int(idx)
        if not 1 <= i <= len(atts):
            return f'index {i} out of range: message {message_id} has {len(atts)} attachment(s).'
        picked = [(i, atts[i - 1])]
    else:
        picked = list(enumerate(atts, 1))

    save_dir = args.get('save_dir') or DOWNLOAD_DIR
    os.makedirs(save_dir, exist_ok=True)
    blocks, lines = [], [f'{len(atts)} attachment(s) on message {message_id}:']
    for n, meta in picked:
        aid = _field(meta, 'Id')
        full = api.get_attachment(account=account, message_id=message_id, attachment_id=aid)
        otype = (full.get('@odata.type') or meta.get('@odata.type') or '').lower()
        name = _field(full, 'Name') or _field(meta, 'Name') or f'attachment-{n}'
        mime = (_field(full, 'ContentType') or _field(meta, 'ContentType') or 'application/octet-stream').split(';')[0].strip()

        if 'itemattachment' in otype:
            try:
                expanded = api.get_attachment(account=account, message_id=message_id, attachment_id=aid, expand_item=True)
                item = _field(expanded, 'Item') or {}
            except Exception:
                item = {}
            if item:
                formatted_item = format_message_full(item)
                md_path = os.path.join(save_dir, _safe_filename(name, f'item-{n}') + '.md')
                with open(md_path, 'w', encoding='utf-8') as fh:
                    fh.write(formatted_item)
                lines.append(f'  #{n} "{name}": embedded Outlook item saved as Markdown to {md_path}:')
                lines.append(formatted_item)
            else:
                lines.append(f'  #{n} "{name}": embedded Outlook item (email/event).')
            continue

        if 'referenceattachment' in otype:
            url = _field(full, 'SourceUrl') or ''
            lines.append(f'  #{n} "{name}": link to a cloud file (OneDrive/SharePoint), no bytes on the message. {url}'.rstrip())
            continue

        content_b64 = _field(full, 'ContentBytes')
        if not content_b64:
            lines.append(f'  #{n} "{name}": no content bytes returned (type {otype or "unknown"}).')
            continue
        data = base64.b64decode(content_b64)
        path = os.path.join(save_dir, _safe_filename(name, f'attachment-{n}'))
        with open(path, 'wb') as fh:
            fh.write(data)
        lines.append(f'  #{n} "{name}" saved: {path} ({len(data) / 1024:.1f} KB, {mime})')
        if mime.startswith('image/'):
            blocks.append({'type': 'image', 'data': content_b64, 'mimeType': mime})
    return [{'type': 'text', 'text': '\n'.join(lines)}] + blocks


def tool_manage_email(args):
    account = args.get('account')
    message_id = args.get('message_id')
    action = args.get('action')
    destination_folder = args.get('destination_folder')
    if not message_id or not action:
        return "message_id and action are required"
    info = tokens.get_token_for_account(account)
    res = api.manage_email(
        account=account,
        message_id=message_id,
        action=action,
        destination_folder=destination_folder
    )
    lines = [f"Executed '{action}' on {len(res)} message(s) for {info['upn']}:"]
    for r in res:
        lines.append(f"- {r['id'][:32]}...: {r['status']}")
    return '\n'.join(lines)


# ---------------------------------------------------------------- Tool Registry

ACCOUNT_ARG = {
    'type': 'string',
    'description': 'Filter or select by account email (e.g. part of address like "roar", "serco", "prisma", "bandpey", "notarify"). If omitted, uses default account.'
}

TOOLS = [
    {
        'name': 'list_calendars',
        'description': 'Read calendars using the same account authorization as mail.',
        'inputSchema': {'type': 'object', 'properties': {'account': ACCOUNT_ARG}},
        'fn': lambda args: json.dumps(api.list_calendars(args.get('account')), ensure_ascii=False)
    },
    {
        'name': 'list_calendar_events',
        'description': 'Read a bounded calendar view. Reports has_more when truncated. No RSVP or state changes.',
        'inputSchema': {'type': 'object', 'required': ['start', 'end'], 'properties': {
            'account': ACCOUNT_ARG, 'start': {'type': 'string'}, 'end': {'type': 'string'},
            'calendar_id': {'type': 'string'}, 'limit': {'type': 'integer', 'minimum': 1, 'maximum': 1000}}},
        'fn': lambda args: json.dumps(api.list_calendar_events(args.get('account'), args['start'], args['end'], args.get('calendar_id'), args.get('limit', 100)), ensure_ascii=False)
    },
    {
        'name': 'list_accounts',
        'description': 'List all signed-in Outlook accounts detected on OE Copilot (virtual phone) and local PC, with their token and sync status.',
        'inputSchema': {'type': 'object', 'properties': {}},
        'fn': tool_list_accounts
    },
    {
        'name': 'list_folders',
        'description': 'List mail folders (Inbox, SentItems, Drafts, Archive, DeletedItems, and custom/child folders) with unread and total message counts.',
        'inputSchema': {'type': 'object', 'properties': {
            'account': ACCOUNT_ARG,
            'include_child_folders': {'type': 'boolean', 'description': 'Include nested subfolders (default true).'}
        }},
        'fn': tool_list_folders
    },
    {
        'name': 'list_emails',
        'description': 'List emails in a folder (default Inbox) sorted newest first. Returns summaries with message ID, thread ID, From, To, Cc, subject, and preview.',
        'inputSchema': {'type': 'object', 'properties': {
            'account': ACCOUNT_ARG,
            'folder': {'type': 'string', 'description': 'Folder name (e.g. Inbox, SentItems, Archive, Drafts, DeletedItems, custom folder, or "all") or folder ID. Default: Inbox.'},
            'query': {'type': 'string', 'description': 'Optional search keyword to filter emails.'},
            'from_address': {'type': 'string', 'description': 'Optional filter by sender email address or display name.'},
            'unread_only': {'type': 'boolean', 'description': 'If true, only return unread emails.'},
            'has_attachments': {'type': 'boolean', 'description': 'If true, only return emails with attachments.'},
            'flagged_only': {'type': 'boolean', 'description': 'If true, only return flagged emails.'},
            'start_date': {'type': 'string', 'description': 'Optional start date filter (ISO YYYY-MM-DD or YYYY-MM-DDTHH:MM:SSZ).'},
            'end_date': {'type': 'string', 'description': 'Optional end date filter (ISO YYYY-MM-DD or YYYY-MM-DDTHH:MM:SSZ).'},
            'limit': {'type': 'integer', 'description': 'Max emails to return (default 20).'}
        }},
        'fn': tool_list_emails
    },
    {
        'name': 'read_email',
        'description': 'Read full content of a specific email by its message ID, including From, To, Cc, Bcc, clean Markdown body (tables, links, formatting), and attachments list.',
        'inputSchema': {'type': 'object', 'required': ['message_id'], 'properties': {
            'message_id': {'type': 'string', 'description': 'The message ID from list_emails or search_emails.'},
            'account': ACCOUNT_ARG
        }},
        'fn': tool_read_email
    },
    {
        'name': 'read_thread',
        'description': 'Read an entire email conversation thread by its thread_id (ConversationId) or any message_id within the thread. Messages are returned in chronological order (oldest to newest) with full participants list, Markdown body content, and attachments.',
        'inputSchema': {'type': 'object', 'required': ['thread_id'], 'properties': {
            'thread_id': {'type': 'string', 'description': 'ConversationId (from list_emails/search_emails) or any message ID from the thread.'},
            'account': ACCOUNT_ARG,
            'limit': {'type': 'integer', 'description': 'Max messages in thread to return (default 50).'}
        }},
        'fn': tool_read_thread
    },
    {
        'name': 'search_emails',
        'description': 'Search emails across all folders (or a specific folder) by keyword, sender, date range, unread status, or attachment presence.',
        'inputSchema': {'type': 'object', 'properties': {
            'query': {'type': 'string', 'description': 'Search query string.'},
            'from_address': {'type': 'string', 'description': 'Filter by sender email or name.'},
            'folder': {'type': 'string', 'description': 'Optional folder name/ID to restrict search (defaults to all folders).'},
            'unread_only': {'type': 'boolean', 'description': 'If true, only return unread emails.'},
            'has_attachments': {'type': 'boolean', 'description': 'If true, only return emails with attachments.'},
            'start_date': {'type': 'string', 'description': 'Optional start date filter (YYYY-MM-DD).'},
            'end_date': {'type': 'string', 'description': 'Optional end date filter (YYYY-MM-DD).'},
            'account': ACCOUNT_ARG,
            'limit': {'type': 'integer', 'description': 'Max results (default 20).'}
        }},
        'fn': tool_search_emails
    },
    {
        'name': 'send_email',
        'description': (
            'Send a new email, reply to an existing email/thread, forward an email, or save as draft. '
            'Supports rich HTML formatting (is_html=true or HTML tags in body) on BOTH new emails and replies/forwards, '
            'multiple CC/BCC recipients, selective replies to specific people (via to, cc, replace_cc, or exclude_recipients), '
            'and local file attachments.'
        ),
        'inputSchema': {'type': 'object', 'properties': {
            'account': ACCOUNT_ARG,
            'to': {'description': 'Recipient email address or list of addresses. Required for new/forward emails; on replies, overrides ToRecipients so you can reply only to specific people.', 'type': ['string', 'array']},
            'subject': {'type': 'string', 'description': 'Subject of the email (optional on replies/forwards unless changing subject).'},
            'body': {'type': 'string', 'description': 'Text or rich HTML body content (tables, bold, badges, lists, highlights).'},
            'cc': {'description': 'CC recipient(s) as string (comma/semicolon separated) or list of addresses.', 'type': ['string', 'array']},
            'bcc': {'description': 'BCC recipient(s) as string or list of addresses.', 'type': ['string', 'array']},
            'is_html': {'type': 'boolean', 'description': 'True if body contains rich HTML formatting (works on new emails, replies, and forwards). Auto-detected if HTML tags are present.'},
            'reply_to_id': {'type': 'string', 'description': 'Message ID to reply to. Keeps thread history and ConversationId intact.'},
            'reply_all': {'type': 'boolean', 'description': 'If replying, true to start from all thread recipients (reply-all), false to start from sender only (default false).'},
            'replace_cc': {'type': 'boolean', 'description': 'When replying with reply_all=true, set replace_cc=true to replace the thread CC list with `cc` instead of merging.'},
            'exclude_recipients': {'description': 'Email address(es) to exclude from To/Cc when replying-all to a thread.', 'type': ['string', 'array']},
            'forward_to_id': {'type': 'string', 'description': 'Message ID to forward (preserves original attachments and quoted message).'},
            'importance': {'type': 'string', 'enum': ['Low', 'Normal', 'High'], 'description': 'Message importance level.'},
            'attachments': {'type': 'array', 'items': {'type': 'string'}, 'description': 'List of local file paths to attach to the email, reply, or forward.'},
            'save_as_draft': {'type': 'boolean', 'description': 'If true, saves the new email, reply, or forward in Drafts without sending immediately.'}
        }},
        'fn': tool_send_email
    },
    {
        'name': 'forward_email',
        'description': 'Forward an existing email (preserving its original server-side attachments and quoted thread) to one or more recipients, with optional rich HTML/text comment, CC/BCC, and extra local attachments.',
        'inputSchema': {'type': 'object', 'required': ['message_id', 'to'], 'properties': {
            'message_id': {'type': 'string', 'description': 'Message ID to forward.'},
            'to': {'description': 'Recipient email address or list of addresses.', 'type': ['string', 'array']},
            'body': {'type': 'string', 'description': 'Optional text or rich HTML comment added above the forwarded message.'},
            'cc': {'description': 'CC recipient(s).', 'type': ['string', 'array']},
            'bcc': {'description': 'BCC recipient(s).', 'type': ['string', 'array']},
            'subject': {'type': 'string', 'description': 'Optional custom subject override.'},
            'is_html': {'type': 'boolean', 'description': 'True if body is HTML.'},
            'attachments': {'type': 'array', 'items': {'type': 'string'}, 'description': 'Optional additional local file paths to attach.'},
            'save_as_draft': {'type': 'boolean', 'description': 'If true, saves the forward as a draft without sending.'},
            'account': ACCOUNT_ARG
        }},
        'fn': tool_forward_email
    },
    {
        'name': 'get_attachment',
        'description': 'Download attachment(s) of an email to disk, view image attachments inline, and expand embedded Outlook email/event attachments (ItemAttachment) into readable Markdown.',
        'inputSchema': {'type': 'object', 'required': ['message_id'], 'properties': {
            'message_id': {'type': 'string', 'description': 'The message ID holding the attachment(s).'},
            'attachment_id': {'type': 'string', 'description': 'Attachment Id (from the attachments list in read_email/read_thread). Optional.'},
            'index': {'type': 'integer', 'description': 'Which attachment to fetch, 1-based. Optional: all are fetched when omitted.'},
            'save_dir': {'type': 'string', 'description': 'Folder to save into. Optional, defaults to %LOCALAPPDATA%\\OutlookMCP\\downloads.'},
            'account': ACCOUNT_ARG
        }},
        'fn': tool_get_attachment
    },
    {
        'name': 'manage_email',
        'description': 'Mark email(s) as read/unread, flag/unflag, archive, delete (move to DeletedItems), or move to a specific folder.',
        'inputSchema': {'type': 'object', 'required': ['message_id', 'action'], 'properties': {
            'message_id': {'description': 'Message ID or list of message IDs to act on.', 'type': ['string', 'array']},
            'action': {
                'type': 'string',
                'enum': ['mark_read', 'mark_unread', 'flag', 'unflag', 'archive', 'delete', 'move'],
                'description': 'Action to perform on the message(s).'
            },
            'destination_folder': {'type': 'string', 'description': 'Target folder name or ID (required when action is "move").'},
            'account': ACCOUNT_ARG
        }},
        'fn': tool_manage_email
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
            'instructions': (
                'Access to OE Copilot and New Outlook emails and threads across all connected mailboxes. '
                'Read inbox/folders, navigate full conversation threads, search, download/expand attachments, '
                'and send/reply/forward with rich HTML formatting, selective recipients, multiple CC/BCC, and file attachments.'
            )
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
            res = tool['fn'](p.get('arguments') or {})
            content = res if isinstance(res, list) else [{'type': 'text', 'text': res}]
            reply(msg_id, {'content': content})
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
            try:
                import ast
                args = ast.literal_eval(raw_arg)
            except Exception:
                args = {}
        res = tool['fn'](args)
        if isinstance(res, list):
            for b in res:
                print(b['text'] if b.get('type') == 'text'
                      else '[%s %s, %d B base64]' % (b.get('type'), b.get('mimeType'), len(b.get('data') or '')))
        else:
            print(res)
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
