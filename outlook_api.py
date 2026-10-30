"""Outlook API client for New Outlook / Exchange / Microsoft Graph.

Provides clean Python functions over Outlook REST API v2.0 and Microsoft Graph v1.0.
"""
import base64
import html
import json
import os
import re
import urllib.error
import urllib.parse
import urllib.request
import datetime as dt

import outlook_tokens as tokens


# ---------------------------------------------------------------- HTML -> Markdown

def html_to_markdown(raw_html: str) -> str:
    """Convert HTML email body into clean, human-readable markdown text."""
    if not raw_html:
        return ""
    s = raw_html
    # Remove head, style, script, xml, meta
    s = re.sub(r'<(style|script|head|title)[^>]*>.*?</\1>', '', s, flags=re.S | re.I)
    s = re.sub(r'<!\[if[^\]]*\]>.*?<!\[endif\]>', '', s, flags=re.S | re.I)
    s = re.sub(r'<!--.*?-->', '', s, flags=re.S)
    
    # Links
    def _link(m):
        url = m.group(1).strip()
        txt = re.sub(r'<[^>]+>', '', m.group(2)).strip()
        if not txt or txt == url:
            return url
        return f"[{txt}]({url})"
    s = re.sub(r'<a\s+[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', _link, s, flags=re.S | re.I)
    
    # Blockquotes / replies
    def _quote(m):
        inner = html_to_markdown(m.group(1))
        return '\n' + '\n'.join('> ' + l for l in inner.splitlines() if l.strip()) + '\n'
    s = re.sub(r'<blockquote[^>]*>(.*?)</blockquote>', _quote, s, flags=re.S | re.I)
    
    # Lists
    s = re.sub(r'<li[^>]*>', '\n* ', s, flags=re.I)
    s = re.sub(r'</li>', '', s, flags=re.I)
    
    # Line breaks and block elements
    s = re.sub(r'<br\s*/?>', '\n', s, flags=re.I)
    s = re.sub(r'</(p|div|tr|h\d)>', '\n\n', s, flags=re.I)
    s = re.sub(r'<hr\s*/?>', '\n---\n', s, flags=re.I)
    
    # Strip remaining tags
    s = re.sub(r'<[^>]+>', '', s)
    s = html.unescape(s)
    
    # Normalize whitespaces & zero-width characters
    s = s.replace('\xa0', ' ').replace('\u200b', '')
    s = re.sub(r'[ \t]+', ' ', s)
    s = re.sub(r' *\n *', '\n', s)
    s = re.sub(r'\n{3,}', '\n\n', s)
    return s.strip()


def parse_recipients(val) -> list:
    """Normalize recipients to list of {"EmailAddress": {"Address": addr}}."""
    if not val:
        return []
    if isinstance(val, str):
        parts = [p.strip() for p in re.split(r'[,;]', val) if p.strip()]
    elif isinstance(val, (list, tuple)):
        parts = val
    else:
        parts = [str(val)]
    
    res = []
    for p in parts:
        if isinstance(p, dict):
            if 'EmailAddress' in p:
                res.append(p)
            elif 'Address' in p:
                res.append({'EmailAddress': {'Address': p['Address'], 'Name': p.get('Name', '')}})
            elif 'address' in p:
                res.append({'EmailAddress': {'Address': p['address'], 'Name': p.get('name', '')}})
        else:
            # string e.g. "Name <addr@domain>" or "addr@domain"
            m = re.search(r'<([^>]+)>', p)
            if m:
                addr = m.group(1).strip()
                name = p[:m.start()].strip(' "\'')
                res.append({'EmailAddress': {'Address': addr, 'Name': name}})
            else:
                res.append({'EmailAddress': {'Address': p.strip()}})
    return res


# ---------------------------------------------------------------- HTTP Client

def _request(account: str = None, path: str = '', method: str = 'GET', data: dict = None, retry: bool = True):
    info = tokens.get_token_for_account(account)
    base_url = info['base_url'].rstrip('/')
    url = f"{base_url}/{path.lstrip('/')}"
    
    headers = {
        'Authorization': f"Bearer {info['token']}",
        'Accept': 'application/json',
        'User-Agent': 'OutlookMCP/1.0.0 (Windows)'
    }
    
    body_bytes = None
    if data is not None:
        headers['Content-Type'] = 'application/json; charset=utf-8'
        body_bytes = json.dumps(data).encode('utf-8')
    
    req = urllib.request.Request(url, data=body_bytes, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            status = resp.status
            if status == 204 or resp.length == 0:
                return {}
            raw = resp.read()
            if not raw:
                return {}
            return json.loads(raw.decode('utf-8', errors='replace'))
    except urllib.error.HTTPError as e:
        if e.code == 401 and retry:
            # Refresh token and retry once
            tokens.get_accounts(refresh=True)
            return _request(account, path, method, data, retry=False)
        err_body = ''
        try:
            err_body = e.read().decode('utf-8', errors='replace')
        except Exception:
            pass
        raise RuntimeError(f"Outlook API error {e.code} on {method} {url}: {err_body or e.reason}")


# ---------------------------------------------------------------- Core Operations

def list_accounts() -> list:
    """Return all detected Outlook accounts."""
    accts = tokens.get_accounts()
    now = dt.datetime.now().timestamp()
    out = []
    for upn, info in accts.items():
        exp = info.get('exp', 0)
        out.append({
            'account': upn,
            'endpoint_type': info.get('endpoint_type'),
            'expires_in_minutes': max(0, int((exp - now) / 60))
        })
    return out


def list_folders(account: str = None) -> list:
    """List mail folders for the account."""
    res = _request(account, 'me/MailFolders?$top=50')
    folders = []
    for f in res.get('value', []):
        folders.append({
            'id': f.get('Id'),
            'name': f.get('DisplayName'),
            'total_count': f.get('TotalItemCount', 0),
            'unread_count': f.get('UnreadItemCount', 0)
        })
    return folders


def list_emails(account: str = None, folder: str = 'Inbox', query: str = None, limit: int = 25, unread_only: bool = False) -> list:
    """List recent emails with search and folder options."""
    # Determine base folder path
    if not folder or folder.lower() in ('inbox', ''):
        path = 'me/MailFolders/Inbox/messages'
    elif folder.lower() in ('sent', 'sentitems'):
        path = 'me/MailFolders/SentItems/messages'
    elif folder.lower() in ('drafts',):
        path = 'me/MailFolders/Drafts/messages'
    elif folder.lower() in ('archive',):
        path = 'me/MailFolders/Archive/messages'
    elif folder.lower() in ('deleted', 'deleteditems', 'trash'):
        path = 'me/MailFolders/DeletedItems/messages'
    elif folder.startswith('AAMk'):
        # Folder ID given directly
        path = f'me/MailFolders/{urllib.parse.quote(folder)}/messages'
    else:
        # Search folder by name
        folders = list_folders(account)
        matched = [f for f in folders if folder.casefold() in f['name'].casefold()]
        if matched:
            path = f"me/MailFolders/{urllib.parse.quote(matched[0]['id'])}/messages"
        else:
            path = 'me/messages'

    params = [
        f'$top={max(1, min(limit, 100))}',
        '$select=Id,Subject,From,Sender,ToRecipients,ReceivedDateTime,CreatedDateTime,BodyPreview,HasAttachments,IsRead,ConversationId,ParentFolderId'
    ]

    filters = []
    if unread_only:
        filters.append('IsRead eq false')
    if filters:
        params.append('$filter=' + urllib.parse.quote(' and '.join(filters)))

    if query:
        params.append('$search=' + urllib.parse.quote(f'"{query}"'))

    full_path = f"{path}?{'&'.join(params)}"
    res = _request(account, full_path)
    return res.get('value', [])


def get_message(account: str = None, message_id: str = None) -> dict:
    """Retrieve full details of one message, including body and attachments."""
    if not message_id:
        raise ValueError('message_id is required')
    path = f"me/messages/{urllib.parse.quote(message_id)}"
    msg = _request(account, path)
    
    # If it has attachments, list them
    attachments = []
    if msg.get('HasAttachments'):
        try:
            att_res = _request(account, f"{path}/attachments?$select=Id,Name,ContentType,Size,IsInline")
            attachments = att_res.get('value', [])
        except Exception:
            pass
    msg['AttachmentsList'] = attachments
    return msg


def read_thread(account: str = None, thread_id: str = None, limit: int = 50) -> list:
    """Retrieve all messages in a thread (ConversationId) sorted chronologically."""
    if not thread_id:
        raise ValueError('thread_id or message_id is required')
    
    cid = thread_id
    # If user passed a message ID instead of ConversationId
    if not thread_id.startswith('AAQk'):
        try:
            m = _request(account, f"me/messages/{urllib.parse.quote(thread_id)}?$select=ConversationId")
            if m.get('ConversationId'):
                cid = m['ConversationId']
        except Exception:
            pass

    filter_expr = urllib.parse.quote(f"ConversationId eq '{cid}'")
    params = f"$filter={filter_expr}&$top={max(1, min(limit, 100))}&$select=Id,Subject,From,Sender,ToRecipients,CcRecipients,ReceivedDateTime,CreatedDateTime,Body,BodyPreview,HasAttachments,IsRead,ConversationId"
    res = _request(account, f"me/messages?{params}")
    messages = res.get('value', [])
    
    # Sort ascending by arrival time
    def _msg_time(m):
        return m.get('ReceivedDateTime') or m.get('CreatedDateTime') or ''
    messages.sort(key=_msg_time)
    
    # Fetch attachments list for messages with attachments
    for m in messages:
        if m.get('HasAttachments'):
            try:
                mid = m['Id']
                att_res = _request(account, f"me/messages/{urllib.parse.quote(mid)}/attachments?$select=Id,Name,ContentType,Size,IsInline")
                m['AttachmentsList'] = att_res.get('value', [])
            except Exception:
                m['AttachmentsList'] = []
        else:
            m['AttachmentsList'] = []
            
    return messages


def search_emails(account: str = None, query: str = None, limit: int = 25) -> list:
    """Full-text search across all messages."""
    if not query:
        raise ValueError('query is required')
    params = [
        f'$top={max(1, min(limit, 100))}',
        f'$search={urllib.parse.quote(f"\"{query}\"")}',
        '$select=Id,Subject,From,Sender,ToRecipients,ReceivedDateTime,CreatedDateTime,BodyPreview,HasAttachments,IsRead,ConversationId'
    ]
    path = f"me/messages?{'&'.join(params)}"
    res = _request(account, path)
    return res.get('value', [])


def _attachment_objects(attachments: list) -> list:
    """Build FileAttachment objects from local file paths."""
    att_objs = []
    for att in attachments or []:
        if isinstance(att, str) and os.path.isfile(att):
            with open(att, 'rb') as fp:
                content_b64 = base64.b64encode(fp.read()).decode('ascii')
            att_objs.append({
                '@odata.type': '#Microsoft.OutlookServices.FileAttachment',
                'Name': os.path.basename(att),
                'ContentBytes': content_b64
            })
    return att_objs


def _merge_recipients(existing: list, extra: list) -> list:
    """Append extra recipients to existing ones, skipping duplicate addresses."""
    merged = list(existing or [])
    seen = {r['EmailAddress']['Address'].lower() for r in merged}
    for r in extra:
        addr = r['EmailAddress']['Address'].lower()
        if addr not in seen:
            seen.add(addr)
            merged.append(r)
    return merged


def _send_reply_draft(account, reply_to_id, body, reply_all, to, cc, bcc, attachments) -> dict:
    """Reply through a draft: the reply action takes no recipients, a draft does."""
    create = 'createreplyall' if reply_all else 'createreply'
    path = f"me/messages/{urllib.parse.quote(reply_to_id)}/{create}"
    draft = _request(account, path, method='POST', data={'Comment': body})
    draft_path = f"me/messages/{urllib.parse.quote(draft['Id'])}"

    to_recipients = parse_recipients(to) or draft.get('ToRecipients', [])
    cc_recipients = _merge_recipients(draft.get('CcRecipients'), parse_recipients(cc))
    bcc_recipients = _merge_recipients(draft.get('BccRecipients'), parse_recipients(bcc))

    try:
        _request(account, draft_path, method='PATCH', data={
            'ToRecipients': to_recipients,
            'CcRecipients': cc_recipients,
            'BccRecipients': bcc_recipients
        })
        for att in _attachment_objects(attachments):
            _request(account, f"{draft_path}/attachments", method='POST', data=att)
        _request(account, f"{draft_path}/send", method='POST')
    except Exception:
        _request(account, draft_path, method='DELETE')
        raise

    addr = lambda rs: [r['EmailAddress']['Address'] for r in rs]
    return {'status': 'sent', 'type': create, 'target_message_id': reply_to_id,
            'to': addr(to_recipients), 'cc': addr(cc_recipients), 'bcc': addr(bcc_recipients)}


def send_email(account: str = None, to: list = None, subject: str = '', body: str = '',
               cc: list = None, bcc: list = None, is_html: bool = False,
               reply_to_id: str = None, reply_all: bool = False,
               attachments: list = None) -> dict:
    """Send an email or reply to an existing email/thread."""
    if reply_to_id:
        # Reply mode
        method = 'replyall' if reply_all else 'reply'
        if to or cc or bcc or attachments:
            return _send_reply_draft(account, reply_to_id, body, reply_all, to, cc, bcc, attachments)
        payload = {
            'Comment': body
        }
        path = f"me/messages/{urllib.parse.quote(reply_to_id)}/{method}"
        _request(account, path, method='POST', data=payload)
        return {'status': 'sent', 'type': method, 'target_message_id': reply_to_id}

    if not to:
        raise ValueError('Recipient (to) is required for new email')

    to_recipients = parse_recipients(to)
    cc_recipients = parse_recipients(cc) if cc else []
    bcc_recipients = parse_recipients(bcc) if bcc else []

    message = {
        'Subject': subject,
        'Body': {
            'ContentType': 'HTML' if is_html else 'Text',
            'Content': body
        },
        'ToRecipients': to_recipients
    }
    if cc_recipients:
        message['CcRecipients'] = cc_recipients
    if bcc_recipients:
        message['BccRecipients'] = bcc_recipients

    att_objs = _attachment_objects(attachments)
    if att_objs:
        message['Attachments'] = att_objs

    payload = {
        'Message': message,
        'SaveToSentItems': 'true'
    }
    _request(account, 'me/sendmail', method='POST', data=payload)
    return {'status': 'sent', 'subject': subject, 'to': [r['EmailAddress']['Address'] for r in to_recipients]}
