"""Outlook API client for New Outlook / OE Copilot / Exchange / Microsoft Graph.

Provides clean Python functions over Outlook REST API v2.0 and Microsoft Graph v1.0.
"""
import base64
import datetime as dt
import html
import json
import mimetypes
import os
import re
import urllib.error
import urllib.parse
import urllib.request

import outlook_tokens as tokens


# ---------------------------------------------------------------- HTML <-> Markdown / Rich Body

def _unwrap_safelink(url: str) -> str:
    """Unwrap Microsoft Defender SafeLinks to their original target URL."""
    if not url:
        return url
    try:
        parsed = urllib.parse.urlparse(url)
        if 'safelinks.protection.outlook.com' in (parsed.netloc or '').lower():
            qs = urllib.parse.parse_qs(parsed.query)
            if 'url' in qs and qs['url']:
                return qs['url'][0]
    except Exception:
        pass
    return url


def html_to_markdown(raw_html: str) -> str:
    """Convert HTML email body into clean, human-readable Markdown text."""
    if not raw_html:
        return ""
    s = raw_html
    # Remove head, style, script, xml, meta, conditional comments
    s = re.sub(r'<(style|script|head|title|xml)[^>]*>.*?</\1>', '', s, flags=re.S | re.I)
    s = re.sub(r'<!\[if[^\]]*\]>.*?<!\[endif\]>', '', s, flags=re.S | re.I)
    s = re.sub(r'<!--.*?-->', '', s, flags=re.S)

    # Headings
    for level in range(1, 7):
        hashes = '#' * level
        s = re.sub(
            rf'<h{level}[^>]*>(.*?)</h{level}>',
            lambda m, h=hashes: f"\n\n{h} {re.sub(r'<[^>]+>', '', m.group(1)).strip()}\n\n",
            s,
            flags=re.S | re.I
        )

    # Pre / code
    s = re.sub(
        r'<pre[^>]*>(.*?)</pre>',
        lambda m: f"\n```\n{html.unescape(re.sub(r'<[^>]+>', '', m.group(1)).strip())}\n```\n",
        s,
        flags=re.S | re.I
    )
    s = re.sub(
        r'<code[^>]*>(.*?)</code>',
        lambda m: f"`{html.unescape(re.sub(r'<[^>]+>', '', m.group(1)).strip())}`",
        s,
        flags=re.S | re.I
    )

    # Bold / Italic
    s = re.sub(
        r'<(b|strong)(?:\s[^>]*)?>(.*?)</\1>',
        lambda m: f"**{m.group(2).strip()}**" if m.group(2).strip() else '',
        s,
        flags=re.S | re.I
    )
    s = re.sub(
        r'<(i|em)(?:\s[^>]*)?>(.*?)</\1>',
        lambda m: f"*{m.group(2).strip()}*" if m.group(2).strip() else '',
        s,
        flags=re.S | re.I
    )

    # Links (with SafeLinks unwrapping)
    def _link(m):
        url = _unwrap_safelink(html.unescape(m.group(1).strip()))
        txt = re.sub(r'<[^>]+>', '', m.group(2)).strip()
        txt_clean = _unwrap_safelink(html.unescape(txt))
        if not txt_clean or txt_clean == url:
            return url
        return f"[{txt_clean}]({url})"
    s = re.sub(r'<a\s+[^>]*href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', _link, s, flags=re.S | re.I)

    # Tables: convert rows and cells into pipe-delimited Markdown rows
    def _table(m):
        rows = re.findall(r'<tr[^>]*>(.*?)</tr>', m.group(1), flags=re.S | re.I)
        if not rows:
            return m.group(0)
        md_rows = []
        for idx, r in enumerate(rows):
            cells = re.findall(r'<(?:td|th)[^>]*>(.*?)</(?:td|th)>', r, flags=re.S | re.I)
            if not cells:
                continue
            cleaned = []
            for c in cells:
                c_txt = re.sub(r'<br\s*/?>', ' ', c, flags=re.I)
                c_txt = re.sub(r'<[^>]+>', ' ', c_txt)
                c_txt = html.unescape(c_txt).replace('\xa0', ' ')
                c_txt = re.sub(r'\s+', ' ', c_txt).strip().replace('|', '\\|')
                cleaned.append(c_txt)
            if any(cleaned):
                md_rows.append('| ' + ' | '.join(cleaned) + ' |')
                if len(md_rows) == 1:
                    md_rows.append('| ' + ' | '.join('---' for _ in cleaned) + ' |')
        return '\n\n' + '\n'.join(md_rows) + '\n\n' if md_rows else ''
    s = re.sub(r'<table[^>]*>(.*?)</table>', _table, s, flags=re.S | re.I)

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
    s = re.sub(r'</(p|div|tr|ul|ol|table)>', '\n\n', s, flags=re.I)
    s = re.sub(r'<hr\s*[^>]*/?>', '\n---\n', s, flags=re.I)

    # Strip remaining tags
    s = re.sub(r'<[^>]+>', '', s)
    s = html.unescape(s)

    # Normalize whitespaces & zero-width characters
    s = s.replace('\xa0', ' ').replace('\u200b', '').replace('\r\n', '\n')
    s = re.sub(r'[ \t]+', ' ', s)
    s = re.sub(r' *\n *', '\n', s)
    s = re.sub(r'\n{3,}', '\n\n', s)
    return s.strip()


_HTML_TAG_RE = re.compile(
    r'<\s*(?:p|div|span|b|strong|i|em|u|mark|br|hr|ul|ol|li|table|thead|tbody|tr|th|td|h[1-6]|a|blockquote|pre|code|font|img)\b',
    re.I
)


def looks_like_html(text: str) -> bool:
    """Return True if text already contains HTML formatting tags."""
    return bool(text and _HTML_TAG_RE.search(text))


def format_body_for_html(body: str, is_html: bool = False) -> str:
    """Prepare body content for an HTML email / reply / forward comment.

    If `is_html` is True or `body` contains HTML tags, preserves the rich HTML directly.
    Otherwise, escapes plain text and converts newlines to `<br>` so Exchange's HTML
    composer never collapses line breaks.
    """
    if not body:
        return ""
    if is_html or looks_like_html(body):
        return body
    escaped = html.escape(body).replace('\r\n', '\n').replace('\n', '<br>\n')
    return f'<div style="font-family:Aptos,Calibri,sans-serif;font-size:11pt;">{escaped}</div>'


def parse_recipients(val) -> list:
    """Normalize recipients to list of {"EmailAddress": {"Address": addr, ...}}."""
    if not val:
        return []
    if isinstance(val, str):
        parts = [p.strip() for p in re.split(r'[,;]', val) if p.strip()]
    elif isinstance(val, (list, tuple)):
        parts = []
        for item in val:
            if isinstance(item, str) and (',' in item or ';' in item) and '<' not in item:
                parts.extend(p.strip() for p in re.split(r'[,;]', item) if p.strip())
            else:
                parts.append(item)
    else:
        parts = [str(val)]

    res = []
    seen = set()
    for p in parts:
        entry = None
        if isinstance(p, dict):
            if 'EmailAddress' in p:
                entry = p
            elif 'Address' in p:
                entry = {'EmailAddress': {'Address': p['Address'], 'Name': p.get('Name', '')}}
            elif 'address' in p:
                entry = {'EmailAddress': {'Address': p['address'], 'Name': p.get('name', '')}}
        elif isinstance(p, str):
            m = re.search(r'<([^>]+)>', p)
            if m:
                addr = m.group(1).strip()
                name = p[:m.start()].strip(' "\'')
                entry = {'EmailAddress': {'Address': addr, 'Name': name}} if name else {'EmailAddress': {'Address': addr}}
            elif p.strip():
                entry = {'EmailAddress': {'Address': p.strip()}}
        if entry:
            addr_key = (entry.get('EmailAddress', {}).get('Address') or '').strip().lower()
            if addr_key and addr_key not in seen:
                seen.add(addr_key)
                res.append(entry)
    return res


def parse_address_set(val) -> set:
    """Return a lowercase set of email addresses from a string or list."""
    return {
        r['EmailAddress']['Address'].strip().lower()
        for r in parse_recipients(val)
        if r.get('EmailAddress', {}).get('Address')
    }


# ---------------------------------------------------------------- HTTP Client

def _request(account: str = None, path: str = '', method: str = 'GET', data: dict = None, retry: bool = True):
    info = tokens.get_token_for_account(account)
    base_url = info['base_url'].rstrip('/')
    url = f"{base_url}/{path.lstrip('/')}"

    headers = {
        'Authorization': f"Bearer {info['token']}",
        'Accept': 'application/json',
        'User-Agent': 'OutlookMCP/2.0.0'
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
    """Return all detected Outlook accounts (local + OE Copilot phone)."""
    accts = tokens.get_accounts()
    now = dt.datetime.now().timestamp()
    out = []
    for upn, info in accts.items():
        exp = info.get('exp', 0)
        out.append({
            'account': upn,
            'source': info.get('source', 'local'),
            'endpoint_type': info.get('endpoint_type'),
            'expires_in_minutes': max(0, int((exp - now) / 60)),
            'phone_status': info.get('phone_status'),
            'phone_msg': info.get('phone_msg')
        })
    return out


def list_folders(account: str = None, include_child_folders: bool = True) -> list:
    """List mail folders (including child folders) for the account."""
    res = _request(account, 'me/MailFolders?$top=100')
    folders = []
    for f in res.get('value', []):
        fid = f.get('Id')
        fname = f.get('DisplayName')
        child_count = int(f.get('ChildFolderCount') or 0)
        folders.append({
            'id': fid,
            'name': fname,
            'total_count': f.get('TotalItemCount', 0),
            'unread_count': f.get('UnreadItemCount', 0),
            'child_count': child_count
        })
        if include_child_folders and child_count > 0 and fid:
            try:
                cres = _request(account, f"me/MailFolders/{urllib.parse.quote(fid)}/childfolders?$top=100")
                for cf in cres.get('value', []):
                    folders.append({
                        'id': cf.get('Id'),
                        'name': f"{fname}/{cf.get('DisplayName')}",
                        'display_name': cf.get('DisplayName'),
                        'total_count': cf.get('TotalItemCount', 0),
                        'unread_count': cf.get('UnreadItemCount', 0),
                        'child_count': int(cf.get('ChildFolderCount') or 0)
                    })
            except Exception:
                pass
    return folders


WELL_KNOWN_FOLDERS = {
    'inbox': 'Inbox',
    'posta in arrivo': 'Inbox',
    'sent': 'SentItems',
    'sentitems': 'SentItems',
    'sent items': 'SentItems',
    'posta inviata': 'SentItems',
    'drafts': 'Drafts',
    'bozze': 'Drafts',
    'archive': 'Archive',
    'archivio': 'Archive',
    'deleted': 'DeletedItems',
    'deleteditems': 'DeletedItems',
    'deleted items': 'DeletedItems',
    'trash': 'DeletedItems',
    'posta eliminata': 'DeletedItems',
    'junk': 'JunkEmail',
    'junkemail': 'JunkEmail',
    'spam': 'JunkEmail',
    'posta indesiderata': 'JunkEmail',
    'outbox': 'Outbox',
    'posta in uscita': 'Outbox',
}


def resolve_folder_id(account: str, folder: str) -> str:
    """Resolve a folder name or well-known alias to a MailFolders path segment or ID."""
    if not folder:
        return 'Inbox'
    clean = folder.strip()
    if clean.lower() in WELL_KNOWN_FOLDERS:
        return WELL_KNOWN_FOLDERS[clean.lower()]
    if clean.startswith('AAMk') or clean.startswith('AQMk'):
        return clean
    folders = list_folders(account, include_child_folders=True)
    # Exact match first
    for f in folders:
        if clean.casefold() in (f['name'].casefold(), (f.get('display_name') or '').casefold()):
            return f['id']
    # Substring match second
    for f in folders:
        if clean.casefold() in f['name'].casefold():
            return f['id']
    return None


def _normalize_iso_date(d_str: str, end_of_day: bool = False) -> str:
    if not d_str:
        return None
    s = d_str.strip()
    if len(s) == 10 and re.match(r'^\d{4}-\d{2}-\d{2}$', s):
        return f"{s}T23:59:59Z" if end_of_day else f"{s}T00:00:00Z"
    if not s.endswith('Z') and '+' not in s[10:] and '-' not in s[10:]:
        return s + 'Z'
    return s


def _matches_local_filters(msg: dict, unread_only: bool = False, from_address: str = None,
                           has_attachments: bool = None, start_date: str = None,
                           end_date: str = None, flagged_only: bool = False) -> bool:
    if unread_only and msg.get('IsRead') is not False:
        return False
    if has_attachments is True and not msg.get('HasAttachments'):
        return False
    if has_attachments is False and msg.get('HasAttachments'):
        return False
    if flagged_only:
        flag_status = ((msg.get('Flag') or {}).get('FlagStatus') or '').lower()
        if flag_status != 'flagged':
            return False
    if from_address:
        needle = from_address.casefold()
        sender_obj = (msg.get('From') or msg.get('Sender') or {}).get('EmailAddress') or {}
        haystack = f"{sender_obj.get('Name', '')} {sender_obj.get('Address', '')}".casefold()
        if needle not in haystack:
            return False
    msg_dt = msg.get('ReceivedDateTime') or msg.get('CreatedDateTime') or ''
    if start_date and msg_dt and msg_dt < _normalize_iso_date(start_date, False):
        return False
    if end_date and msg_dt and msg_dt > _normalize_iso_date(end_date, True):
        return False
    return True


def list_emails(account: str = None, folder: str = 'Inbox', query: str = None,
                limit: int = 20, unread_only: bool = False, from_address: str = None,
                has_attachments: bool = None, start_date: str = None,
                end_date: str = None, flagged_only: bool = False) -> list:
    """List emails with folder, search, sender, attachment, flag, and date filters."""
    if folder and folder.strip().lower() in ('all', '*', 'any'):
        path = 'me/messages'
    else:
        fid = resolve_folder_id(account, folder or 'Inbox')
        path = f"me/MailFolders/{urllib.parse.quote(fid)}/messages" if fid else 'me/messages'

    need_post_filter = bool(from_address or flagged_only or (
        query and (unread_only or has_attachments is not None or start_date or end_date)
    ))
    fetch_top = min(max(limit * 4, 50), 100) if need_post_filter else max(1, min(limit, 100))

    select_fields = (
        'Id,Subject,From,Sender,ToRecipients,CcRecipients,ReceivedDateTime,CreatedDateTime,'
        'BodyPreview,HasAttachments,IsRead,IsDraft,Importance,Flag,ConversationId,ParentFolderId'
    )
    params = [
        f'$top={fetch_top}',
        f'$select={select_fields}'
    ]

    if query:
        params.append('$search=' + urllib.parse.quote(f'"{query}"'))
    else:
        filters = []
        if unread_only:
            filters.append('IsRead eq false')
        if has_attachments is True:
            filters.append('HasAttachments eq true')
        elif has_attachments is False:
            filters.append('HasAttachments eq false')
        if start_date:
            filters.append(f"ReceivedDateTime ge {_normalize_iso_date(start_date, False)}")
        if end_date:
            filters.append(f"ReceivedDateTime le {_normalize_iso_date(end_date, True)}")
        if filters:
            params.append('$filter=' + urllib.parse.quote(' and '.join(filters)))
        params.append('$orderby=' + urllib.parse.quote('ReceivedDateTime desc'))

    full_path = f"{path}?{'&'.join(params)}"
    res = _request(account, full_path)
    items = res.get('value', [])

    if need_post_filter:
        items = [
            m for m in items
            if _matches_local_filters(
                m, unread_only=unread_only, from_address=from_address,
                has_attachments=has_attachments, start_date=start_date,
                end_date=end_date, flagged_only=flagged_only
            )
        ]

    return items[:limit]


def get_message(account: str = None, message_id: str = None) -> dict:
    """Retrieve full details of one message, including body and attachments."""
    if not message_id:
        raise ValueError('message_id is required')
    path = f"me/messages/{urllib.parse.quote(message_id)}"
    msg = _request(account, path)

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
    """Retrieve all messages in a thread (ConversationId or message_id) sorted oldest to newest."""
    if not thread_id:
        raise ValueError('thread_id or message_id is required')

    cid = thread_id
    if not thread_id.startswith('AAQk'):
        try:
            m = _request(account, f"me/messages/{urllib.parse.quote(thread_id)}?$select=ConversationId")
            if m.get('ConversationId'):
                cid = m['ConversationId']
        except Exception:
            pass

    filter_expr = urllib.parse.quote(f"ConversationId eq '{cid}'")
    select_fields = (
        'Id,Subject,From,Sender,ToRecipients,CcRecipients,BccRecipients,ReplyTo,'
        'ReceivedDateTime,SentDateTime,CreatedDateTime,Body,BodyPreview,'
        'HasAttachments,IsRead,IsDraft,Importance,Flag,ConversationId'
    )
    params = f"$filter={filter_expr}&$top={max(1, min(limit, 100))}&$select={select_fields}"
    res = _request(account, f"me/messages?{params}")
    messages = res.get('value', [])

    def _msg_time(m):
        return m.get('ReceivedDateTime') or m.get('SentDateTime') or m.get('CreatedDateTime') or ''
    messages.sort(key=_msg_time)

    for m in messages:
        if m.get('HasAttachments'):
            try:
                mid = m['Id']
                att_res = _request(
                    account,
                    f"me/messages/{urllib.parse.quote(mid)}/attachments?$select=Id,Name,ContentType,Size,IsInline"
                )
                m['AttachmentsList'] = att_res.get('value', [])
            except Exception:
                m['AttachmentsList'] = []
        else:
            m['AttachmentsList'] = []

    return messages


def search_emails(account: str = None, query: str = None, folder: str = None,
                  from_address: str = None, unread_only: bool = False,
                  has_attachments: bool = None, start_date: str = None,
                  end_date: str = None, limit: int = 20) -> list:
    """Full-text search across all folders (or a specific folder) with optional filters."""
    if not query and not from_address:
        raise ValueError('query or from_address is required')
    return list_emails(
        account=account,
        folder=folder or 'all',
        query=query,
        limit=limit,
        unread_only=unread_only,
        from_address=from_address,
        has_attachments=has_attachments,
        start_date=start_date,
        end_date=end_date
    )


def list_attachments(account: str = None, message_id: str = None) -> list:
    """List attachment metadata (no content bytes) for one message."""
    if not message_id:
        raise ValueError('message_id is required')
    path = f"me/messages/{urllib.parse.quote(message_id)}/attachments?$select=Id,Name,ContentType,Size,IsInline"
    res = _request(account, path)
    return res.get('value', [])


def get_attachment(account: str = None, message_id: str = None, attachment_id: str = None,
                   expand_item: bool = False) -> dict:
    """Fetch one attachment in full, including ContentBytes or expanded ItemAttachment."""
    if not message_id:
        raise ValueError('message_id is required')
    if not attachment_id:
        raise ValueError('attachment_id is required')
    path = f"me/messages/{urllib.parse.quote(message_id)}/attachments/{urllib.parse.quote(attachment_id)}"
    if expand_item:
        info = tokens.get_token_for_account(account)
        if info.get('endpoint_type') == 'graph':
            path += '?$expand=microsoft.graph.itemAttachment/item'
        else:
            path += '?$expand=Microsoft.OutlookServices.ItemAttachment/Item'
    return _request(account, path)


def _attachment_objects(attachments: list) -> list:
    """Build FileAttachment objects from local file paths."""
    att_objs = []
    for att in attachments or []:
        if isinstance(att, str):
            p = os.path.expanduser(att.strip())
            if not os.path.isfile(p):
                raise FileNotFoundError(f"Attachment file not found: {p}")
            with open(p, 'rb') as fp:
                content_b64 = base64.b64encode(fp.read()).decode('ascii')
            ctype, _ = mimetypes.guess_type(p)
            obj = {
                '@odata.type': '#Microsoft.OutlookServices.FileAttachment',
                'Name': os.path.basename(p),
                'ContentBytes': content_b64
            }
            if ctype:
                obj['ContentType'] = ctype
            att_objs.append(obj)
    return att_objs


def _merge_recipients(existing: list, extra: list) -> list:
    """Append extra recipients to existing ones, skipping duplicate addresses."""
    merged = list(existing or [])
    seen = {
        (r.get('EmailAddress', {}).get('Address') or '').strip().lower()
        for r in merged
        if r.get('EmailAddress', {}).get('Address')
    }
    for r in extra or []:
        addr = (r.get('EmailAddress', {}).get('Address') or '').strip().lower()
        if addr and addr not in seen:
            seen.add(addr)
            merged.append(r)
    return merged


def _filter_excluded(recs: list, excluded: set) -> list:
    if not excluded:
        return list(recs or [])
    return [
        r for r in (recs or [])
        if (r.get('EmailAddress', {}).get('Address') or '').strip().lower() not in excluded
    ]


def _addr_list(recs: list) -> list:
    return [
        r['EmailAddress']['Address']
        for r in (recs or [])
        if r.get('EmailAddress', {}).get('Address')
    ]


def _reply_or_forward_via_draft(
    account: str,
    target_id: str,
    action: str,
    body: str,
    is_html: bool = False,
    to=None,
    cc=None,
    bcc=None,
    replace_cc: bool = False,
    exclude_recipients=None,
    subject: str = None,
    importance: str = None,
    attachments: list = None,
    save_as_draft: bool = False
) -> dict:
    """Create a reply, reply-all, or forward draft, customize recipients/attachments/body, and send or keep as draft."""
    html_comment = format_body_for_html(body, is_html=is_html)
    path = f"me/messages/{urllib.parse.quote(target_id)}/{action}"
    draft = _request(account, path, method='POST', data={'Comment': html_comment})
    draft_id = draft.get('Id')
    if not draft_id:
        raise RuntimeError(f"Failed to create draft via {action} on message {target_id}")
    draft_path = f"me/messages/{urllib.parse.quote(draft_id)}"

    excluded = parse_address_set(exclude_recipients)
    parsed_to = parse_recipients(to)
    parsed_cc = parse_recipients(cc)
    parsed_bcc = parse_recipients(bcc)

    if parsed_to:
        to_recipients = parsed_to
    else:
        to_recipients = draft.get('ToRecipients') or []

    if replace_cc or action != 'createreplyall':
        cc_recipients = parsed_cc if parsed_cc else (
            [] if replace_cc else (draft.get('CcRecipients') or [])
        )
    else:
        cc_recipients = _merge_recipients(draft.get('CcRecipients'), parsed_cc)

    bcc_recipients = _merge_recipients(draft.get('BccRecipients'), parsed_bcc)

    if excluded:
        to_recipients = _filter_excluded(to_recipients, excluded)
        cc_recipients = _filter_excluded(cc_recipients, excluded)
        bcc_recipients = _filter_excluded(bcc_recipients, excluded)

    if not to_recipients and not save_as_draft:
        _request(account, draft_path, method='DELETE')
        raise ValueError(
            "No 'To' recipients left after applying recipient filters. Specify 'to' explicitly."
        )

    patch_data = {
        'ToRecipients': to_recipients,
        'CcRecipients': cc_recipients,
        'BccRecipients': bcc_recipients
    }
    if subject and subject.strip():
        patch_data['Subject'] = subject.strip()
    if importance and importance.capitalize() in ('Low', 'Normal', 'High'):
        patch_data['Importance'] = importance.capitalize()

    try:
        _request(account, draft_path, method='PATCH', data=patch_data)
        for att in _attachment_objects(attachments):
            _request(account, f"{draft_path}/attachments", method='POST', data=att)
        if not save_as_draft:
            _request(account, f"{draft_path}/send", method='POST')
    except Exception:
        if not save_as_draft:
            try:
                _request(account, draft_path, method='DELETE')
            except Exception:
                pass
        raise

    return {
        'status': 'draft' if save_as_draft else 'sent',
        'type': action,
        'target_message_id': target_id,
        'draft_id': draft_id if save_as_draft else None,
        'subject': (subject or draft.get('Subject') or '').strip(),
        'to': _addr_list(to_recipients),
        'cc': _addr_list(cc_recipients),
        'bcc': _addr_list(bcc_recipients)
    }


def send_email(account: str = None, to=None, subject: str = '', body: str = '',
               cc=None, bcc=None, is_html: bool = False,
               reply_to_id: str = None, reply_all: bool = False,
               forward_to_id: str = None, replace_cc: bool = False,
               exclude_recipients=None, importance: str = None,
               attachments: list = None, save_as_draft: bool = False) -> dict:
    """Send a new email, reply/reply-all to a thread, forward an email, or save as draft."""
    if reply_to_id:
        action = 'createreplyall' if reply_all else 'createreply'
        return _reply_or_forward_via_draft(
            account=account,
            target_id=reply_to_id,
            action=action,
            body=body,
            is_html=is_html,
            to=to,
            cc=cc,
            bcc=bcc,
            replace_cc=replace_cc,
            exclude_recipients=exclude_recipients,
            subject=subject,
            importance=importance,
            attachments=attachments,
            save_as_draft=save_as_draft
        )

    if forward_to_id:
        if not to and not save_as_draft:
            raise ValueError("Recipient ('to') is required when forwarding an email")
        return _reply_or_forward_via_draft(
            account=account,
            target_id=forward_to_id,
            action='createforward',
            body=body,
            is_html=is_html,
            to=to,
            cc=cc,
            bcc=bcc,
            replace_cc=True,
            exclude_recipients=exclude_recipients,
            subject=subject,
            importance=importance,
            attachments=attachments,
            save_as_draft=save_as_draft
        )

    if not to and not save_as_draft:
        raise ValueError('Recipient (to) is required for new email')

    to_recipients = parse_recipients(to)
    cc_recipients = parse_recipients(cc) if cc else []
    bcc_recipients = parse_recipients(bcc) if bcc else []

    use_html = bool(is_html or looks_like_html(body))
    message = {
        'Subject': subject,
        'Body': {
            'ContentType': 'HTML' if use_html else 'Text',
            'Content': body
        },
        'ToRecipients': to_recipients
    }
    if cc_recipients:
        message['CcRecipients'] = cc_recipients
    if bcc_recipients:
        message['BccRecipients'] = bcc_recipients
    if importance and importance.capitalize() in ('Low', 'Normal', 'High'):
        message['Importance'] = importance.capitalize()

    att_objs = _attachment_objects(attachments)

    if save_as_draft:
        draft = _request(account, 'me/messages', method='POST', data=message)
        draft_id = draft.get('Id')
        for att in att_objs:
            _request(account, f"me/messages/{urllib.parse.quote(draft_id)}/attachments", method='POST', data=att)
        return {
            'status': 'draft',
            'draft_id': draft_id,
            'subject': subject,
            'to': _addr_list(to_recipients),
            'cc': _addr_list(cc_recipients),
            'bcc': _addr_list(bcc_recipients)
        }

    if att_objs:
        message['Attachments'] = att_objs

    payload = {
        'Message': message,
        'SaveToSentItems': 'true'
    }
    _request(account, 'me/sendmail', method='POST', data=payload)
    return {
        'status': 'sent',
        'subject': subject,
        'to': _addr_list(to_recipients),
        'cc': _addr_list(cc_recipients),
        'bcc': _addr_list(bcc_recipients)
    }


def manage_email(account: str = None, message_id=None, action: str = None,
                 destination_folder: str = None) -> list:
    """Perform read/unread, flag/unflag, archive, delete, or move actions on one or more messages."""
    if not message_id:
        raise ValueError('message_id is required')
    ids = [message_id] if isinstance(message_id, str) else list(message_id)
    act = (action or '').strip().lower()
    if act not in ('mark_read', 'mark_unread', 'flag', 'unflag', 'archive', 'delete', 'move'):
        raise ValueError(
            f"Unsupported action '{action}'. Use: mark_read, mark_unread, flag, unflag, archive, delete, move."
        )

    dest_id = None
    if act == 'archive':
        dest_id = 'Archive'
    elif act == 'delete':
        dest_id = 'DeletedItems'
    elif act == 'move':
        if not destination_folder:
            raise ValueError('destination_folder is required when action is move')
        dest_id = resolve_folder_id(account, destination_folder)
        if not dest_id:
            raise ValueError(f"Destination folder '{destination_folder}' not found")

    results = []
    for mid in ids:
        mpath = f"me/messages/{urllib.parse.quote(mid)}"
        if act == 'mark_read':
            _request(account, mpath, method='PATCH', data={'IsRead': True})
            results.append({'id': mid, 'status': 'marked_read'})
        elif act == 'mark_unread':
            _request(account, mpath, method='PATCH', data={'IsRead': False})
            results.append({'id': mid, 'status': 'marked_unread'})
        elif act == 'flag':
            _request(account, mpath, method='PATCH', data={'Flag': {'FlagStatus': 'Flagged'}})
            results.append({'id': mid, 'status': 'flagged'})
        elif act == 'unflag':
            _request(account, mpath, method='PATCH', data={'Flag': {'FlagStatus': 'NotFlagged'}})
            results.append({'id': mid, 'status': 'unflagged'})
        elif act in ('archive', 'delete', 'move'):
            moved = _request(account, f"{mpath}/move", method='POST', data={'DestinationId': dest_id})
            results.append({'id': mid, 'new_id': moved.get('Id', mid), 'status': f"moved_to_{dest_id}"})
    return results
