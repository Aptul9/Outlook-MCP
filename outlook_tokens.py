"""Token harvester for Outlook MCP (New Outlook IdentityCache + OE Copilot virtual phone).

Extracts and decrypts OAuth2 bearer tokens from:
1. Local Windows IdentityCache (%LOCALAPPDATA%\\Microsoft\\IdentityCache) via DPAPI.
2. Remote OE Copilot virtual Android phone over SSH (host in OUTLOOK_MCP_SSH_HOST,
   default 'oecopilot-vm'), decrypting Outlook mobile's HxCoreSecureStringAES256Gcm
   tokens from acompliAcct.db and Android Keystore persistent.sqlite.

Standard library only on Windows. Cached tokens are stored DPAPI-encrypted on disk.
"""
import base64
import ctypes
import json
import os
import subprocess
import time
import urllib.error
import urllib.parse
import urllib.request
from ctypes import wintypes

CACHE_DIR = os.path.join(os.environ.get('LOCALAPPDATA', '.'), 'OutlookMCP')
STATE_FILE = os.path.join(CACHE_DIR, 'tokens.bin')

SSH_HOST = os.environ.get('OUTLOOK_MCP_SSH_HOST', 'oecopilot-vm')
REMOTE_DIR = os.environ.get('OUTLOOK_MCP_REMOTE_DIR', '/opt/oe-copilot')

# ---------------------------------------------------------------- DPAPI

class _BLOB(ctypes.Structure):
    _fields_ = [('cbData', wintypes.DWORD), ('pbData', ctypes.POINTER(ctypes.c_char))]

_crypt32 = ctypes.WinDLL('crypt32', use_last_error=True)
_k32 = ctypes.WinDLL('kernel32', use_last_error=True)

_crypt32.CryptProtectData.argtypes = [ctypes.POINTER(_BLOB), ctypes.POINTER(wintypes.LPWSTR),
                                      ctypes.POINTER(_BLOB), ctypes.c_void_p, ctypes.c_void_p,
                                      wintypes.DWORD, ctypes.POINTER(_BLOB)]
_crypt32.CryptProtectData.restype = wintypes.BOOL

_crypt32.CryptUnprotectData.argtypes = [ctypes.POINTER(_BLOB), ctypes.POINTER(wintypes.LPWSTR),
                                        ctypes.POINTER(_BLOB), ctypes.c_void_p, ctypes.c_void_p,
                                        wintypes.DWORD, ctypes.POINTER(_BLOB)]
_crypt32.CryptUnprotectData.restype = wintypes.BOOL
_k32.LocalFree.argtypes = [ctypes.c_void_p]

CRYPTPROTECT_UI_FORBIDDEN = 1


def dpapi_unprotect(data: bytes) -> bytes:
    src = ctypes.create_string_buffer(data, len(data))
    blob_in = _BLOB(len(data), ctypes.cast(src, ctypes.POINTER(ctypes.c_char)))
    blob_out = _BLOB()
    if not _crypt32.CryptUnprotectData(ctypes.byref(blob_in), None, None, None, None,
                                        CRYPTPROTECT_UI_FORBIDDEN, ctypes.byref(blob_out)):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        _k32.LocalFree(blob_out.pbData)


def dpapi_protect(data: bytes) -> bytes:
    src = ctypes.create_string_buffer(data, len(data))
    blob_in = _BLOB(len(data), ctypes.cast(src, ctypes.POINTER(ctypes.c_char)))
    blob_out = _BLOB()
    if not _crypt32.CryptProtectData(ctypes.byref(blob_in), None, None, None, None,
                                      CRYPTPROTECT_UI_FORBIDDEN, ctypes.byref(blob_out)):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        return ctypes.string_at(blob_out.pbData, blob_out.cbData)
    finally:
        _k32.LocalFree(blob_out.pbData)


def jwt_claims(token: str) -> dict:
    try:
        part = token.split('.')[1]
        part += '=' * (-len(part) % 4)
        return json.loads(base64.urlsafe_b64decode(part))
    except Exception:
        return {}


def _classify_token(secret: str, now: float, source: str = 'local', extra: dict = None) -> dict:
    """Validate JWT and classify its endpoint (outlook_rest vs graph)."""
    if not secret or len(secret.split('.')) < 3:
        return None
    claims = jwt_claims(secret)
    upn = claims.get('upn') or claims.get('email') or claims.get('unique_name')
    if not upn:
        return None
    exp = int(claims.get('exp') or 0)
    if exp <= now:
        return None

    aud = str(claims.get('aud', ''))
    scp = str(claims.get('scp', ''))

    endpoint_type = None
    base_url = None

    if 'outlook.office.com' in aud or 'outlook.office365.com' in aud:
        if any(s in scp for s in ('Mail.Read', 'Mail.ReadWrite', 'EAS.AccessAsUser', 'OWA.AccessAsUser')):
            endpoint_type = 'outlook_rest'
            base_url = 'https://outlook.office.com/api/v2.0'
    elif 'graph.microsoft.com' in aud or aud == '00000003-0000-0000-c000-000000000000':
        if 'Mail.Read' in scp or 'Mail.ReadWrite' in scp:
            endpoint_type = 'graph'
            base_url = 'https://graph.microsoft.com/v1.0'

    if not endpoint_type:
        return None

    info = {
        'upn': upn,
        'token': secret,
        'exp': exp,
        'endpoint_type': endpoint_type,
        'base_url': base_url,
        'scopes': scp,
        'source': source
    }
    if extra:
        info.update(extra)
    return info


def _merge_account(accounts: dict, info: dict):
    """Insert or update account token, preferring outlook_rest and later expiration."""
    if not info:
        return
    upn = info['upn']
    # Case-insensitive key lookup so Antonio.Maritato@... and antonio.maritato@... merge
    existing_key = next((k for k in accounts if k.casefold() == upn.casefold()), None)
    existing = accounts.get(existing_key) if existing_key else None
    replace = False
    if not existing:
        replace = True
    elif existing['endpoint_type'] == 'graph' and info['endpoint_type'] == 'outlook_rest':
        replace = True
    elif existing['endpoint_type'] == info['endpoint_type'] and info['exp'] > existing['exp']:
        replace = True

    if replace:
        target_key = existing_key or upn
        accounts[target_key] = info


# ---------------------------------------------------------------- Remote OE Copilot Harvest

_REMOTE_HARVEST_SCRIPT = r"""
import base64, json, os, shutil, sqlite3, sys, tempfile
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

base = sys.argv[1] if len(sys.argv) > 1 else '/opt/oe-copilot'
ks_path = os.path.join(base, 'telefono/data/misc/keystore/persistent.sqlite')
acct_path = os.path.join(base, 'telefono/data/data/com.microsoft.office.outlook/databases/acompliAcct.db')
accessi_path = os.path.join(base, 'stato/accessi.json')

out = {'tokens': [], 'accessi': {}}
if os.path.exists(accessi_path):
    try:
        out['accessi'] = json.load(open(accessi_path)).get('caselle', {})
    except Exception:
        pass

if not (os.path.exists(ks_path) and os.path.exists(acct_path)):
    print(json.dumps(out))
    sys.exit(0)

tmp = tempfile.mkdtemp(prefix='oe-mail-mcp-')
try:
    for src, name in ((ks_path, 'ks.db'), (acct_path, 'acct.db')):
        shutil.copy2(src, os.path.join(tmp, name))
        for suf in ('-wal', '-shm', '-journal'):
            if os.path.exists(src + suf):
                shutil.copy2(src + suf, os.path.join(tmp, name + suf))

    cks = sqlite3.connect(os.path.join(tmp, 'ks.db'))
    row = cks.execute(
        "SELECT b.blob FROM blobentry b JOIN keyentry k ON b.keyentryid = k.id "
        "WHERE k.alias = 'HxCoreSecureStringAES256Gcm' AND b.subcomponent_type = 0"
    ).fetchone()
    cks.close()
    if not row:
        print(json.dumps(out))
        sys.exit(0)

    key = row[0][13:45]
    aes = AESGCM(key)

    def dec(enc):
        if not enc:
            return None
        raw = base64.urlsafe_b64decode(enc + '=' * (-len(enc) % 4))
        if len(raw) < 33:
            return None
        return aes.decrypt(raw[5:17], raw[17:], None).decode('utf-8')

    cac = sqlite3.connect(os.path.join(tmp, 'acct.db'))
    for (tok_enc,) in cac.execute(
        "SELECT token FROM tokenStore WHERE token IS NOT NULL AND "
        "(resource LIKE '%outlook%' OR resource LIKE '%graph%')"
    ):
        try:
            jwt = dec(tok_enc)
            if jwt:
                out['tokens'].append(jwt)
        except Exception:
            pass

    for (dtok_enc,) in cac.execute(
        "SELECT directToken FROM mailAccounts WHERE directToken IS NOT NULL"
    ):
        try:
            jwt = dec(dtok_enc)
            if jwt:
                out['tokens'].append(jwt)
        except Exception:
            pass
    cac.close()
finally:
    shutil.rmtree(tmp, ignore_errors=True)

print(json.dumps(out))
"""


def harvest_remote_tokens(now: float = None) -> dict:
    """Harvest Outlook tokens from the OE Copilot virtual phone over SSH."""
    if os.environ.get('OUTLOOK_MCP_NO_REMOTE', '').lower() in ('1', 'true', 'yes'):
        return {}
    if now is None:
        now = time.time()
    accounts = {}
    try:
        proc = subprocess.run(
            ['ssh', '-o', 'BatchMode=yes', '-o', 'ConnectTimeout=8', SSH_HOST,
             'python3', '-', REMOTE_DIR],
            input=_REMOTE_HARVEST_SCRIPT,
            capture_output=True,
            text=True,
            encoding='utf-8',
            timeout=15
        )
        if proc.returncode != 0 or not proc.stdout.strip():
            return {}
        payload = json.loads(proc.stdout.strip().splitlines()[-1])
    except Exception:
        return {}

    accessi = payload.get('accessi') or {}
    for secret in payload.get('tokens') or []:
        info = _classify_token(secret, now, source='oe-copilot')
        if not info:
            continue
        acc_status = accessi.get(info['upn'].lower()) or {}
        if acc_status:
            info['phone_status'] = acc_status.get('stato', 'ok')
            info['phone_msg'] = acc_status.get('msg', '')
        _merge_account(accounts, info)

    return accounts


# ---------------------------------------------------------------- Authenticated OE Copilot broker

class _NoBrokerRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        # Never forward the broker credential to a redirect target.
        return None


def harvest_broker_tokens(now: float = None) -> dict:
    """Read the optional authenticated agent /api/token-posta endpoint.

    Invalid configuration, unavailable brokers and invalid rows fail closed for
    this source only. Response bodies and credentials are never logged.
    """
    url = os.environ.get('OUTLOOK_MCP_BROKER_URL', '')
    key = os.environ.get('OUTLOOK_MCP_BROKER_KEY', '')
    if not url or not key:
        return {}
    if now is None:
        now = time.time()
    try:
        parsed = urllib.parse.urlsplit(url)
        if (parsed.scheme not in ('https', 'http') or not parsed.hostname
                or parsed.username or parsed.password or parsed.query or parsed.fragment):
            return {}
        if parsed.scheme == 'http' and parsed.hostname not in ('127.0.0.1', '::1', 'localhost'):
            return {}
        req = urllib.request.Request(url, headers={
            'Authorization': 'Bearer ' + key, 'Accept': 'application/json'
        })
        opener = urllib.request.build_opener(_NoBrokerRedirect())
        with opener.open(req, timeout=10) as response:
            payload = json.load(response)
        rows = payload.get('account', [])
        if not isinstance(rows, list):
            return {}
    except urllib.error.HTTPError as error:
        error.close()
        return {}
    except Exception:
        return {}

    allowed = {u.strip().casefold() for u in
               os.environ.get('OUTLOOK_MCP_BROKER_ACCOUNTS', '').split(',') if u.strip()}
    accounts = {}
    for row in rows:
        try:
            if not isinstance(row, dict) or not isinstance(row.get('token'), str):
                continue
            info = _classify_token(row['token'], now + 120, source='oe-copilot-broker')
            if not info or not isinstance(row.get('upn'), str):
                continue
            if row['upn'].casefold() != info['upn'].casefold():
                continue
            if allowed and info['upn'].casefold() not in allowed:
                continue
            # Derive audience, scopes and API base from JWT, never broker URLs.
            info['exp'] = min(info['exp'], int(row.get('exp', 0)))
            if info['exp'] <= now + 120:
                continue
            _merge_account(accounts, info)
        except (TypeError, ValueError, OverflowError):
            continue
    return accounts


# ---------------------------------------------------------------- Local Harvest

def harvest_local_tokens(now: float = None) -> dict:
    """Scan local Windows IdentityCache and decrypt tokens via DPAPI."""
    if now is None:
        now = time.time()
    id_cache = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'IdentityCache')
    accounts = {}
    if not os.path.exists(id_cache):
        return accounts

    for root, dirs, files in os.walk(id_cache):
        if '\\AT' not in root:
            continue
        for f in files:
            if not f.endswith('.bin'):
                continue
            p = os.path.join(root, f)
            try:
                with open(p, 'rb') as fp:
                    raw = fp.read()
                dec = dpapi_unprotect(raw)
                data = json.loads(dec.decode('utf-8'))
            except Exception:
                continue

            for k, v in data.items():
                secret = v.get('secret')
                info = _classify_token(secret, now, source='local')
                if info:
                    _merge_account(accounts, info)

    return accounts


def harvest_tokens() -> dict:
    """Harvest local, SSH phone and optional authenticated broker tokens."""
    now = time.time()
    accounts = harvest_local_tokens(now)
    from web_session import harvest as harvest_web_tokens
    for source in (harvest_remote_tokens, harvest_broker_tokens, harvest_web_tokens):
        for info in source(now).values():
            _merge_account(accounts, info)

    if accounts:
        try:
            os.makedirs(CACHE_DIR, exist_ok=True)
            state = {
                'harvested_at': int(now),
                'accounts': accounts
            }
            tmp = STATE_FILE + '.tmp'
            with open(tmp, 'wb') as fp:
                fp.write(dpapi_protect(json.dumps(state).encode('utf-8')))
            os.replace(tmp, STATE_FILE)
        except Exception:
            pass

    return accounts


def load_cached_tokens() -> dict:
    """Load cached tokens from STATE_FILE if still valid (with 120s safety margin)."""
    now = time.time()
    try:
        with open(STATE_FILE, 'rb') as fp:
            dec = dpapi_unprotect(fp.read())
        state = json.loads(dec.decode('utf-8'))
        accounts = state.get('accounts', {})
        valid = {u: a for u, a in accounts.items() if a.get('exp', 0) > now + 120}
        return valid
    except Exception:
        return {}


def get_accounts(refresh=False) -> dict:
    """Get all available accounts with valid tokens."""
    if not refresh:
        cached = load_cached_tokens()
        if cached:
            return cached
    accounts = harvest_tokens()
    if not accounts:
        accounts = load_cached_tokens()
    return accounts


def get_token_for_account(account=None) -> dict:
    """Return token info dict for specific or default account."""
    accts = get_accounts()
    if not accts:
        accts = harvest_tokens()
    if not accts:
        raise RuntimeError(
            'No active Outlook account tokens found (checked local IdentityCache and OE Copilot phone).'
        )

    if not account:
        upn = next(iter(accts))
        return accts[upn]

    acc_lower = account.casefold()
    for upn, info in accts.items():
        if acc_lower in upn.casefold():
            return info

    # If requested account wasn't in cache, force a fresh harvest once
    accts = harvest_tokens()
    for upn, info in accts.items():
        if acc_lower in upn.casefold():
            return info

    raise ValueError(f"Account '{account}' not found among active accounts: {', '.join(accts.keys())}")


if __name__ == '__main__':
    accts = get_accounts(refresh=True)
    print(f"Found {len(accts)} Outlook account(s):")
    for upn, info in accts.items():
        exp_min = int((info['exp'] - time.time()) / 60)
        src = info.get('source', 'local')
        pstat = f", phone: {info['phone_status']}" if info.get('phone_status') else ""
        print(f"  {upn}: {info['endpoint_type']} [{src}{pstat}], expires in {exp_min} min")
