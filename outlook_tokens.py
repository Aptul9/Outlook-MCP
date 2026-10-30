"""Token harvester for New Outlook (olk.exe).

Extracts and decrypts OAuth2 bearer tokens cached by Microsoft IdentityBroker / OneAuth
in %LOCALAPPDATA%\\Microsoft\\IdentityCache.

Standard library only, Windows DPAPI via ctypes. No elevation required, no app restarts.
"""
import base64
import ctypes
import json
import os
import time
from ctypes import wintypes

CACHE_DIR = os.path.join(os.environ.get('LOCALAPPDATA', '.'), 'OutlookMCP')
STATE_FILE = os.path.join(CACHE_DIR, 'tokens.bin')

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


# ---------------------------------------------------------------- Harvest

def harvest_tokens() -> dict:
    """Scan IdentityCache, decrypt tokens, and return map of accounts:
    {
      "user@example.com": {
         "upn": "user@example.com",
         "token": "...",
         "exp": 123456789,
         "endpoint_type": "outlook_rest" | "graph",
         "base_url": "https://outlook.office.com/api/v2.0" | "https://graph.microsoft.com/v1.0",
         "scopes": "..."
      }
    }
    """
    id_cache = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'IdentityCache')
    now = time.time()
    accounts = {}

    if os.path.exists(id_cache):
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
                    if not secret or len(secret.split('.')) < 3:
                        continue
                    claims = jwt_claims(secret)
                    upn = claims.get('upn') or claims.get('email') or claims.get('unique_name')
                    if not upn:
                        continue
                    exp = claims.get('exp', 0)
                    if exp <= now:
                        continue

                    aud = str(claims.get('aud', ''))
                    scp = str(claims.get('scp', ''))

                    # Check suitability
                    endpoint_type = None
                    base_url = None

                    if 'outlook.office.com' in aud or 'outlook.office365.com' in aud:
                        if 'Mail.Read' in scp or 'Mail.ReadWrite' in scp or 'EAS.AccessAsUser' in scp or 'OWA.AccessAsUser' in scp:
                            endpoint_type = 'outlook_rest'
                            base_url = 'https://outlook.office.com/api/v2.0'
                    elif 'graph.microsoft.com' in aud or aud == '00000003-0000-0000-c000-000000000000':
                        if 'Mail.Read' in scp or 'Mail.ReadWrite' in scp:
                            endpoint_type = 'graph'
                            base_url = 'https://graph.microsoft.com/v1.0'

                    if endpoint_type:
                        # Prioritize outlook_rest over graph if both exist, or later expiry
                        existing = accounts.get(upn)
                        replace = False
                        if not existing:
                            replace = True
                        elif existing['endpoint_type'] == 'graph' and endpoint_type == 'outlook_rest':
                            replace = True
                        elif existing['endpoint_type'] == endpoint_type and exp > existing['exp']:
                            replace = True

                        if replace:
                            accounts[upn] = {
                                'upn': upn,
                                'token': secret,
                                'exp': int(exp),
                                'endpoint_type': endpoint_type,
                                'base_url': base_url,
                                'scopes': scp
                            }

    if accounts:
        # Cache to disk protected by DPAPI
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
    """Load cached tokens from STATE_FILE if still valid."""
    now = time.time()
    try:
        with open(STATE_FILE, 'rb') as fp:
            dec = dpapi_unprotect(fp.read())
        state = json.loads(dec.decode('utf-8'))
        accounts = state.get('accounts', {})
        # Filter expired
        valid = {u: a for u, a in accounts.items() if a.get('exp', 0) > now}
        return valid
    except Exception:
        return {}


def get_accounts(refresh=False) -> dict:
    """Get all available accounts with valid tokens."""
    if not refresh:
        cached = load_cached_tokens()
        if cached:
            return cached
    # Harvest from IdentityCache
    accounts = harvest_tokens()
    if not accounts:
        # Try cached even if near expiry
        accounts = load_cached_tokens()
    return accounts


def get_token_for_account(account=None) -> dict:
    """Return token info dict for specific or default account."""
    accts = get_accounts()
    if not accts:
        accts = harvest_tokens()
    if not accts:
        raise RuntimeError('No active Outlook account tokens found. Ensure New Outlook is opened and signed in.')

    if not account:
        # Pick primary / first account
        upn = next(iter(accts))
        return accts[upn]

    acc_lower = account.casefold()
    for upn, info in accts.items():
        if acc_lower in upn.casefold():
            return info

    raise ValueError(f"Account '{account}' not found among active accounts: {', '.join(accts.keys())}")


if __name__ == '__main__':
    accts = get_accounts(refresh=True)
    print(f"Found {len(accts)} Outlook accounts:")
    for upn, info in accts.items():
        exp_min = int((info['exp'] - time.time()) / 60)
        print(f"  {upn}: {info['endpoint_type']} ({info['base_url']}), expires in {exp_min} min")
