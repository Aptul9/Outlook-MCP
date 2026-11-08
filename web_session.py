"""Outlook web authorization from an existing Microsoft browser session.

No browser control or tenant policy changes. Session format matches TeamsRelay.
The caller owns encrypted persistence and must provide the expected mailbox.
"""
import base64
import hashlib
import json
import secrets
import time
import urllib.error
import urllib.parse
import urllib.request

CLIENT_ID = '9199bf20-a13f-4107-85dc-02114787ef48'
REDIRECT = 'https://outlook.cloud.microsoft/mail/'
SCOPE = 'https://outlook.office.com/.default offline_access openid profile'
LOGIN = 'https://login.microsoftonline.com'


class NeedsSignIn(RuntimeError):
    pass


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, *args):
        return None


def request(url, headers, data=None):
    try:
        with urllib.request.build_opener(NoRedirect).open(
                urllib.request.Request(url, headers=headers, data=data), timeout=25) as response:
            return response.status, response.headers, response.read(2_000_000)
    except urllib.error.HTTPError as error:
        return error.code, error.headers, error.read(2_000_000)
    except (OSError, ValueError):
        raise RuntimeError('Microsoft authorization request failed') from None


def claims(token):
    try:
        return json.loads(base64.urlsafe_b64decode(token.split('.')[1] + '==='))
    except (ValueError, IndexError, TypeError):
        return {}


def validate(result, account):
    c = claims(result.get('access_token', ''))
    identity = c.get('upn') or c.get('preferred_username') or c.get('unique_name')
    if not identity or identity.casefold() != account.casefold():
        raise NeedsSignIn('Microsoft returned a different mailbox')
    if c.get('aud') not in ('https://outlook.office.com', 'https://outlook.office365.com'):
        raise NeedsSignIn('Microsoft returned an unexpected resource')
    if int(c.get('exp', 0)) <= time.time() + 120:
        raise NeedsSignIn('Microsoft returned an expired token')
    scopes = set(str(c.get('scp', '')).split())
    if not scopes.intersection(('Mail.Read', 'Mail.ReadWrite')):
        raise NeedsSignIn('Microsoft token lacks mail permission')
    return result


def token_response(status, body, account):
    if status == 429 or status >= 500:
        raise RuntimeError('Microsoft token service temporarily unavailable')
    try:
        result = json.loads(body)
    except (ValueError, TypeError):
        raise RuntimeError('Microsoft token response invalid') from None
    if status != 200 or not result.get('access_token'):
        code = result.get('error', '')
        # Do not emit descriptions, URLs, codes or response bodies containing credentials.
        if code in ('invalid_grant', 'interaction_required', 'login_required', 'consent_required'):
            raise NeedsSignIn('Microsoft requires interactive Outlook sign-in')
        raise RuntimeError('Microsoft refused Outlook token request')
    return validate(result, account)


def authorize(session, tenant, account, transport=request):
    if not tenant or any(c not in 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-' for c in tenant):
        raise ValueError('Invalid Microsoft tenant')
    if session.get('loginHint', '').casefold() != account.casefold():
        raise ValueError('Browser session mailbox mismatch')
    user_agent = session.get('userAgent', '')
    if not user_agent or '\r' in user_agent or '\n' in user_agent:
        raise ValueError('Invalid browser User-Agent')
    now = time.time()
    cookies = []
    endpoint = LOGIN + '/' + tenant + '/oauth2/v2.0/'
    authorize_path = urllib.parse.urlparse(endpoint + 'authorize').path
    for c in session.get('cookies', []):
        if c.get('domain', '').lstrip('.').lower() != 'login.microsoftonline.com':
            continue
        path = c.get('path', '/')
        if not (authorize_path == path or authorize_path.startswith(path.rstrip('/') + '/')):
            continue
        if c.get('expires', -1) != -1 and c.get('expires', 0) <= now:
            continue
        name, value = c.get('name', ''), c.get('value', '')
        if not name or any(ord(x) <= 32 or ord(x) >= 127 or x in ';=' for x in name):
            continue
        if any(ord(x) <= 32 or ord(x) >= 127 or x == ';' for x in value):
            continue
        cookies.append((name, value))
    if not any(name in ('ESTSAUTH', 'ESTSAUTHPERSISTENT', 'ESTSAUTHLIGHT') for name, _ in cookies):
        raise NeedsSignIn('Microsoft browser session expired')
    verifier, state = secrets.token_urlsafe(32), secrets.token_urlsafe(24)
    params = dict(client_id=CLIENT_ID, redirect_uri=REDIRECT, response_type='code',
                  response_mode='query', scope=SCOPE, prompt='none', login_hint=account,
                  state=state, code_challenge_method='S256',
                  code_challenge=base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).decode().rstrip('='))
    status, headers, _ = transport(endpoint + 'authorize?' + urllib.parse.urlencode(params),
                                   {'User-Agent': user_agent, 'Cookie': '; '.join(n + '=' + v for n, v in cookies)})
    if status == 429 or status >= 500:
        raise RuntimeError('Microsoft authorization temporarily unavailable')
    back = urllib.parse.urlparse(headers.get('Location', ''))
    if status not in (302, 303) or back.scheme + '://' + back.netloc + back.path != REDIRECT:
        raise NeedsSignIn('Microsoft requires interactive Outlook sign-in')
    query = urllib.parse.parse_qs(back.query)
    if query.get('state') != [state]:
        raise RuntimeError('Microsoft authorization state mismatch')
    if not query.get('code') or query.get('error'):
        raise NeedsSignIn('Microsoft requires interactive Outlook sign-in')
    body = urllib.parse.urlencode(dict(client_id=CLIENT_ID, redirect_uri=REDIRECT,
             grant_type='authorization_code', code=query['code'][0], code_verifier=verifier, scope=SCOPE)).encode()
    status, _, raw = transport(endpoint + 'token', {'User-Agent': user_agent,
                  'Origin': 'https://outlook.cloud.microsoft', 'Content-Type': 'application/x-www-form-urlencoded'}, body)
    return token_response(status, raw, account)


def refresh(token, tenant, account, transport=request):
    if not tenant or any(c not in 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-' for c in tenant):
        raise ValueError('Invalid Microsoft tenant')
    body = urllib.parse.urlencode(dict(client_id=CLIENT_ID, grant_type='refresh_token',
                  refresh_token=token, scope=SCOPE)).encode()
    status, _, raw = transport(LOGIN + '/' + tenant + '/oauth2/v2.0/token',
                {'Origin': 'https://outlook.cloud.microsoft', 'Content-Type': 'application/x-www-form-urlencoded'}, body)
    return token_response(status, raw, account)


def harvest(now=None):
    """Optional web-session source; tokens remain in the MCP DPAPI cache.

    OUTLOOK_MCP_WEB_SESSION_FILE points to a DPAPI-encrypted TeamsRelay-format
    session, including tenantId. No browser or production session is modified.
    """
    import os
    import outlook_tokens as tokens
    path = os.environ.get('OUTLOOK_MCP_WEB_SESSION_FILE', '')
    if not path:
        return {}
    try:
        with open(path, 'rb') as source:
            session = json.loads(tokens.dpapi_unprotect(source.read()))
        account = session['loginHint']
        result = authorize(session, session['tenantId'], account)
        info = tokens._classify_token(result['access_token'], time.time() if now is None else now, source='web_session')
        return {info['upn']: info} if info else {}
    except (OSError, ValueError, KeyError, RuntimeError):
        return {}

if __name__ == '__main__':
    import argparse
    import os
    import sys
    parser = argparse.ArgumentParser(description='Import a Microsoft browser session into the Outlook MCP encrypted store')
    parser.add_argument('--tenant', required=True)
    parser.add_argument('--output', required=True)
    args = parser.parse_args()
    import outlook_tokens as tokens
    session = json.load(sys.stdin)
    session['tenantId'] = args.tenant
    # Prove Outlook access before persisting credentials.
    authorize(session, args.tenant, session['loginHint'])
    parent = os.path.dirname(os.path.abspath(args.output))
    os.makedirs(parent, exist_ok=True)
    temp = args.output + '.new'
    with open(temp, 'wb') as destination:
        destination.write(tokens.dpapi_protect(json.dumps(session).encode()))
    os.replace(temp, args.output)
    print('Outlook session authorized and stored encrypted')
