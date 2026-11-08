import base64
import contextlib
import io
import json
import os
import threading
import unittest
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from unittest.mock import patch

import outlook_tokens as tokens

NOW = 2000000000


def jwt(upn='user@example.test', exp=NOW + 600, **extra):
    claims = dict(upn=upn, exp=exp, aud='https://graph.microsoft.com', scp='Mail.Read')
    claims.update(extra)
    part = base64.urlsafe_b64encode(json.dumps(claims).encode()).decode().rstrip('=')
    return 'fake.' + part + '.unsigned'


def row(**kwargs):
    return dict(upn='user@example.test', exp=NOW + 600, token=jwt(), **kwargs)


class BrokerTests(unittest.TestCase):
    def setUp(self):
        self.payload = {'account': [row()]}
        self.status = 200
        self.redirect = None
        self.seen = []
        owner = self

        class Handler(BaseHTTPRequestHandler):
            def do_GET(self):
                owner.seen.append((self.path, self.headers.get('Authorization')))
                self.send_response(owner.status)
                if owner.redirect:
                    self.send_header('Location', owner.redirect)
                self.end_headers()
                body = owner.payload
                self.wfile.write(body if isinstance(body, bytes) else json.dumps(body).encode())

            def log_message(self, *args):
                pass

        self.server = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
        self.thread = threading.Thread(target=self.server.serve_forever)
        self.thread.start()
        self.url = 'http://127.0.0.1:%s/api/token-posta' % self.server.server_port
        self.env = patch.dict(os.environ, {'OUTLOOK_MCP_BROKER_URL': self.url,
                                          'OUTLOOK_MCP_BROKER_KEY': 'fake-broker-key',
                                          'OUTLOOK_MCP_BROKER_ACCOUNTS': ''})
        self.env.start()

    def tearDown(self):
        self.env.stop()
        self.server.shutdown()
        self.thread.join()
        self.server.server_close()

    def fetch(self):
        return tokens.harvest_broker_tokens(NOW)

    def test_authenticated_request_and_untrusted_metadata(self):
        self.payload['account'][0].update(base_url='https://evil.test', endpoint_type='evil')
        info = self.fetch()['user@example.test']
        self.assertEqual(self.seen, [('/api/token-posta', 'Bearer fake-broker-key')])
        self.assertEqual(info['base_url'], 'https://graph.microsoft.com/v1.0')
        self.assertEqual(info['source'], 'oe-copilot-broker')

    def test_outlook_mail_and_calendar_scopes_use_existing_classifier(self):
        self.payload = {'account': [dict(
            upn='user@example.test', exp=NOW + 600,
            token=jwt(aud='https://outlook.office.com',
                      scp='Mail.ReadWrite Calendars.ReadWrite'))]}
        info = self.fetch()['user@example.test']
        self.assertEqual(info['endpoint_type'], 'outlook_rest')
        self.assertEqual(info['base_url'], 'https://outlook.office.com/api/v2.0')
        self.assertIn('Calendars.ReadWrite', info['scopes'])

    def test_expired_margin_and_row_expiry(self):
        for exp in (NOW - 1, NOW, NOW + 120):
            self.payload = {'account': [dict(upn='user@example.test', exp=exp, token=jwt(exp=exp))]}
            self.assertEqual(self.fetch(), {})
        self.payload = {'account': [dict(upn='user@example.test', exp=NOW, token=jwt())]}
        self.assertEqual(self.fetch(), {})

    def test_exact_case_insensitive_account_allowlist(self):
        with patch.dict(os.environ, {'OUTLOOK_MCP_BROKER_ACCOUNTS': 'USER@example.test'}):
            self.assertEqual(len(self.fetch()), 1)
        with patch.dict(os.environ, {'OUTLOOK_MCP_BROKER_ACCOUNTS': 'user'}):
            self.assertEqual(self.fetch(), {})

    def test_mismatched_identity_and_invalid_rows(self):
        self.payload = {'account': [None, {}, {'token': 42}, {'upn': 'user@example.test', 'token': jwt()}, dict(upn='other@example.test', token=jwt(), exp=NOW+600), dict(upn='user@example.test', token=jwt(exp='bad'), exp=NOW+600), row()]}
        self.assertEqual(len(self.fetch()), 1)

    def test_failure_does_not_log_credentials_or_response(self):
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            for status, body in ((401, b'fake-broker-key'), (500, jwt().encode()), (200, b'not json'), (200, {'account': None}), (200, [])):
                self.status, self.payload = status, body
                self.assertEqual(self.fetch(), {})
        self.assertEqual(out.getvalue() + err.getvalue(), '')

    def test_redirect_not_followed(self):
        self.status, self.redirect = 302, self.url + '/redirected'
        self.assertEqual(self.fetch(), {})
        self.assertEqual(len(self.seen), 1)

    def test_disabled_and_insecure_remote_configuration(self):
        for config in ({'OUTLOOK_MCP_BROKER_KEY': ''}, {'OUTLOOK_MCP_BROKER_URL': 'http://remote.test/api/token-posta'}, {'OUTLOOK_MCP_BROKER_URL': self.url + '?secret=x'}):
            with patch.dict(os.environ, config):
                self.assertEqual(self.fetch(), {})
        self.assertEqual(self.seen, [])

    def test_integrated_harvest_retains_sources_and_dpapi_cache(self):
        with patch.object(tokens, 'harvest_local_tokens', return_value={}) as local, patch.object(tokens, 'harvest_remote_tokens', return_value={}) as remote, patch.object(tokens.time, 'time', return_value=NOW), patch.object(tokens, 'dpapi_protect', side_effect=lambda data: b'encrypted'), patch('builtins.open', unittest.mock.mock_open()) as file, patch.object(tokens.os, 'makedirs'), patch.object(tokens.os, 'replace'):
            self.assertEqual(len(tokens.harvest_tokens()), 1)
            local.assert_called_once_with(NOW)
            remote.assert_called_once_with(NOW)
            file().write.assert_called_once_with(b'encrypted')

    def test_existing_account_selection(self):
        accounts = self.fetch()
        with patch.object(tokens, 'get_accounts', return_value=accounts):
            self.assertEqual(tokens.get_token_for_account('USER@')['upn'], 'user@example.test')


if __name__ == '__main__':
    unittest.main()
