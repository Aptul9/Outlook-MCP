import base64
import json
import time
import unittest
import urllib.parse
import web_session as w

ACCOUNT='msc@example.com'
def token(**overrides):
    c=dict(upn=ACCOUNT,aud='https://outlook.office.com',exp=int(time.time())+3600,scp='Mail.ReadWrite Calendars.ReadWrite')
    c.update(overrides)
    return 'a.'+base64.urlsafe_b64encode(json.dumps(c).encode()).decode().rstrip('=')+'.c'

class Tests(unittest.TestCase):
    def session(self):
        return dict(loginHint=ACCOUNT,userAgent='Chrome',cookies=[dict(name='ESTSAUTH',value='private',domain='.login.microsoftonline.com',path='/',expires=-1)])
    def transport(self,url,headers,data=None):
        if '/authorize?' in url:
            p=urllib.parse.parse_qs(urllib.parse.urlparse(url).query)
            self.assertEqual(p['client_id'],[w.CLIENT_ID])
            self.assertEqual(headers['Cookie'],'ESTSAUTH=private')
            return 302,{'Location':w.REDIRECT+'?'+urllib.parse.urlencode(dict(code='secret',state=p['state'][0]))},b''
        self.assertEqual(headers['Origin'],'https://outlook.cloud.microsoft')
        return 200,{},json.dumps(dict(access_token=token(),refresh_token='rotated')).encode()
    def test_authorize(self):
        self.assertEqual(w.authorize(self.session(),'tenant',ACCOUNT,self.transport)['refresh_token'],'rotated')
    def test_refresh(self):
        self.assertEqual(w.refresh('old','tenant',ACCOUNT,self.transport)['refresh_token'],'rotated')
    def test_wrong_identity(self):
        with self.assertRaises(w.NeedsSignIn):w.validate(dict(access_token=token(upn='other@example.com')),ACCOUNT)
    def test_wrong_audience(self):
        with self.assertRaises(w.NeedsSignIn):w.validate(dict(access_token=token(aud='https://graph.microsoft.com')),ACCOUNT)
    def test_expired(self):
        with self.assertRaises(w.NeedsSignIn):w.validate(dict(access_token=token(exp=0)),ACCOUNT)
    def test_missing_mail(self):
        with self.assertRaises(w.NeedsSignIn):w.validate(dict(access_token=token(scp='Calendars.Read')),ACCOUNT)
    def test_bad_state(self):
        with self.assertRaisesRegex(RuntimeError,'state mismatch'):
            w.authorize(self.session(),'tenant',ACCOUNT,lambda *a:(302,{'Location':w.REDIRECT+'?code=secret&state=bad'},b''))
    def test_bad_redirect(self):
        with self.assertRaises(w.NeedsSignIn):
            w.authorize(self.session(),'tenant',ACCOUNT,lambda *a:(302,{'Location':'https://evil.example/?code=secret'},b''))
    def test_ca_denial(self):
        def deny(url,*args):
            p=urllib.parse.parse_qs(urllib.parse.urlparse(url).query)
            return 302,{'Location':w.REDIRECT+'?'+urllib.parse.urlencode(dict(error='interaction_required',state=p['state'][0]))},b''
        with self.assertRaises(w.NeedsSignIn):w.authorize(self.session(),'tenant',ACCOUNT,deny)
    def test_no_cookie(self):
        s=self.session();s['cookies']=[]
        with self.assertRaises(w.NeedsSignIn):w.authorize(s,'tenant',ACCOUNT,self.transport)
    def test_cookie_host_scoping(self):
        s=self.session();s['cookies'].append(dict(name='evil',value='private',domain='other.example',path='/',expires=-1))
        w.authorize(s,'tenant',ACCOUNT,self.transport)
    def test_tenant_injection(self):
        with self.assertRaises(ValueError):w.refresh('old','../common',ACCOUNT,self.transport)
    def test_no_secret_in_error(self):
        with self.assertRaises(RuntimeError) as e:w.token_response(400,b'{"error_description":"secret"}',ACCOUNT)
        self.assertNotIn('secret',str(e.exception))
    def test_transient(self):
        with self.assertRaises(RuntimeError) as e:w.token_response(503,b'',ACCOUNT)
        self.assertNotIsInstance(e.exception,w.NeedsSignIn)

if __name__=='__main__':unittest.main()
