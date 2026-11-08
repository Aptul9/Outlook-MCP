import unittest
from unittest.mock import patch
import outlook_api as api

class CalendarTests(unittest.TestCase):
    def test_shared_transport(self):
        with patch.object(api,'_request',return_value={'value':[{'Id':'one'}]}) as request:
            self.assertEqual(api.list_calendars('msc'),[{'Id':'one'}])
            request.assert_called_once_with('msc','me/calendars?$top=100')
    def test_truncation_and_quoted_id(self):
        with patch.object(api,'_request',return_value={'value':[{}],'@odata.nextLink':'https://untrusted.example'}) as request:
            result=api.list_calendar_events('msc','2026-11-07T00:00:00Z','2026-11-08T00:00:00Z','a/b',5000)
            self.assertTrue(result['has_more']);self.assertNotIn('nextLink',result)
            self.assertIn('a%2Fb',request.call_args.args[1]);self.assertIn('%24top=1000',request.call_args.args[1])
    def test_invalid_ranges(self):
        for start,end in [('2026-11-07','2026-11-08'),('2026-11-08T00:00:00Z','2026-11-07T00:00:00Z'),('2020-01-01T00:00:00Z','2026-01-01T00:00:00Z')]:
            with self.assertRaises(ValueError):api.list_calendar_events('msc',start,end)

if __name__=='__main__':unittest.main()
