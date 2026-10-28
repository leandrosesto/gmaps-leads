import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import gmaps_leads as g


class WhatsappLinkTests(unittest.TestCase):
    def test_inserts_mobile_nine(self):
        self.assertEqual(g.whatsapp_link("+54 223 456-7890"), "https://wa.me/5492234567890")

    def test_keeps_existing_nine(self):
        self.assertEqual(g.whatsapp_link("+54 9 223 456-7890"), "https://wa.me/5492234567890")

    def test_empty(self):
        self.assertEqual(g.whatsapp_link(None), "")


class SearchTests(unittest.TestCase):
    def _page(self, ids, token=None):
        data = {"places": [{"id": i, "displayName": {"text": f"N{i}"}, "nationalPhoneNumber": "0223 1"} for i in ids]}
        if token:
            data["nextPageToken"] = token
        resp = MagicMock(status_code=200)
        resp.json.return_value = data
        return resp

    @patch("gmaps_leads.time.sleep")
    def test_pagination_and_limit(self, _sleep):
        session = MagicMock()
        session.post.side_effect = [self._page(["a", "b"], "t1"), self._page(["c", "d"])]
        rows = g.search("k", "q", "Mar del Plata", 3, session)
        self.assertEqual([r["id"] for r in rows], ["a", "b", "c"])

    @patch("gmaps_leads.search")
    def test_dedup_and_closed(self, search):
        base = {"Teléfono": "1"}
        search.return_value = [
            {**base, "id": "1", "_status": "OPERATIONAL"},
            {**base, "id": "1", "_status": "OPERATIONAL"},
            {**base, "id": "2", "_status": "CLOSED_PERMANENTLY"},
        ]
        rows = g.collect("k", ["q"], "c", 60, False)
        self.assertEqual([r["id"] for r in rows], ["1"])

    def test_api_error_message(self):
        session = MagicMock()
        resp = MagicMock(status_code=403, text="")
        resp.json.return_value = {"error": {"message": "API key not valid"}}
        session.post.return_value = resp
        with self.assertRaisesRegex(g.ApiError, "API key not valid"):
            g.search("k", "q", "c", 20, session)


if __name__ == "__main__":
    unittest.main()
