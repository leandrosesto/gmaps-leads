import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import drive_upload as d
import gmaps_leads as g


def resp(status=200, payload=None):
    r = MagicMock(status_code=status, text="")
    r.json.return_value = payload or {}
    return r


class FolderTests(unittest.TestCase):
    def test_reuses_existing_folder(self):
        session = MagicMock()
        session.get.return_value = resp(payload={"files": [{"id": "F1", "name": "LEADS"}]})
        self.assertEqual(d.find_or_create_folder(session, "LEADS"), "F1")
        session.post.assert_not_called()

    def test_creates_folder_when_missing(self):
        session = MagicMock()
        session.get.return_value = resp(payload={"files": []})
        session.post.return_value = resp(payload={"id": "NEW"})
        self.assertEqual(d.find_or_create_folder(session, "LEADS"), "NEW")
        self.assertEqual(session.post.call_args.kwargs["json"]["mimeType"], d.FOLDER_MIME)

    def test_quote_in_name_is_escaped(self):
        session = MagicMock()
        session.get.return_value = resp(payload={"files": [{"id": "X"}]})
        d.find_or_create_folder(session, "Mi 'carpeta'")
        self.assertIn("\\'carpeta\\'", session.get.call_args.kwargs["params"]["q"])

    def test_api_error_is_readable(self):
        session = MagicMock()
        session.get.return_value = resp(403, {"error": {"message": "Drive API disabled"}})
        with self.assertRaisesRegex(d.DriveError, "Drive API disabled"):
            d.find_or_create_folder(session)


class UploadTests(unittest.TestCase):
    def test_upload_sends_parent_and_content(self):
        session = MagicMock()
        session.post.return_value = resp(payload={"id": "A", "name": "x.xlsx", "webViewLink": "https://drive/x"})
        with tempfile.TemporaryDirectory() as tmp:
            f = Path(tmp) / "x.xlsx"
            f.write_bytes(b"DATA")
            out = d.upload_file(session, f, "F1")
        body = session.post.call_args.kwargs["data"]
        self.assertIn(b'"parents": ["F1"]', body)
        self.assertIn(b"DATA", body)
        self.assertEqual(out["webViewLink"], "https://drive/x")


class SlugTests(unittest.TestCase):
    def test_slugify(self):
        self.assertEqual(g.slugify("Kinesiólogo"), "kinesiologo")
        self.assertEqual(g.slugify("estudio contable"), "estudio-contable")
        self.assertEqual(g.slugify("???"), "leads")


if __name__ == "__main__":
    unittest.main()
