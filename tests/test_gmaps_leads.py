import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import gmaps_leads as g


class WhatsappLinkTests(unittest.TestCase):
    def test_landline_style_number(self):
        self.assertEqual(g.whatsapp_link("0223 667-6854"), "https://wa.me/5492236676854")

    def test_number_with_15(self):
        self.assertEqual(g.whatsapp_link("0223 15-567-1515"), "https://wa.me/5492235671515")

    def test_number_without_leading_zero(self):
        self.assertEqual(g.whatsapp_link("223 15 567 1515"), "https://wa.me/5492235671515")

    def test_international_format(self):
        self.assertEqual(g.whatsapp_link("+54 223 456-7890"), "https://wa.me/5492234567890")
        self.assertEqual(g.whatsapp_link("+54 9 223 456-7890"), "https://wa.me/5492234567890")

    def test_invalid_or_empty(self):
        self.assertEqual(g.whatsapp_link(""), "")
        self.assertEqual(g.whatsapp_link("0800 222 3333 44"), "")


class CleanTextTests(unittest.TestCase):
    def test_strips_icon_prefix(self):
        self.assertEqual(g._clean_button_text(" 0223 667-6854"), "0223 667-6854")


class ShortUrlTests(unittest.TestCase):
    def test_uses_place_id(self):
        url = "https://www.google.com/maps/place/Kinesio+ARM/data=!4m7!3m6!19sChIJwRwYCczfhJURlZZDo8DPFXg?authuser=0&hl=es"
        self.assertEqual(
            g.short_maps_url(url),
            "https://www.google.com/maps/place/?q=place_id:ChIJwRwYCczfhJURlZZDo8DPFXg",
        )

    def test_fallback_strips_query(self):
        self.assertEqual(g.short_maps_url("https://www.google.com/maps/place/X?hl=es"), "https://www.google.com/maps/place/X")


class ExtractInstagramTests(unittest.TestCase):
    def test_picks_most_frequent_profile(self):
        hrefs = [
            "https://www.instagram.com/kine.mdq/",
            "https://instagram.com/kine.mdq?hl=es",
            "https://www.instagram.com/otro/",
            "https://www.instagram.com/p/ABC123/",
            "https://www.instagram.com/sharer/sharer.php",
        ]
        self.assertEqual(g.extract_instagram(hrefs), "https://www.instagram.com/kine.mdq/")

    def test_no_profile(self):
        self.assertEqual(g.extract_instagram(["https://www.instagram.com/p/XYZ/", "https://www.instagram.com/explore/"]), "")
        self.assertEqual(g.extract_instagram([]), "")


class LinkLabelTests(unittest.TestCase):
    def test_instagram_handle(self):
        self.assertEqual(g.instagram_handle("http://www.instagram.com/kinesiakinesiologiamdq"), "@kinesiakinesiologiamdq")
        self.assertEqual(g.instagram_handle("https://www.instagram.com/fballesteros.kine/?hl=es"), "@fballesteros.kine")

    def test_domain_label(self):
        self.assertEqual(g.domain_label("https://www.sanarkinesiologia.com.ar/contacto"), "sanarkinesiologia.com.ar")


class OutputTests(unittest.TestCase):
    def test_writes_xlsx_and_csv(self):
        import tempfile

        row = {c: "" for c in g.COLUMNS}
        row["Nombre"] = "Test"
        with tempfile.TemporaryDirectory() as tmp:
            g.write_xlsx([row], Path(tmp) / "a.xlsx")
            g.write_csv([row], Path(tmp) / "a.csv")
            self.assertTrue((Path(tmp) / "a.xlsx").stat().st_size > 0)
            self.assertIn("Test", (Path(tmp) / "a.csv").read_text(encoding="utf-8-sig"))


if __name__ == "__main__":
    unittest.main()
