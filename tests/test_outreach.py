import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import gmaps_leads as g
import outreach as o


class RenderTests(unittest.TestCase):
    T = "Hola! [[Vi {detalle}. ]]¿En {negocio} cómo lo hacen?"

    def test_with_detail(self):
        self.assertEqual(o.render_message(self.T, "Kine X", "su web"), "Hola! Vi su web. ¿En Kine X cómo lo hacen?")

    def test_without_detail_drops_block(self):
        self.assertEqual(o.render_message(self.T, "Kine X", "  "), "Hola! ¿En Kine X cómo lo hacen?")


class TemplateTests(unittest.TestCase):
    def test_splits_variants_and_ignores_comments(self):
        with tempfile.TemporaryDirectory() as tmp:
            f = Path(tmp) / "m.txt"
            f.write_text("# nota\nUno\n---\nDos\n", encoding="utf-8")
            self.assertEqual(o.load_templates(f), ["Uno", "Dos"])

    def test_empty_raises(self):
        with tempfile.TemporaryDirectory() as tmp:
            f = Path(tmp) / "m.txt"
            f.write_text("# nada\n", encoding="utf-8")
            with self.assertRaises(ValueError):
                o.load_templates(f)


class RunTests(unittest.TestCase):
    def _sheet(self, tmp: Path) -> Path:
        base = {c: "" for c in g.COLUMNS}
        rows = [
            {**base, "Nombre": "Con IG", "Instagram": "https://www.instagram.com/conig/"},
            {**base, "Nombre": "Sin IG"},
            {**base, "Nombre": "Ya contactado", "Instagram": "https://www.instagram.com/ya/", "Contactado": "Sí"},
        ]
        path = tmp / "leads_test.xlsx"
        g.write_xlsx(rows, path)
        return path

    def test_marks_only_confirmed_lead_and_keeps_link(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._sheet(Path(tmp))
            answers = iter(["mi detalle", "", "s"])  # detalle, Enter para abrir, ¿enviaste? sí
            with patch("builtins.input", lambda *_: next(answers)), patch("outreach.webbrowser.open") as op, patch(
                "outreach.copy_to_clipboard", return_value=True
            ):
                self.assertEqual(o.run(path, ["Hola {negocio}"], 15), 0)
            op.assert_called_once_with("https://www.instagram.com/conig/")
            ws = load_workbook(path).active
            header = [c.value for c in ws[1]]
            row = {h: ws.cell(row=2, column=i + 1) for i, h in enumerate(header)}
            self.assertEqual(row["Contactado"].value, "Sí")
            self.assertEqual(row["Fecha contacto"].value, date.today().isoformat())
            self.assertEqual(row["Notas"].value, "Variante 1")
            self.assertEqual(row["Instagram"].hyperlink.target, "https://www.instagram.com/conig/")

    def test_not_sent_leaves_lead_pending(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._sheet(Path(tmp))
            answers = iter(["", "", "n", "q"])
            with patch("builtins.input", lambda *_: next(answers)), patch("outreach.webbrowser.open"), patch(
                "outreach.copy_to_clipboard", return_value=True
            ):
                o.run(path, ["Hola"], 15)
            ws = load_workbook(path).active
            self.assertFalse(ws.cell(row=2, column=[c.value for c in ws[1]].index("Contactado") + 1).value)

    def test_daily_cap_stops_before_prompting(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._sheet(Path(tmp))
            with patch("builtins.input", side_effect=AssertionError("no debía preguntar")):
                self.assertEqual(o.run(path, ["Hola"], 0), 0)


if __name__ == "__main__":
    unittest.main()
