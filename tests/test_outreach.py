import io
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import MagicMock, patch

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import drive_upload as d
import gmaps_leads as g
import outreach as o


def leads_xlsx(path_or_none=None, names=("A", "B")) -> bytes:
    base = {c: "" for c in g.COLUMNS}
    rows = [
        {**base, "Nombre": n, "Dirección": f"Calle {n}", "Instagram": f"https://www.instagram.com/{n.lower()}/"}
        for n in names
    ]
    with tempfile.TemporaryDirectory() as tmp:
        p = Path(tmp) / "x.xlsx"
        g.write_xlsx(rows, p)
        return p.read_bytes()


def cell_values(data: bytes, row: int) -> dict:
    ws = load_workbook(io.BytesIO(data)).active
    header = [c.value for c in ws[1]]
    return {h: ws.cell(row=row, column=i + 1) for i, h in enumerate(header)}


class FakeDrive:
    """Sesión falsa de Drive con archivos en memoria."""

    def __init__(self, files):
        self.files = files  # id -> {"name": str, "data": bytes}
        self.patches = 0

    def get(self, url, params=None):
        r = MagicMock(status_code=200)
        if url == d.FILES_URL:
            r.json.return_value = {"files": [{"id": i, "name": f["name"], "modifiedTime": "2026-10-28T00:00:00Z"} for i, f in self.files.items()]}
        else:
            r.content = self.files[url.rsplit("/", 1)[1]]["data"]
        return r

    def patch(self, url, params=None, data=None, headers=None):
        self.files[url.rsplit("/", 1)[1]]["data"] = data
        self.patches += 1
        r = MagicMock(status_code=200)
        r.json.return_value = {"id": url.rsplit("/", 1)[1]}
        return r


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


class DriveStoreTests(unittest.TestCase):
    def _store(self, files, current="k1"):
        session = FakeDrive(files)
        listing = [{"id": i, "name": f["name"]} for i, f in files.items()]
        return session, o.DriveStore(session, listing[0] if current == "k1" else listing[1], listing)

    def test_mark_updates_only_that_lead_and_keeps_hyperlink(self):
        session, store = self._store({"k1": {"name": "leads_kine.xlsx", "data": leads_xlsx()}})
        self.assertTrue(store.mark(("b", "calle b"), "Variante 2"))
        data = session.files["k1"]["data"]
        a, b = cell_values(data, 2), cell_values(data, 3)
        self.assertEqual(b["Contactado"].value, "Sí")
        self.assertEqual(b["Notas"].value, "Variante 2")
        self.assertEqual(b["Fecha contacto"].value, date.today().isoformat())
        self.assertFalse(a["Contactado"].value)
        self.assertEqual(b["Instagram"].hyperlink.target, "https://www.instagram.com/b/")

    def test_mark_applies_to_latest_version_even_if_rows_moved(self):
        files = {"k1": {"name": "leads_kine.xlsx", "data": leads_xlsx(names=("A", "B"))}}
        session, store = self._store(files)
        files["k1"]["data"] = leads_xlsx(names=("Z", "B", "A"))  # alguien reordenó/agregó filas en Drive
        self.assertTrue(store.mark(("a", "calle a"), "Variante 1"))
        data = session.files["k1"]["data"]
        self.assertEqual(cell_values(data, 4)["Contactado"].value, "Sí")  # 'A' ahora es la fila 4
        self.assertFalse(cell_values(data, 2)["Contactado"].value)

    def test_mark_returns_false_if_lead_missing(self):
        session, store = self._store({"k1": {"name": "leads_kine.xlsx", "data": leads_xlsx()}})
        self.assertFalse(store.mark(("nadie", "x"), "Variante 1"))
        self.assertEqual(session.patches, 0)

    def test_notes_are_appended_not_overwritten(self):
        wb = load_workbook(io.BytesIO(leads_xlsx()))
        ws = wb.active
        cols = o.header_columns(ws)
        ws.cell(row=2, column=cols["Notas"]).value = "dueño amable"
        buf = io.BytesIO()
        wb.save(buf)
        session, store = self._store({"k1": {"name": "leads_kine.xlsx", "data": buf.getvalue()}})
        store.mark(("a", "calle a"), "Variante 1")
        self.assertEqual(cell_values(session.files["k1"]["data"], 2)["Notas"].value, "dueño amable | Variante 1")

    def test_daily_count_spans_all_sheets(self):
        today = date.today().isoformat()
        sheets = {}
        for fid, name in (("k1", "leads_kine.xlsx"), ("k2", "leads_contable.xlsx")):
            wb = load_workbook(io.BytesIO(leads_xlsx()))
            ws = wb.active
            cols = o.header_columns(ws)
            ws.cell(row=2, column=cols["Fecha contacto"]).value = today
            buf = io.BytesIO()
            wb.save(buf)
            sheets[fid] = {"name": name, "data": buf.getvalue()}
        _, store = self._store(sheets)
        self.assertEqual(store.contacted_today(), 2)


class RunTests(unittest.TestCase):
    def test_flow_marks_confirmed_lead_in_drive(self):
        files = {"k1": {"name": "leads_kine.xlsx", "data": leads_xlsx()}}
        session = FakeDrive(files)
        store = o.DriveStore(session, {"id": "k1", "name": "leads_kine.xlsx"}, [{"id": "k1", "name": "leads_kine.xlsx"}])
        answers = iter(["mi detalle", "", "s", "q"])  # lead A: detalle, abrir, enviado; lead B: salir
        with patch("builtins.input", lambda *_: next(answers)), patch("outreach.webbrowser.open") as op, patch(
            "outreach.copy_to_clipboard", return_value=True
        ):
            self.assertEqual(o.run(store, ["Hola {negocio}"], 15), 0)
        op.assert_called_once_with("https://www.instagram.com/a/")
        data = session.files["k1"]["data"]
        self.assertEqual(cell_values(data, 2)["Contactado"].value, "Sí")
        self.assertFalse(cell_values(data, 3)["Contactado"].value)

    def test_not_sent_leaves_lead_pending(self):
        files = {"k1": {"name": "leads_kine.xlsx", "data": leads_xlsx()}}
        session = FakeDrive(files)
        store = o.DriveStore(session, {"id": "k1", "name": "x"}, [{"id": "k1", "name": "x"}])
        answers = iter(["", "", "n", "q"])
        with patch("builtins.input", lambda *_: next(answers)), patch("outreach.webbrowser.open"), patch(
            "outreach.copy_to_clipboard", return_value=True
        ):
            o.run(store, ["Hola"], 15)
        self.assertEqual(session.patches, 0)

    def test_daily_cap_stops_before_prompting(self):
        files = {"k1": {"name": "leads_kine.xlsx", "data": leads_xlsx()}}
        store = o.DriveStore(FakeDrive(files), {"id": "k1", "name": "x"}, [{"id": "k1", "name": "x"}])
        with patch("builtins.input", side_effect=AssertionError("no debía preguntar")):
            self.assertEqual(o.run(store, ["Hola"], 0), 0)

    def test_failed_save_stops_the_run(self):
        files = {"k1": {"name": "leads_kine.xlsx", "data": leads_xlsx()}}
        store = o.DriveStore(FakeDrive(files), {"id": "k1", "name": "x"}, [{"id": "k1", "name": "x"}])
        answers = iter(["", "", "s"])
        with patch("builtins.input", lambda *_: next(answers)), patch("outreach.webbrowser.open"), patch(
            "outreach.copy_to_clipboard", return_value=True
        ), patch.object(o.DriveStore, "mark", return_value=False):
            self.assertEqual(o.run(store, ["Hola"], 15), 1)


class LocalStoreTests(unittest.TestCase):
    def test_local_mark(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "l.xlsx"
            p.write_bytes(leads_xlsx())
            self.assertTrue(o.LocalStore(p).mark(("a", "calle a"), "Variante 1"))
            self.assertEqual(cell_values(p.read_bytes(), 2)["Contactado"].value, "Sí")


class DriveListTests(unittest.TestCase):
    def test_list_download_update(self):
        session = FakeDrive({"k1": {"name": "leads_kine.xlsx", "data": b"OLD"}})
        self.assertEqual([f["name"] for f in d.list_sheets(session, "F")], ["leads_kine.xlsx"])
        self.assertEqual(d.download_file(session, "k1"), b"OLD")
        d.update_file(session, "k1", b"NEW")
        self.assertEqual(session.files["k1"]["data"], b"NEW")


if __name__ == "__main__":
    unittest.main()
