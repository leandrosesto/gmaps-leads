"""Modo asistido para contactar leads por Instagram. VOS enviás cada mensaje.

Trabaja sobre las planillas de tu carpeta LEADS de Google Drive (una por rubro) y
guarda el seguimiento ahí mismo, así lo ves desde cualquier PC. Por cada lead sin contactar:
  1. te muestra el negocio y el mensaje ya armado,
  2. abre su perfil de Instagram en tu navegador y copia el mensaje al portapapeles,
  3. espera a que vos lo pegues y lo envíes, y te pregunta si lo enviaste,
  4. marca el lead como contactado (fecha y variante) en la planilla de Drive.

No inicia sesión en Instagram ni envía nada por su cuenta.

Uso:
    python outreach.py                       # elegís la planilla de la carpeta LEADS de Drive
    python outreach.py --rubro kinesiologo   # la que tenga ese rubro en el nombre
    python outreach.py output/leads_x.xlsx   # una planilla local (sin Drive)
"""
from __future__ import annotations

import argparse
import io
import logging
import platform
import re
import subprocess
import sys
import webbrowser
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook

DEFAULT_TEMPLATES = Path(__file__).resolve().parent / "mensajes.txt"
DEFAULT_DAILY_CAP = 15
REQUIRED = ("Nombre", "Dirección", "Instagram", "Contactado", "Fecha contacto", "Notas")

log = logging.getLogger("outreach")


# ---------- plantillas ----------

def load_templates(path: Path) -> list[str]:
    """Lee las variantes del archivo de plantillas (separadas por '---')."""
    lines = [ln for ln in path.read_text(encoding="utf-8").splitlines() if not ln.lstrip().startswith("#")]
    variants = [v.strip() for v in "\n".join(lines).split("\n---\n")]
    variants = [v for v in variants if v]
    if not variants:
        raise ValueError(f"No hay plantillas en {path}")
    return variants


def render_message(template: str, negocio: str, detalle: str) -> str:
    """Reemplaza las variables y borra los bloques [[ ]] si no hay detalle."""
    detalle = detalle.strip()
    if detalle:
        text = re.sub(r"\[\[(.*?)\]\]", r"\1", template, flags=re.DOTALL)
    else:
        text = re.sub(r"\[\[.*?\]\]", "", template, flags=re.DOTALL)
    return text.replace("{negocio}", negocio).replace("{detalle}", detalle).strip()


def copy_to_clipboard(text: str) -> bool:
    system = platform.system()
    try:
        if system == "Windows":
            subprocess.run(["clip"], input=text.encode("utf-16le"), check=True, timeout=10)
        elif system == "Darwin":
            subprocess.run(["pbcopy"], input=text.encode("utf-8"), check=True, timeout=10)
        else:
            subprocess.run(["xclip", "-selection", "clipboard"], input=text.encode("utf-8"), check=True, timeout=10)
    except (OSError, subprocess.SubprocessError):
        return False
    return True


# ---------- lectura de la planilla ----------

def header_columns(ws) -> dict[str, int]:
    header = [c.value for c in ws[1]]
    missing = [c for c in REQUIRED if c not in header]
    if missing:
        raise ValueError(f"A la planilla le faltan columnas: {', '.join(missing)}. ¿La generó este programa?")
    return {name: header.index(name) + 1 for name in REQUIRED}


def lead_key(ws, row: int, cols: dict[str, int]) -> tuple[str, str]:
    """Identifica un lead por nombre + dirección (no depende del número de fila)."""
    get = lambda name: str(ws.cell(row=row, column=cols[name]).value or "").strip().lower()
    return get("Nombre"), get("Dirección")


def find_row(ws, cols: dict[str, int], key: tuple[str, str]) -> int | None:
    for r in range(2, ws.max_row + 1):
        if lead_key(ws, r, cols) == key:
            return r
    return None


def profile_url(cell) -> str:
    """URL del perfil: el hipervínculo si existe, si no el texto de la celda."""
    if cell.hyperlink and cell.hyperlink.target:
        return cell.hyperlink.target
    value = str(cell.value or "").strip()
    if value.startswith("@"):
        return f"https://www.instagram.com/{value[1:]}/"
    return value


def pending_rows(ws, cols: dict[str, int]) -> list[int]:
    """Filas con Instagram y sin contactar."""
    return [
        r
        for r in range(2, ws.max_row + 1)
        if ws.cell(row=r, column=cols["Instagram"]).value and not str(ws.cell(row=r, column=cols["Contactado"]).value or "").strip()
    ]


def count_contacted_today(ws, cols: dict[str, int]) -> int:
    today = date.today().isoformat()
    return sum(1 for r in range(2, ws.max_row + 1) if str(ws.cell(row=r, column=cols["Fecha contacto"]).value or "") == today)


def apply_mark(ws, cols: dict[str, int], row: int, variant_label: str) -> None:
    ws.cell(row=row, column=cols["Contactado"]).value = "Sí"
    ws.cell(row=row, column=cols["Fecha contacto"]).value = date.today().isoformat()
    notes = ws.cell(row=row, column=cols["Notas"])
    notes.value = f"{notes.value} | {variant_label}" if notes.value else variant_label


def _to_bytes(wb: Workbook) -> bytes:
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# ---------- dónde viven las planillas ----------

class LocalStore:
    """Una planilla .xlsx en esta PC."""

    def __init__(self, path: Path):
        self.path = path
        self.label = path.name

    def load(self) -> Workbook:
        return load_workbook(self.path)

    def contacted_today(self) -> int:
        ws = self.load().active
        return count_contacted_today(ws, header_columns(ws))

    def mark(self, key: tuple[str, str], variant_label: str) -> bool:
        wb = self.load()
        ws = wb.active
        cols = header_columns(ws)
        row = find_row(ws, cols, key)
        if row is None:
            log.error("No encontré ese lead en la planilla.")
            return False
        apply_mark(ws, cols, row, variant_label)
        try:
            wb.save(self.path)
        except PermissionError:
            log.error("No pude guardar %s. Cerrá el archivo en Excel y probá de nuevo.", self.path)
            return False
        return True


class DriveStore:
    """Una planilla de la carpeta LEADS de Drive. Cada marca baja la última versión,
    modifica solo la fila de ese lead y la vuelve a subir (no pisa cambios hechos en otra PC)."""

    def __init__(self, session, file: dict, all_files: list[dict]):
        self.session = session
        self.file_id = file["id"]
        self.label = file["name"]
        self.all_files = all_files

    def _download(self) -> Workbook:
        import drive_upload

        return load_workbook(io.BytesIO(drive_upload.download_file(self.session, self.file_id)))

    def load(self) -> Workbook:
        return self._download()

    def contacted_today(self) -> int:
        """Cuenta los contactos de hoy en TODAS las planillas de LEADS: el tope es por cuenta, no por rubro."""
        import drive_upload

        total = 0
        for f in self.all_files:
            try:
                ws = load_workbook(io.BytesIO(drive_upload.download_file(self.session, f["id"]))).active
                total += count_contacted_today(ws, header_columns(ws))
            except (ValueError, drive_upload.DriveError) as exc:
                log.warning("No pude leer %s para contar el tope diario: %s", f["name"], exc)
        return total

    def mark(self, key: tuple[str, str], variant_label: str) -> bool:
        import drive_upload

        try:
            wb = self._download()
            ws = wb.active
            cols = header_columns(ws)
            row = find_row(ws, cols, key)
            if row is None:
                log.error("No encontré ese lead en la planilla de Drive (¿lo borraste o le cambiaste el nombre?).")
                return False
            apply_mark(ws, cols, row, variant_label)
            drive_upload.update_file(self.session, self.file_id, _to_bytes(wb))
        except drive_upload.DriveError as exc:
            log.error("No pude actualizar la planilla en Drive: %s", exc)
            return False
        return True


def choose_drive_sheet(session, rubro: str | None) -> DriveStore | None:
    import drive_upload

    folder_id = drive_upload.find_or_create_folder(session, drive_upload.DEFAULT_FOLDER)
    files = drive_upload.list_sheets(session, folder_id)
    if not files:
        log.error("No hay planillas en la carpeta LEADS de Drive. Generá una con gmaps_leads.py.")
        return None
    candidates = [f for f in files if rubro.lower() in f["name"].lower()] if rubro else files
    if not candidates:
        log.error("Ninguna planilla de LEADS tiene '%s' en el nombre.", rubro)
        return None
    if len(candidates) == 1:
        chosen = candidates[0]
    else:
        print("Planillas en LEADS:")
        for i, f in enumerate(candidates, start=1):
            print(f"  {i}. {f['name']}  (modificada {f.get('modifiedTime', '')[:10]})")
        raw = input("Elegí el número: ").strip()
        if not raw.isdigit() or not 1 <= int(raw) <= len(candidates):
            log.error("Opción inválida.")
            return None
        chosen = candidates[int(raw) - 1]
    return DriveStore(session, chosen, files)


# ---------- el flujo principal ----------

def run(store, templates: list[str], daily_cap: int) -> int:
    wb = store.load()
    ws = wb.active
    try:
        cols = header_columns(ws)
    except ValueError as exc:
        log.error("%s", exc)
        return 1

    todo = pending_rows(ws, cols)
    if not todo:
        print(f"No hay leads con Instagram pendientes de contactar en {store.label}.")
        return 0
    sent_today = store.contacted_today()
    print(f"\n{store.label}: {len(todo)} leads pendientes. Hoy ya contactaste {sent_today} (tope diario sugerido: {daily_cap}).")

    for row in todo:
        if sent_today >= daily_cap:
            print(f"\nLlegaste al tope de {daily_cap} por hoy. Seguí mañana: escribir de a pocos protege tu cuenta.")
            break
        negocio = str(ws.cell(row=row, column=cols["Nombre"]).value or "")
        key = lead_key(ws, row, cols)
        url = profile_url(ws.cell(row=row, column=cols["Instagram"]))
        variant_no = sent_today % len(templates)
        print("\n" + "=" * 60)
        print(f"{negocio}\n{url}")
        detalle = input("Detalle personal que viste en su perfil (Enter para omitir, 's' saltar, 'q' salir): ")
        if detalle.strip().lower() == "q":
            break
        if detalle.strip().lower() == "s":
            continue
        message = render_message(templates[variant_no], negocio, detalle)
        print(f"\n--- Mensaje (variante {variant_no + 1}) ---\n{message}\n")
        input("Enter para abrir el perfil y copiar el mensaje... ")
        webbrowser.open(url)
        if copy_to_clipboard(message):
            print("Mensaje copiado. Pegalo en el DM (Ctrl+V), revisalo y enviá vos.")
        else:
            print("No pude usar el portapapeles: copiá el mensaje de arriba a mano.")
        answer = input("¿Lo enviaste? [s] sí  [n] no  [q] salir: ").strip().lower()
        if answer == "q":
            break
        if answer != "s":
            continue
        if not store.mark(key, f"Variante {variant_no + 1}"):
            print("No se pudo anotar en la planilla. Cortamos acá para no perder el control de a quién escribiste.")
            return 1
        sent_today += 1
        print(f"Anotado en {store.label}. Van {sent_today} hoy.")
    return 0


def build_store(args) -> LocalStore | DriveStore | None:
    if args.sheet:
        if not args.sheet.is_file():
            log.error("No existe el archivo %s", args.sheet)
            return None
        return LocalStore(args.sheet)
    try:
        import drive_upload
    except ImportError:
        log.error("Faltan librerías de Google (pip install -r requirements.txt).")
        return None
    if not (drive_upload.find_credentials() or drive_upload.TOKEN_PATH.is_file()):
        log.error("Drive no está configurado (ver README). Mientras tanto podés pasar una planilla local: python outreach.py output/leads_x.xlsx")
        return None
    try:
        return choose_drive_sheet(drive_upload.get_session(), args.rubro)
    except drive_upload.DriveError as exc:
        log.error("%s", exc)
        return None


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    parser = argparse.ArgumentParser(description="Contactar leads por Instagram en modo asistido (enviás vos).")
    parser.add_argument("sheet", nargs="?", type=Path, help="Planilla local .xlsx (si no la pasás, usa Drive)")
    parser.add_argument("--rubro", help="Filtra las planillas de Drive por texto en el nombre, por ejemplo: kinesiologo")
    parser.add_argument("--templates", type=Path, default=DEFAULT_TEMPLATES, help="Archivo de plantillas")
    parser.add_argument("--daily-cap", type=int, default=DEFAULT_DAILY_CAP, help="Tope sugerido por día (default: 15)")
    args = parser.parse_args(argv)

    try:
        templates = load_templates(args.templates)
    except (OSError, ValueError) as exc:
        log.error("%s", exc)
        return 1
    store = build_store(args)
    if store is None:
        return 1
    return run(store, templates, args.daily_cap)


if __name__ == "__main__":
    sys.exit(main())
