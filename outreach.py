"""Modo asistido para contactar leads por Instagram. VOS enviás cada mensaje.

Por cada lead sin contactar de la planilla, el programa:
  1. te muestra el negocio y el mensaje ya armado,
  2. abre su perfil de Instagram en tu navegador y copia el mensaje al portapapeles,
  3. espera a que vos lo pegues y lo envíes, y te pregunta si lo enviaste,
  4. marca el lead como contactado (con fecha y variante) en la planilla.

No inicia sesión en Instagram ni envía nada por su cuenta.

Uso:
    python outreach.py                      # usa la planilla más reciente de output/
    python outreach.py output/leads_x.xlsx  # una planilla específica
"""
from __future__ import annotations

import argparse
import logging
import platform
import re
import subprocess
import sys
import webbrowser
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

DEFAULT_TEMPLATES = Path(__file__).resolve().parent / "mensajes.txt"
DEFAULT_DAILY_CAP = 15

log = logging.getLogger("outreach")


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


def newest_sheet() -> Path | None:
    files = sorted(Path("output").glob("leads_*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None


def profile_url(cell) -> str:
    """URL del perfil: el hipervínculo si existe, si no el texto de la celda."""
    if cell.hyperlink and cell.hyperlink.target:
        return cell.hyperlink.target
    value = str(cell.value or "").strip()
    if value.startswith("@"):
        return f"https://www.instagram.com/{value[1:]}/"
    return value


def pending_leads(ws, cols: dict[str, int]) -> list[int]:
    """Filas con Instagram y sin contactar."""
    rows = []
    for r in range(2, ws.max_row + 1):
        has_ig = ws.cell(row=r, column=cols["Instagram"]).value
        done = str(ws.cell(row=r, column=cols["Contactado"]).value or "").strip()
        if has_ig and not done:
            rows.append(r)
    return rows


def contacted_today(ws, cols: dict[str, int]) -> int:
    today = date.today().isoformat()
    return sum(
        1 for r in range(2, ws.max_row + 1) if str(ws.cell(row=r, column=cols["Fecha contacto"]).value or "") == today
    )


def save_workbook(wb, path: Path) -> bool:
    try:
        wb.save(path)
    except PermissionError:
        log.error("No pude guardar %s. Cerrá el archivo en Excel y probá de nuevo.", path)
        return False
    return True


def run(path: Path, templates: list[str], daily_cap: int) -> int:
    wb = load_workbook(path)
    ws = wb.active
    header = [c.value for c in ws[1]]
    required = ("Nombre", "Instagram", "Contactado", "Fecha contacto", "Notas")
    missing = [c for c in required if c not in header]
    if missing:
        log.error("A la planilla le faltan columnas: %s. ¿La generó este programa?", ", ".join(missing))
        return 1
    cols = {name: header.index(name) + 1 for name in required}

    todo = pending_leads(ws, cols)
    if not todo:
        print("No hay leads con Instagram pendientes de contactar en esta planilla.")
        return 0
    sent_today = contacted_today(ws, cols)
    print(f"{len(todo)} leads pendientes. Hoy ya contactaste {sent_today} (tope diario sugerido: {daily_cap}).")

    for row in todo:
        if sent_today >= daily_cap:
            print(f"\nLlegaste al tope de {daily_cap} por hoy. Seguí mañana: escribir de a pocos protege tu cuenta.")
            break
        negocio = str(ws.cell(row=row, column=cols["Nombre"]).value or "")
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
        ws.cell(row=row, column=cols["Contactado"]).value = "Sí"
        ws.cell(row=row, column=cols["Fecha contacto"]).value = date.today().isoformat()
        ws.cell(row=row, column=cols["Notas"]).value = f"Variante {variant_no + 1}"
        if not save_workbook(wb, path):
            return 1
        sent_today += 1
        print(f"Anotado. Van {sent_today} hoy.")
    return 0


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    parser = argparse.ArgumentParser(description="Contactar leads por Instagram en modo asistido (enviás vos).")
    parser.add_argument("sheet", nargs="?", type=Path, help="Planilla .xlsx (default: la más reciente de output/)")
    parser.add_argument("--templates", type=Path, default=DEFAULT_TEMPLATES, help="Archivo de plantillas")
    parser.add_argument("--daily-cap", type=int, default=DEFAULT_DAILY_CAP, help="Tope sugerido por día (default: 15)")
    args = parser.parse_args(argv)

    sheet = args.sheet or newest_sheet()
    if not sheet or not sheet.is_file():
        log.error("No encontré ninguna planilla. Generá una con gmaps_leads.py o pasá la ruta.")
        return 1
    try:
        templates = load_templates(args.templates)
    except (OSError, ValueError) as exc:
        log.error("%s", exc)
        return 1
    return run(sheet, templates, args.daily_cap)


if __name__ == "__main__":
    sys.exit(main())
