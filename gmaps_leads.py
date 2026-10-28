"""Busca negocios en Google Maps con un navegador automatizado (sin API ni tarjeta).

Uso:
    python gmaps_leads.py "kinesiólogo" "clínica" --city "Mar del Plata"
"""
from __future__ import annotations

import argparse
import csv
import logging
import random
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.sync_api import Error as PWError, Page, TimeoutError as PWTimeout, sync_playwright

MAX_RESULTS = 120
COLUMNS = [
    "Nombre",
    "Rubro buscado",
    "Dirección",
    "Teléfono",
    "WhatsApp (verificar)",
    "Instagram",
    "Web",
    "Rating",
    "Google Maps",
    "Contactado",
    "Fecha contacto",
    "Respuesta",
    "Notas",
]

log = logging.getLogger("gmaps_leads")


def whatsapp_link(phone: str | None) -> str:
    """Arma un link wa.me desde un teléfono argentino. Solo sirve si es celular con WhatsApp."""
    if not phone:
        return ""
    text = phone.strip()
    if text.startswith("+54"):
        digits = re.sub(r"\D", "", text)
        if not digits.startswith("549"):
            digits = "549" + digits[2:]
        return f"https://wa.me/{digits}"
    # Formato local: 0223 15-456-7890 / 223 456-7890 -> quitar 0 inicial y el 15
    m = re.match(r"^0?(\d{2,4})[\s-]+15[\s-]*([\d\s-]{6,10})$", text)
    digits = (m.group(1) + re.sub(r"\D", "", m.group(2))) if m else re.sub(r"\D", "", text).lstrip("0")
    if len(digits) != 10:
        return ""
    return f"https://wa.me/549{digits}"


def instagram_handle(url: str) -> str:
    m = re.search(r"instagram\.com/([^/?#\s]+)", url, re.IGNORECASE)
    return f"@{m.group(1)}" if m else url


def domain_label(url: str) -> str:
    m = re.match(r"^(?:https?://)?(?:www\.)?([^/?#\s]+)", url, re.IGNORECASE)
    return m.group(1) if m else url


# Columnas que en el Excel se muestran como texto corto clickeable
LINK_LABELS = {
    "WhatsApp (verificar)": lambda url: "Abrir WhatsApp",
    "Instagram": instagram_handle,
    "Web": domain_label,
    "Google Maps": lambda url: "Ver en Maps",
}


_IG_RESERVED = {"p", "reel", "reels", "explore", "accounts", "sharer", "share", "stories", "tv", "direct", "embed", "web", "about", "legal"}


def extract_instagram(hrefs: list[str]) -> str:
    """Devuelve el perfil de Instagram más repetido entre los links de una página."""
    counts: dict[str, int] = {}
    for href in hrefs:
        m = re.search(r"instagram\.com/([A-Za-z0-9._]+)", href, re.IGNORECASE)
        if m and m.group(1).lower() not in _IG_RESERVED:
            handle = m.group(1)
            counts[handle] = counts.get(handle, 0) + 1
    if not counts:
        return ""
    return f"https://www.instagram.com/{max(counts, key=counts.get)}/"


def find_instagram_on_site(page: Page, url: str) -> str:
    """Abre el sitio del negocio y busca su link de Instagram (sin loguearse en nada)."""
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=15000)
        page.wait_for_timeout(1200)
        hrefs = page.eval_on_selector_all('a[href*="instagram.com"]', "els => els.map(e => e.href)")
    except (PWTimeout, PWError):
        return ""
    return extract_instagram(hrefs)


def short_maps_url(url: str) -> str:
    """Convierte el link largo de la ficha en uno corto basado en el place ID."""
    m = re.search(r"!19s(ChIJ[\w-]+)", url)
    if m:
        return f"https://www.google.com/maps/place/?q=place_id:{m.group(1)}"
    return url.split("?")[0]


def _text(page: Page, selector: str) -> str:
    loc = page.locator(selector).first
    try:
        return (loc.inner_text(timeout=1500) or "").strip()
    except PWTimeout:
        return ""


def _clean_button_text(raw: str) -> str:
    """Los botones de la ficha traen un ícono privado de Unicode al principio."""
    return re.sub(r"^[-\s]+", "", raw).strip()


def accept_or_reject_consent(page: Page) -> None:
    """Si Google muestra el cartel de cookies, lo rechaza."""
    for label in ("Rechazar todo", "Reject all"):
        btn = page.get_by_role("button", name=label)
        if btn.count():
            btn.first.click()
            page.wait_for_load_state("domcontentloaded")
            return


def collect_place_links(page: Page, query: str, city: str, limit: int) -> list[str]:
    url = f"https://www.google.com/maps/search/{quote(f'{query} en {city}')}?hl=es&gl=ar"
    page.goto(url, wait_until="domcontentloaded")
    accept_or_reject_consent(page)
    try:
        page.wait_for_selector('div[role="feed"], h1', timeout=20000)
    except PWTimeout:
        log.warning("No cargó la lista de resultados para '%s'", query)
        return []

    feed = page.locator('div[role="feed"]')
    if not feed.count():
        # Google fue directo a una única ficha
        return [page.url]

    links: list[str] = []
    stale_rounds = 0
    while len(links) < limit and stale_rounds < 4:
        hrefs = page.eval_on_selector_all(
            'div[role="feed"] a[href*="/maps/place/"]', "els => els.map(e => e.href)"
        )
        new = [h for h in dict.fromkeys(hrefs) if h not in links]
        links += new
        stale_rounds = 0 if new else stale_rounds + 1
        if page.get_by_text("Llegaste al final de la lista").count():
            break
        feed.first.evaluate("el => el.scrollBy(0, el.scrollHeight)")
        time.sleep(random.uniform(1.2, 2.0))
    return links[:limit]


def scrape_place(page: Page, url: str, query: str) -> dict | None:
    page.goto(url, wait_until="domcontentloaded")
    try:
        page.wait_for_selector("h1", timeout=15000)
    except PWTimeout:
        return None
    name = _text(page, "h1")
    if not name:
        return None
    address = _clean_button_text(_text(page, 'button[data-item-id="address"]'))
    phone = _clean_button_text(_text(page, 'button[data-item-id^="phone"]'))
    website = ""
    site = page.locator('a[data-item-id="authority"]').first
    if site.count():
        website = site.get_attribute("href") or ""
    rating = ""
    stars = page.locator('span[role="img"][aria-label*="estrella"]').first
    if stars.count():
        m = re.search(r"([\d,\.]+)\s*estrella", stars.get_attribute("aria-label") or "")
        rating = m.group(1).replace(",", ".") if m else ""
    closed = bool(page.get_by_text("Cerrado permanentemente").count())
    return {
        "Nombre": name,
        "Rubro buscado": query,
        "Dirección": address,
        "Teléfono": phone,
        "WhatsApp (verificar)": whatsapp_link(phone),
        "Instagram": website if "instagram.com" in website.lower() else "",
        "Web": "" if "instagram.com" in website.lower() else website,
        "Rating": rating,
        "Google Maps": short_maps_url(page.url),
        "Contactado": "",
        "Fecha contacto": "",
        "Respuesta": "",
        "Notas": "",
        "_closed": closed,
    }


def collect(
    queries: list[str], city: str, limit: int, only_phone: bool, headless: bool, find_ig: bool = False
) -> list[dict]:
    rows: list[dict] = []
    seen: set[tuple[str, str]] = set()
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=headless)
        context = browser.new_context(locale="es-AR", timezone_id="America/Argentina/Buenos_Aires")
        page = context.new_page()
        site_page = context.new_page()
        try:
            for query in queries:
                log.info("Buscando: %s en %s", query, city)
                links = collect_place_links(page, query, city, limit)
                log.info("%d fichas encontradas, leyendo datos...", len(links))
                for idx, link in enumerate(links, start=1):
                    try:
                        row = scrape_place(page, link, query)
                    except PWTimeout:
                        log.warning("Timeout en la ficha %d, se omite", idx)
                        continue
                    time.sleep(random.uniform(0.8, 1.8))
                    if not row or row["_closed"]:
                        continue
                    key = (row["Nombre"].lower(), row["Dirección"].lower())
                    if key in seen or (only_phone and not row["Teléfono"]):
                        continue
                    seen.add(key)
                    if find_ig and row["Web"] and not row["Instagram"]:
                        row["Instagram"] = find_instagram_on_site(site_page, row["Web"])
                    rows.append(row)
                    if idx % 10 == 0:
                        log.info("  %d/%d leídas, %d válidas", idx, len(links), len(rows))
        finally:
            browser.close()
    return rows


def write_xlsx(rows: list[dict], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"
    ws.append(COLUMNS)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    link_cols = {COLUMNS.index(name) + 1: label for name, label in LINK_LABELS.items()}
    for row in rows:
        ws.append([row[c] for c in COLUMNS])
        for col, label in link_cols.items():
            cell = ws.cell(row=ws.max_row, column=col)
            if cell.value:
                cell.hyperlink = cell.value
                cell.value = label(cell.value)
                cell.font = Font(color="0563C1", underline="single")
    widths = {"Nombre": 32, "Dirección": 40, "Web": 24, "Instagram": 22, "Notas": 30}
    for idx, name in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = widths.get(name, 16)
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)


def write_csv(rows: list[dict], path: Path) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=COLUMNS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Genera una lista de leads desde Google Maps.")
    p.add_argument("queries", nargs="+", help='Rubros a buscar, por ejemplo: "kinesiólogo" "clínica"')
    p.add_argument("--city", default="Mar del Plata", help="Ciudad (default: Mar del Plata)")
    p.add_argument("--max", type=int, default=40, dest="limit", help="Máximo por rubro (default: 40)")
    p.add_argument("--out", type=Path, help="Archivo de salida (default: output/leads_<fecha>.xlsx)")
    p.add_argument("--csv", action="store_true", help="Exportar CSV en vez de Excel")
    p.add_argument("--only-phone", action="store_true", help="Descartar negocios sin teléfono")
    p.add_argument("--show", action="store_true", help="Mostrar el navegador mientras trabaja")
    p.add_argument(
        "--instagram",
        action="store_true",
        help="Buscar el Instagram en el sitio web de cada negocio (más lento, sin loguearse en nada)",
    )
    p.add_argument("--no-drive", action="store_true", help="No subir la planilla a Google Drive")
    p.add_argument("--drive-folder", default="LEADS", help="Carpeta de Drive donde se sube (default: LEADS)")
    return p


def slugify(text: str) -> str:
    """'Kinesiólogo' -> 'kinesiologo' para usar en nombres de archivo."""
    import unicodedata

    ascii_text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "-", ascii_text.lower()).strip("-") or "leads"


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    args = build_parser().parse_args(argv)
    args.limit = max(1, min(args.limit, MAX_RESULTS))

    rows = collect(args.queries, args.city, args.limit, args.only_phone, headless=not args.show, find_ig=args.instagram)
    if not rows:
        log.warning("No se encontraron negocios. Probá con --show para ver qué pasa.")
        return 0

    ext = ".csv" if args.csv else ".xlsx"
    out = args.out or Path("output") / f"leads_{slugify(args.queries[0])}_{datetime.now():%Y%m%d_%H%M}{ext}"
    out.parent.mkdir(parents=True, exist_ok=True)
    (write_csv if out.suffix == ".csv" else write_xlsx)(rows, out)
    log.info("Listo: %d negocios guardados en %s", len(rows), out)

    if not args.no_drive:
        upload_if_configured(out, args.drive_folder)
    return 0


def upload_if_configured(path: Path, folder: str) -> None:
    """Sube el archivo a Drive si hay credenciales configuradas. Un fallo no pierde el archivo local."""
    try:
        import drive_upload
    except ImportError:
        log.warning("Faltan librerías de Google (pip install -r requirements.txt). Se omite la subida a Drive.")
        return
    if not (drive_upload.find_credentials() or drive_upload.TOKEN_PATH.is_file()):
        log.info("Drive no configurado: la planilla quedó solo en tu PC (ver README, sección Google Drive).")
        return
    try:
        link = drive_upload.upload_to_drive(path, folder)
    except Exception as exc:  # cualquier fallo de red/auth no debe tirar abajo el resultado ya guardado
        log.error("No se pudo subir a Drive: %s. El archivo local está en %s", exc, path)
        return
    log.info("Subido a Drive (carpeta %s): %s", folder, link)


if __name__ == "__main__":
    sys.exit(main())
