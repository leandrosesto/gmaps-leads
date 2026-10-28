"""Busca negocios en Google Maps (Places API New) y exporta una planilla de leads.

Uso:
    python gmaps_leads.py "kinesiólogo" "clínica" --city "Mar del Plata"
"""
from __future__ import annotations

import argparse
import csv
import logging
import os
import sys
import time
from datetime import datetime
from pathlib import Path

import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

API_URL = "https://places.googleapis.com/v1/places:searchText"
FIELD_MASK = ",".join(
    [
        "places.id",
        "places.displayName",
        "places.formattedAddress",
        "places.nationalPhoneNumber",
        "places.internationalPhoneNumber",
        "places.websiteUri",
        "places.rating",
        "places.userRatingCount",
        "places.googleMapsUri",
        "places.businessStatus",
        "nextPageToken",
    ]
)
PAGE_SIZE = 20  # máximo permitido por la API
MAX_PAGES = 3  # Google devuelve hasta 60 resultados por búsqueda
MAX_RETRIES = 4
TIMEOUT = 30

COLUMNS = [
    "Nombre",
    "Rubro buscado",
    "Dirección",
    "Teléfono",
    "WhatsApp (verificar)",
    "Web",
    "Rating",
    "Reseñas",
    "Google Maps",
    "Contactado",
    "Fecha contacto",
    "Respuesta",
    "Notas",
]

log = logging.getLogger("gmaps_leads")


class ApiError(RuntimeError):
    """Error devuelto por la API de Google con un mensaje legible."""


def whatsapp_link(international_phone: str | None) -> str:
    """Arma un link wa.me. Solo sirve si el número es celular y tiene WhatsApp."""
    if not international_phone:
        return ""
    digits = "".join(ch for ch in international_phone if ch.isdigit())
    if digits.startswith("54") and not digits.startswith("549"):
        digits = "549" + digits[2:]  # los celulares argentinos llevan el 9
    return f"https://wa.me/{digits}" if digits else ""


def parse_place(place: dict, query: str) -> dict:
    return {
        "id": place.get("id", ""),
        "Nombre": (place.get("displayName") or {}).get("text", ""),
        "Rubro buscado": query,
        "Dirección": place.get("formattedAddress", ""),
        "Teléfono": place.get("nationalPhoneNumber", ""),
        "WhatsApp (verificar)": whatsapp_link(place.get("internationalPhoneNumber")),
        "Web": place.get("websiteUri", ""),
        "Rating": place.get("rating", ""),
        "Reseñas": place.get("userRatingCount", ""),
        "Google Maps": place.get("googleMapsUri", ""),
        "Contactado": "",
        "Fecha contacto": "",
        "Respuesta": "",
        "Notas": "",
        "_status": place.get("businessStatus", ""),
    }


def _post(session: requests.Session, api_key: str, body: dict) -> dict:
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": api_key,
        "X-Goog-FieldMask": FIELD_MASK,
    }
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.post(API_URL, json=body, headers=headers, timeout=TIMEOUT)
        except requests.RequestException as exc:
            if attempt == MAX_RETRIES:
                raise ApiError(f"No se pudo conectar con Google: {exc}") from exc
            time.sleep(2**attempt)
            continue
        if resp.status_code == 200:
            return resp.json()
        if resp.status_code in (429, 500, 502, 503, 504) and attempt < MAX_RETRIES:
            wait = 2**attempt
            log.warning("HTTP %s, reintento en %ss", resp.status_code, wait)
            time.sleep(wait)
            continue
        try:
            detail = resp.json().get("error", {}).get("message", resp.text)
        except ValueError:
            detail = resp.text
        raise ApiError(f"Google respondió {resp.status_code}: {detail}")
    raise ApiError("Se agotaron los reintentos")


def search(api_key: str, query: str, city: str, limit: int, session: requests.Session) -> list[dict]:
    """Devuelve hasta `limit` lugares (tope 60) para '<query> en <city>'."""
    text = f"{query} en {city}"
    results: list[dict] = []
    token: str | None = None
    for _ in range(MAX_PAGES):
        body = {
            "textQuery": text,
            "languageCode": "es",
            "regionCode": "AR",
            "pageSize": min(PAGE_SIZE, limit - len(results)),
        }
        if token:
            body["pageToken"] = token
        data = _post(session, api_key, body)
        results += [parse_place(p, query) for p in data.get("places", [])]
        token = data.get("nextPageToken")
        if not token or len(results) >= limit:
            break
        time.sleep(1)  # el token de página tarda un instante en activarse
    return results[:limit]


def collect(api_key: str, queries: list[str], city: str, limit: int, only_phone: bool) -> list[dict]:
    seen: set[str] = set()
    rows: list[dict] = []
    with requests.Session() as session:
        for query in queries:
            log.info("Buscando: %s en %s", query, city)
            for row in search(api_key, query, city, limit, session):
                if row["id"] in seen or row["_status"] == "CLOSED_PERMANENTLY":
                    continue
                if only_phone and not row["Teléfono"]:
                    continue
                seen.add(row["id"])
                rows.append(row)
            log.info("Acumulados: %d negocios únicos", len(rows))
    return rows


def write_xlsx(rows: list[dict], path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"
    ws.append(COLUMNS)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for row in rows:
        ws.append([row[c] for c in COLUMNS])
    widths = {"Nombre": 32, "Dirección": 40, "WhatsApp (verificar)": 34, "Web": 30, "Google Maps": 30, "Notas": 30}
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
    p.add_argument("--max", type=int, default=60, dest="limit", help="Máximo por rubro, tope 60 (default: 60)")
    p.add_argument("--out", type=Path, help="Archivo de salida (default: output/leads_<fecha>.xlsx)")
    p.add_argument("--csv", action="store_true", help="Exportar CSV en vez de Excel")
    p.add_argument("--only-phone", action="store_true", help="Descartar negocios sin teléfono")
    return p


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    args = build_parser().parse_args(argv)
    load_dotenv()
    api_key = os.environ.get("GOOGLE_MAPS_API_KEY", "").strip()
    if not api_key:
        log.error("Falta GOOGLE_MAPS_API_KEY. Copiá .env.example a .env y pegá tu clave (ver README).")
        return 1
    args.limit = max(1, min(args.limit, PAGE_SIZE * MAX_PAGES))

    try:
        rows = collect(api_key, args.queries, args.city, args.limit, args.only_phone)
    except ApiError as exc:
        log.error("%s", exc)
        return 1
    if not rows:
        log.warning("No se encontraron negocios.")
        return 0

    ext = ".csv" if args.csv else ".xlsx"
    out = args.out or Path("output") / f"leads_{datetime.now():%Y%m%d_%H%M}{ext}"
    out.parent.mkdir(parents=True, exist_ok=True)
    (write_csv if out.suffix == ".csv" else write_xlsx)(rows, out)
    log.info("Listo: %d negocios guardados en %s", len(rows), out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
