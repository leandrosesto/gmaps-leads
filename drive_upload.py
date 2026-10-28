"""Sube archivos a una carpeta de Google Drive (por defecto "LEADS").

Usa la API de Drive con OAuth (gratis, no requiere tarjeta) y el permiso mínimo
`drive.file`: el programa solo ve los archivos y carpetas que él mismo creó.

Uso suelto, para probar la configuración:
    python drive_upload.py --login
"""
from __future__ import annotations

import argparse
import json
import logging
import mimetypes
import os
import sys
from pathlib import Path

from google.auth.transport.requests import AuthorizedSession, Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/drive.file"]
FOLDER_MIME = "application/vnd.google-apps.folder"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
FILES_URL = "https://www.googleapis.com/drive/v3/files"
UPLOAD_URL = "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink"
CONFIG_DIR = Path.home() / ".gmaps-leads"
TOKEN_PATH = CONFIG_DIR / "token.json"
DEFAULT_FOLDER = "LEADS"

log = logging.getLogger("drive_upload")


class DriveError(RuntimeError):
    """Error de configuración o de la API de Drive, con mensaje legible."""


def find_credentials() -> Path | None:
    """Busca credentials.json: variable de entorno, carpeta del proyecto o carpeta de config."""
    candidates = []
    if os.environ.get("GMAPS_LEADS_CREDENTIALS"):
        candidates.append(Path(os.environ["GMAPS_LEADS_CREDENTIALS"]))
    candidates += [Path(__file__).resolve().parent / "credentials.json", CONFIG_DIR / "credentials.json"]
    return next((p for p in candidates if p.is_file()), None)


def get_session(credentials_path: Path | None = None) -> AuthorizedSession:
    """Devuelve una sesión autenticada. La primera vez abre el navegador para que inicies sesión."""
    creds: Credentials | None = None
    if TOKEN_PATH.is_file():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception as exc:  # google-auth lanza varios tipos; si falla, se vuelve a loguear
            log.warning("No se pudo renovar el acceso (%s). Hay que iniciar sesión de nuevo.", exc)
            creds = None
    if not creds or not creds.valid:
        credentials_path = credentials_path or find_credentials()
        if not credentials_path:
            raise DriveError("No encontré credentials.json (ver README, sección Google Drive).")
        flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), SCOPES)
        creds = flow.run_local_server(port=0, prompt="consent")
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        TOKEN_PATH.write_text(creds.to_json(), encoding="utf-8")
    return AuthorizedSession(creds)


def _check(resp, action: str) -> dict:
    if resp.status_code >= 400:
        try:
            detail = resp.json().get("error", {}).get("message", resp.text)
        except ValueError:
            detail = resp.text
        raise DriveError(f"Drive no pudo {action} ({resp.status_code}): {detail}")
    return resp.json()


def find_or_create_folder(session: AuthorizedSession, name: str = DEFAULT_FOLDER) -> str:
    """Devuelve el ID de la carpeta `name`, creándola si no existe."""
    safe = name.replace("\\", "\\\\").replace("'", "\\'")
    query = f"name = '{safe}' and mimeType = '{FOLDER_MIME}' and trashed = false"
    data = _check(session.get(FILES_URL, params={"q": query, "fields": "files(id,name)", "pageSize": 1}), "buscar la carpeta")
    if data.get("files"):
        return data["files"][0]["id"]
    created = _check(session.post(FILES_URL, json={"name": name, "mimeType": FOLDER_MIME}, params={"fields": "id"}), "crear la carpeta")
    log.info("Carpeta '%s' creada en tu Drive", name)
    return created["id"]


def upload_file(session: AuthorizedSession, path: Path, folder_id: str) -> dict:
    """Sube `path` a la carpeta indicada y devuelve {id, name, webViewLink}."""
    mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
    metadata = json.dumps({"name": path.name, "parents": [folder_id]})
    boundary = "gmapsleadsboundary"
    body = (
        f"--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n{metadata}\r\n"
        f"--{boundary}\r\nContent-Type: {mime}\r\n\r\n"
    ).encode("utf-8") + path.read_bytes() + f"\r\n--{boundary}--".encode("utf-8")
    resp = session.post(UPLOAD_URL, data=body, headers={"Content-Type": f"multipart/related; boundary={boundary}"})
    return _check(resp, "subir el archivo")


def list_sheets(session: AuthorizedSession, folder_id: str) -> list[dict]:
    """Planillas .xlsx de la carpeta, ordenadas por nombre: [{id, name, modifiedTime}]."""
    query = f"'{folder_id}' in parents and trashed = false and mimeType = '{XLSX_MIME}'"
    params = {"q": query, "fields": "files(id,name,modifiedTime)", "orderBy": "name", "pageSize": 100}
    return _check(session.get(FILES_URL, params=params), "listar las planillas").get("files", [])


def download_file(session: AuthorizedSession, file_id: str) -> bytes:
    resp = session.get(f"{FILES_URL}/{file_id}", params={"alt": "media"})
    if resp.status_code >= 400:
        _check(resp, "descargar la planilla")
    return resp.content


def update_file(session: AuthorizedSession, file_id: str, data: bytes, mime: str = XLSX_MIME) -> dict:
    """Reemplaza el contenido de un archivo existente (mismo ID, mismo link)."""
    resp = session.patch(
        f"https://www.googleapis.com/upload/drive/v3/files/{file_id}",
        params={"uploadType": "media", "fields": "id"},
        data=data,
        headers={"Content-Type": mime},
    )
    return _check(resp, "actualizar la planilla")


def upload_to_drive(path: Path, folder_name: str = DEFAULT_FOLDER) -> str:
    """Atajo: autentica, ubica la carpeta y sube el archivo. Devuelve el link."""
    session = get_session()
    folder_id = find_or_create_folder(session, folder_name)
    return upload_file(session, path, folder_id).get("webViewLink", "")


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    parser = argparse.ArgumentParser(description="Prueba la conexión con Google Drive.")
    parser.add_argument("--login", action="store_true", help="Inicia sesión y crea/verifica la carpeta")
    parser.add_argument("--folder", default=DEFAULT_FOLDER, help=f"Nombre de la carpeta (default: {DEFAULT_FOLDER})")
    args = parser.parse_args()
    if not args.login:
        parser.print_help()
        return 0
    try:
        session = get_session()
        folder_id = find_or_create_folder(session, args.folder)
    except DriveError as exc:
        log.error("%s", exc)
        return 1
    log.info("Todo listo. Carpeta '%s' en Drive: https://drive.google.com/drive/folders/%s", args.folder, folder_id)
    return 0


if __name__ == "__main__":
    sys.exit(main())
