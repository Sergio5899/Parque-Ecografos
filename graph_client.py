# -*- coding: utf-8 -*-
"""Cliente mínimo de Microsoft Graph para leer/escribir archivos en OneDrive/SharePoint."""
import os
import requests

import graph_auth

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"

# El Excel vive en una biblioteca de SharePoint (no en el OneDrive personal),
# montada localmente como "CCM Servicio - Documentos". Se puede sobrescribir
# por variables de entorno si el sitio o la ruta cambian.
SITE_HOSTNAME = os.environ.get("MS_SITE_HOSTNAME", "comercialcentromedico.sharepoint.com")
SITE_PATH = os.environ.get("MS_SITE_PATH", "/sites/CCMServicio")
FILE_RELATIVE_PATH = os.environ.get(
    "MS_FILE_RELATIVE_PATH", "Manuales Técnicos/US/PARQUE ECOS/ECOS DATOS.xlsx"
)

# Copia que vive en el OneDrive personal del usuario (fallback sin permisos de Sites.*).
MY_DRIVE_FILE_PATH = os.environ.get("MS_MY_DRIVE_FILE_PATH", "PARQUE ECOS/ECOS DATOS.xlsx")


def _headers():
    return {"Authorization": "Bearer " + graph_auth.get_access_token()}


def _req(method, url, headers=None, **kwargs):
    all_headers = _headers()
    if headers:
        all_headers.update(headers)
    res = requests.request(method, url, headers=all_headers, timeout=30, **kwargs)
    if res.status_code >= 400:
        raise RuntimeError(f"Graph {method} {url} -> {res.status_code}: {res.text[:500]}")
    return res


def _find_via_site(filename):
    """Resuelve el archivo dentro del sitio de SharePoint conocido (requiere Sites.ReadWrite.All)."""
    site_url = f"{GRAPH_ROOT}/sites/{SITE_HOSTNAME}:{SITE_PATH}"
    site = _req("GET", site_url).json()
    site_id = site["id"]

    drive = _req("GET", f"{GRAPH_ROOT}/sites/{site_id}/drive").json()
    drive_id = drive["id"]

    from urllib.parse import quote
    item_url = f"{GRAPH_ROOT}/drives/{drive_id}/root:/{quote(FILE_RELATIVE_PATH)}"
    item = _req("GET", item_url).json()
    return drive_id, item["id"], item.get("parentReference", {}).get("path", "")


def _find_via_my_drive(filename):
    """Resuelve el archivo por ruta directa en el OneDrive personal del usuario.
    Se usa ruta en vez de búsqueda porque el índice de búsqueda de Graph tarda
    en ponerse al día tras subir/mover un archivo, mientras que la ruta directa
    refleja el estado real al instante."""
    from urllib.parse import quote
    url = f"{GRAPH_ROOT}/me/drive/root:/{quote(MY_DRIVE_FILE_PATH)}"
    res = requests.get(url, headers=_headers(), timeout=30)
    if res.status_code == 200:
        item = res.json()
        drive_id = item["parentReference"]["driveId"]
        return drive_id, item["id"], item["parentReference"].get("path", "")

    # Fallback: búsqueda por nombre (más lenta de indexar, pero más flexible).
    url = f"{GRAPH_ROOT}/me/drive/root/search(q='{filename}')"
    res = _req("GET", url)
    items = res.json().get("value", [])
    for it in items:
        if it.get("name", "").lower() == filename.lower() and "file" in it:
            drive_id = it["parentReference"]["driveId"]
            return drive_id, it["id"], it["parentReference"].get("path", "")
    return None


def find_item_by_name(filename):
    """Devuelve (drive_id, item_id, path) o None si no se encuentra.
    Primero busca en el OneDrive personal del usuario (donde vive ahora el
    archivo real); si no aparece, prueba el sitio de SharePoint (requiere
    Sites.ReadWrite.All, no siempre disponible)."""
    found = _find_via_my_drive(filename)
    if found:
        return found
    try:
        return _find_via_site(filename)
    except Exception as e:
        print(f"WARN: no se pudo resolver via sitio de SharePoint ({e}).")
        return None


def download_bytes(drive_id, item_id):
    url = f"{GRAPH_ROOT}/drives/{drive_id}/items/{item_id}/content"
    res = _req("GET", url)
    return res.content


def upload_bytes(drive_id, item_id, data):
    url = f"{GRAPH_ROOT}/drives/{drive_id}/items/{item_id}/content"
    _req("PUT", url, data=data, headers={"Content-Type": "application/octet-stream"})


def download_bytes_by_path(path):
    """path relativo a la raíz de /me/drive, ej: '_app_data/refresh_token.txt'. None si no existe."""
    url = f"{GRAPH_ROOT}/me/drive/root:/{path}:/content"
    res = requests.get(url, headers=_headers(), timeout=30)
    if res.status_code == 404:
        return None
    if res.status_code >= 400:
        raise RuntimeError(f"Graph GET {url} -> {res.status_code}: {res.text[:300]}")
    return res.content


def upload_bytes_by_path(path, data):
    url = f"{GRAPH_ROOT}/me/drive/root:/{path}:/content"
    requests.put(url, data=data, headers={**_headers(), "Content-Type": "application/octet-stream"}, timeout=30)
