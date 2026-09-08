# -*- coding: utf-8 -*-
"""Cliente mínimo de Microsoft Graph para leer/escribir archivos en OneDrive."""
import io
import requests

import graph_auth

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"


def _headers():
    return {"Authorization": "Bearer " + graph_auth.get_access_token()}


def _req(method, url, **kwargs):
    res = requests.request(method, url, headers=_headers(), timeout=30, **kwargs)
    if res.status_code >= 400:
        raise RuntimeError(f"Graph {method} {url} -> {res.status_code}: {res.text[:500]}")
    return res


def find_item_by_name(filename):
    """Busca un archivo por nombre en todo el OneDrive del usuario.
    Devuelve (drive_id, item_id, path) o None si no se encuentra."""
    url = f"{GRAPH_ROOT}/me/drive/root/search(q='{filename}')"
    res = _req("GET", url)
    items = res.json().get("value", [])
    for it in items:
        if it.get("name", "").lower() == filename.lower() and "file" in it:
            drive_id = it["parentReference"]["driveId"]
            return drive_id, it["id"], it["parentReference"].get("path", "")
    return None


def download_bytes(drive_id, item_id):
    url = f"{GRAPH_ROOT}/drives/{drive_id}/items/{item_id}/content"
    res = _req("GET", url)
    return res.content


def upload_bytes(drive_id, item_id, data):
    url = f"{GRAPH_ROOT}/drives/{drive_id}/items/{item_id}/content"
    _req("PUT", url, data=data, headers={**_headers(), "Content-Type": "application/octet-stream"})


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
