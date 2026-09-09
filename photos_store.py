# -*- coding: utf-8 -*-
"""
Guarda las fotos de cada equipo. Si hay credenciales de Microsoft Graph
configuradas (mismo criterio que excel_store.USE_GRAPH), las fotos viven en
el OneDrive del usuario, junto al Excel, en una carpeta "fotos/<SN>/". Si no,
se guardan en una carpeta local "fotos/<SN>/" al lado del Excel local.
"""
import os
import re
import time
import mimetypes

import excel_store as store

USE_GRAPH = store.USE_GRAPH
GRAPH_BASE_PATH = "PARQUE ECOS/fotos"
LOCAL_BASE_DIR = os.path.join(store.BASE_DIR, "fotos")

ALLOWED_EXT = {"jpg", "jpeg", "png", "webp", "heic", "heif", "gif"}


class PhotoTooLargeError(Exception):
    pass


MAX_BYTES = 15 * 1024 * 1024  # 15 MB por foto


def _safe_ext(filename):
    ext = (os.path.splitext(filename or "")[1] or "").lower().lstrip(".")
    if ext not in ALLOWED_EXT:
        ext = "jpg"
    return ext


def _safe_name(filename):
    ext = _safe_ext(filename)
    stamp = time.strftime("%Y%m%d-%H%M%S")
    rand = format(int.from_bytes(os.urandom(3), "big"), "06x")
    return f"{stamp}-{rand}.{ext}"


def _content_type(filename):
    ctype = mimetypes.guess_type(filename)[0]
    return ctype or "application/octet-stream"


def _local_dir(sn):
    d = os.path.join(LOCAL_BASE_DIR, store.sanitize_sheet_name(sn))
    os.makedirs(d, exist_ok=True)
    return d


def list_photos(sn):
    sn = store.sanitize_sheet_name(sn)
    if USE_GRAPH:
        import graph_client
        children = graph_client.list_children_by_path(f"{GRAPH_BASE_PATH}/{sn}")
        names = [c["name"] for c in children if c.get("isFile")]
    else:
        d = os.path.join(LOCAL_BASE_DIR, sn)
        names = sorted(os.listdir(d)) if os.path.isdir(d) else []
    names.sort()
    return names


def save_photo(sn, filename, data):
    if len(data) > MAX_BYTES:
        raise PhotoTooLargeError("La foto pesa más de 15 MB.")
    sn = store.sanitize_sheet_name(sn)
    name = _safe_name(filename)
    if USE_GRAPH:
        import graph_client
        path = f"{GRAPH_BASE_PATH}/{sn}"
        graph_client.ensure_folder_path(path)
        graph_client.upload_bytes_by_path(f"{path}/{name}", data)
    else:
        with open(os.path.join(_local_dir(sn), name), "wb") as f:
            f.write(data)
    return name


def get_photo_bytes(sn, filename):
    sn = store.sanitize_sheet_name(sn)
    filename = os.path.basename(filename)
    if USE_GRAPH:
        import graph_client
        data = graph_client.download_bytes_by_path(f"{GRAPH_BASE_PATH}/{sn}/{filename}")
        if data is None:
            return None, None
        return data, _content_type(filename)
    path = os.path.join(LOCAL_BASE_DIR, sn, filename)
    if not os.path.isfile(path):
        return None, None
    with open(path, "rb") as f:
        return f.read(), _content_type(filename)


def delete_photo(sn, filename):
    sn = store.sanitize_sheet_name(sn)
    filename = os.path.basename(filename)
    if USE_GRAPH:
        import graph_client
        graph_client.delete_item_by_path(f"{GRAPH_BASE_PATH}/{sn}/{filename}")
    else:
        path = os.path.join(LOCAL_BASE_DIR, sn, filename)
        if os.path.isfile(path):
            os.remove(path)
