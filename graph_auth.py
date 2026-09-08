# -*- coding: utf-8 -*-
"""
Gestiona el token de acceso a Microsoft Graph a partir de un refresh token
guardado como variable de entorno (MS_REFRESH_TOKEN). Cada vez que Azure AD
emite un refresh token nuevo, se intenta persistir en el propio OneDrive
(_app_data/refresh_token.txt) para que el próximo arranque en frío use
siempre el más reciente, aunque la variable de entorno se quede desfasada.
"""
import os
import time
import threading

import msal

TENANT_ID = os.environ.get("MS_TENANT_ID", "")
CLIENT_ID = os.environ.get("MS_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("MS_CLIENT_SECRET", "")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Files.ReadWrite", "User.Read"]
REDIRECT_URI = "http://localhost:8400/callback"

REFRESH_TOKEN_ONEDRIVE_PATH = "_app_data/refresh_token.txt"

_lock = threading.RLock()
_state = {
    "refresh_token": os.environ.get("MS_REFRESH_TOKEN"),
    "access_token": None,
    "expires_at": 0,
    "tried_onedrive_sync": False,
}


class NotConfigured(Exception):
    pass


def confidential_app():
    if not (TENANT_ID and CLIENT_ID and CLIENT_SECRET):
        raise NotConfigured("Faltan MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET.")
    return msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)


def _exchange_refresh_token(rt):
    app = confidential_app()
    result = app.acquire_token_by_refresh_token(rt, scopes=SCOPES)
    if "access_token" not in result:
        raise RuntimeError("No se pudo renovar el token: " + str(result.get("error_description") or result))
    return result


def _maybe_sync_from_onedrive():
    """Best-effort: si hay un refresh token más reciente guardado en OneDrive, úsalo."""
    if _state["tried_onedrive_sync"]:
        return
    _state["tried_onedrive_sync"] = True
    try:
        import graph_client
        data = graph_client.download_bytes_by_path(REFRESH_TOKEN_ONEDRIVE_PATH)
        if data:
            stored = data.decode().strip()
            if stored and stored != _state["refresh_token"]:
                result = _exchange_refresh_token(stored)
                _state["refresh_token"] = stored
                _apply_result(result)
    except Exception as e:
        print("WARN: no se pudo sincronizar el refresh token desde OneDrive:", e)


def _persist_refresh_token(rt):
    try:
        import graph_client
        graph_client.upload_bytes_by_path(REFRESH_TOKEN_ONEDRIVE_PATH, rt.encode())
    except Exception as e:
        print("WARN: no se pudo guardar el refresh token en OneDrive:", e)


def _apply_result(result):
    _state["access_token"] = result["access_token"]
    _state["expires_at"] = time.time() + result.get("expires_in", 3600)
    new_rt = result.get("refresh_token")
    if new_rt and new_rt != _state["refresh_token"]:
        _state["refresh_token"] = new_rt
        _persist_refresh_token(new_rt)


def get_access_token():
    with _lock:
        if _state["access_token"] and time.time() < _state["expires_at"] - 60:
            return _state["access_token"]
        if not _state["refresh_token"]:
            raise NotConfigured(
                "No hay refresh token configurado (MS_REFRESH_TOKEN). "
                "Ejecuta bootstrap_auth.py para generarlo."
            )
        result = _exchange_refresh_token(_state["refresh_token"])
        _apply_result(result)
        _maybe_sync_from_onedrive()
        return _state["access_token"]
