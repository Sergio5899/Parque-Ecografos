# -*- coding: utf-8 -*-
"""
Ejecuta esto UNA sola vez, en tu PC, para autorizar la app a acceder a tu
OneDrive. Abre tu navegador, inicias sesión con tu cuenta de Microsoft y
aceptas los permisos; el script recoge el resultado y te muestra el
refresh token que hay que guardar como variable de entorno MS_REFRESH_TOKEN
en el servicio de hosting (Render).

Antes de ejecutar, define estas variables de entorno (o edítalas aquí abajo):
  MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET
"""
import os
import sys
import webbrowser
import http.server
import urllib.parse
import threading

import msal

TENANT_ID = os.environ.get("MS_TENANT_ID") or input("MS_TENANT_ID: ").strip()
CLIENT_ID = os.environ.get("MS_CLIENT_ID") or input("MS_CLIENT_ID: ").strip()
CLIENT_SECRET = os.environ.get("MS_CLIENT_SECRET") or input("MS_CLIENT_SECRET: ").strip()

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Files.ReadWrite", "User.Read"]
REDIRECT_URI = "http://localhost:8400/callback"

_result = {}


class Handler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        qs = urllib.parse.urlparse(self.path).query
        params = urllib.parse.parse_qs(qs)
        if "code" in params:
            _result["code"] = params["code"][0]
            body = "Autorización recibida. Ya puedes cerrar esta pestaña y volver a la terminal."
        else:
            body = "No se recibió código de autorización. Revisa la terminal."
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(f"<html><body><h2>{body}</h2></body></html>".encode())

    def log_message(self, format, *args):
        pass


def main():
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    auth_url = app.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI)

    server = http.server.HTTPServer(("localhost", 8400), Handler)
    t = threading.Thread(target=server.handle_request, daemon=True)
    t.start()

    print("Abriendo el navegador para iniciar sesión con tu cuenta Microsoft...")
    webbrowser.open(auth_url)
    print("Esperando autorización (inicia sesión y acepta los permisos en la ventana del navegador)...")
    t.join(timeout=180)

    if "code" not in _result:
        print("No se recibió el código de autorización a tiempo. Vuelve a intentarlo.")
        sys.exit(1)

    result = app.acquire_token_by_authorization_code(
        _result["code"], scopes=SCOPES, redirect_uri=REDIRECT_URI
    )
    if "access_token" not in result:
        print("ERROR:", result.get("error_description") or result)
        sys.exit(1)

    refresh_token = result.get("refresh_token")
    print("\n=== LISTO ===")
    print("Guarda este valor como variable de entorno MS_REFRESH_TOKEN en Render:\n")
    print(refresh_token)
    print("\n(También se ha intentado subir una copia a tu OneDrive en _app_data/refresh_token.txt)")

    try:
        import graph_auth, graph_client
        graph_auth._state["refresh_token"] = refresh_token
        graph_auth._state["access_token"] = result["access_token"]
        import time as _t
        graph_auth._state["expires_at"] = _t.time() + result.get("expires_in", 3600)
        graph_client.upload_bytes_by_path("_app_data/refresh_token.txt", refresh_token.encode())
        print("Copia subida a OneDrive correctamente.")
    except Exception as e:
        print("Aviso: no se pudo subir la copia a OneDrive (no es crítico):", e)


if __name__ == "__main__":
    main()
