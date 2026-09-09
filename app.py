# -*- coding: utf-8 -*-
import os
import webbrowser
import threading

from flask import Flask, jsonify, request, render_template, session, redirect, url_for, Response

import excel_store as store
import photos_store
import alertas

try:
    import graph_auth
except ImportError:
    graph_auth = None

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "clave-de-desarrollo-solo-para-local")

APP_PASSWORD = os.environ.get("APP_PASSWORD")  # si no está definida, no se pide login (uso local)


def _error(msg, code=400):
    return jsonify({"error": msg}), code


@app.before_request
def _require_login():
    if not APP_PASSWORD:
        return None
    if request.endpoint in ("login", "static", "api_check_alertas"):
        return None
    if session.get("authed"):
        return None
    if request.path.startswith("/api/"):
        return _error("No has iniciado sesión.", 401)
    return redirect(url_for("login", next=request.path))


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        if request.form.get("password") == APP_PASSWORD:
            session["authed"] = True
            return redirect(request.args.get("next") or url_for("index"))
        error = "Contraseña incorrecta."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/")
def index():
    return render_template("index.html")


def _handle_store_errors(fn, *args, **kwargs):
    try:
        return fn(*args, **kwargs), None
    except store.SheetExistsError:
        return None, _error("Ya existe un equipo con ese número de serie.", 409)
    except store.SheetNotFoundError:
        return None, _error("Ese equipo ya no existe.", 404)
    except store.AveriaNotFoundError:
        return None, _error("Esa avería no existe.", 404)
    except photos_store.PhotoTooLargeError as e:
        return None, _error(str(e), 413)
    except ValueError as e:
        return None, _error(str(e) or "Datos inválidos.", 400)
    except FileNotFoundError as e:
        return None, _error(str(e) or "No se encuentra el archivo Excel.", 500)
    except PermissionError:
        return None, _error("El Excel está abierto en otro programa (ciérralo e inténtalo de nuevo).", 423)
    except (graph_auth.NotConfigured if graph_auth else Exception) as e:
        return None, _error("Configuración de Microsoft Graph incompleta: " + str(e), 500)
    except RuntimeError as e:
        return None, _error("Error accediendo a OneDrive: " + str(e), 502)


@app.route("/api/equipos", methods=["GET"])
def api_list():
    data, err = _handle_store_errors(store.list_equipos)
    return err if err else jsonify(data)


@app.route("/api/equipos", methods=["POST"])
def api_create():
    record = request.get_json(force=True) or {}
    if not record.get("sn") or not record.get("modelo"):
        return _error("Faltan el modelo o el número de serie.")
    data, err = _handle_store_errors(store.create_equipo, record)
    return err if err else (jsonify(data), 201)


@app.route("/api/equipos/<sn>", methods=["PUT"])
def api_update(sn):
    record = request.get_json(force=True) or {}
    if not record.get("sn") or not record.get("modelo"):
        return _error("Faltan el modelo o el número de serie.")
    data, err = _handle_store_errors(store.update_equipo, sn, record)
    return err if err else jsonify(data)


@app.route("/api/equipos/<sn>", methods=["DELETE"])
def api_delete(sn):
    data, err = _handle_store_errors(store.delete_equipo, sn)
    return err if err else jsonify({"ok": True})


@app.route("/api/equipos/<sn>/fotos", methods=["GET"])
def api_list_fotos(sn):
    data, err = _handle_store_errors(photos_store.list_photos, sn)
    return err if err else jsonify(data)


@app.route("/api/equipos/<sn>/fotos", methods=["POST"])
def api_upload_foto(sn):
    f = request.files.get("foto")
    if not f or not f.filename:
        return _error("No se ha recibido ninguna foto.")
    data, err = _handle_store_errors(photos_store.save_photo, sn, f.filename, f.read())
    return err if err else (jsonify({"name": data}), 201)


@app.route("/api/equipos/<sn>/fotos/<path:filename>", methods=["GET"])
def api_get_foto(sn, filename):
    result, err = _handle_store_errors(photos_store.get_photo_bytes, sn, filename)
    if err:
        return err
    data, content_type = result
    if data is None:
        return _error("Foto no encontrada.", 404)
    return Response(data, mimetype=content_type)


@app.route("/api/equipos/<sn>/fotos/<path:filename>", methods=["DELETE"])
def api_delete_foto(sn, filename):
    data, err = _handle_store_errors(photos_store.delete_photo, sn, filename)
    return err if err else jsonify({"ok": True})


@app.route("/api/equipos/<sn>/historial", methods=["GET"])
def api_list_historial(sn):
    data, err = _handle_store_errors(store.list_historial, sn)
    return err if err else jsonify(data)


@app.route("/api/equipos/<sn>/historial", methods=["POST"])
def api_add_historial(sn):
    body = request.get_json(force=True) or {}
    data, err = _handle_store_errors(store.add_historial_nota, sn, body.get("detalle") or "")
    return err if err else (jsonify(data), 201)


@app.route("/api/equipos/<sn>/averias", methods=["GET"])
def api_list_averias(sn):
    data, err = _handle_store_errors(store.list_averias, sn)
    return err if err else jsonify(data)


@app.route("/api/equipos/<sn>/averias", methods=["POST"])
def api_create_averia(sn):
    record = request.get_json(force=True) or {}
    data, err = _handle_store_errors(store.create_averia, sn, record)
    return err if err else (jsonify(data), 201)


@app.route("/api/equipos/<sn>/averias/<averia_id>", methods=["PUT"])
def api_update_averia(sn, averia_id):
    record = request.get_json(force=True) or {}
    data, err = _handle_store_errors(store.update_averia, sn, averia_id, record)
    return err if err else jsonify(data)


@app.route("/api/alertas/check", methods=["GET", "POST"])
def api_check_alertas():
    token = request.args.get("token") or request.headers.get("X-Alert-Token")
    expected = os.environ.get("ALERT_TOKEN")
    if not expected or token != expected:
        return _error("No autorizado.", 401)
    dry_run = request.args.get("dry_run") in ("1", "true", "yes")
    try:
        resultado = alertas.check_and_notify(dry_run=dry_run)
    except RuntimeError as e:
        return _error(str(e), 502)
    return jsonify(resultado)


def _open_browser():
    webbrowser.open("http://127.0.0.1:5000/")


if __name__ == "__main__":
    if not os.environ.get("MS_CLIENT_ID"):
        threading.Timer(1.0, _open_browser).start()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
