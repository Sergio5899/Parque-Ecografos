# -*- coding: utf-8 -*-
"""
Comprueba qué equipos tienen la revisión de mantenimiento (PM) atrasada según
su periodicidad configurada (campo "periodicidadMeses" de cada equipo; si no
se ha rellenado se asume anual), y envía un email cuando un equipo *pasa* a
estar atrasado. No repite el aviso cada día mientras siga atrasado; si se
pone al día y luego se vuelve a atrasar, avisa de nuevo.

Pensado para ser llamado periódicamente (p. ej. una vez al día) por un cron
externo contra /api/alertas/check.
"""
import os
import smtplib
from email.mime.text import MIMEText
from datetime import date, datetime

import excel_store as store

DEFAULT_PERIODICIDAD_MESES = 12


def _dias_desde(iso):
    if not iso:
        return None
    try:
        d = datetime.strptime(iso, "%Y-%m-%d").date()
    except ValueError:
        return None
    return (date.today() - d).days


def _periodicidad_dias(equipo):
    meses = equipo.get("periodicidadMeses")
    try:
        meses = int(meses)
    except (TypeError, ValueError):
        meses = DEFAULT_PERIODICIDAD_MESES
    if meses <= 0:
        meses = DEFAULT_PERIODICIDAD_MESES
    return meses * 30


def esta_atrasado(equipo):
    dias = _dias_desde(equipo.get("ultimaRevisionPM"))
    if dias is None:
        return False
    return dias > _periodicidad_dias(equipo)


def check_and_notify(dry_run=False):
    """Devuelve {"atrasados": [...], "nuevos": [...], "resueltos": [...]}.
    Solo envía email (y solo persiste estado) si dry_run es False."""
    equipos = store.list_equipos()
    por_sn = {e["sn"]: e for e in equipos}
    atrasados_ahora = {sn for sn, e in por_sn.items() if esta_atrasado(e)}

    ya_alertados = {sn for sn, alertado in store.get_alertas_estado().items() if alertado}
    nuevos = sorted(atrasados_ahora - ya_alertados)
    resueltos = sorted(ya_alertados - atrasados_ahora)

    if nuevos and not dry_run:
        _enviar_email([por_sn[sn] for sn in nuevos])

    if not dry_run:
        estado = {sn: True for sn in atrasados_ahora}
        estado.update({sn: False for sn in resueltos})
        if estado:
            store.set_alertas_estado(estado)

    return {"atrasados": sorted(atrasados_ahora), "nuevos": nuevos, "resueltos": resueltos}


def _enviar_email(equipos):
    host = os.environ.get("SMTP_HOST")
    user = os.environ.get("SMTP_USER")
    password = os.environ.get("SMTP_PASS")
    to_addr = os.environ.get("ALERT_EMAIL_TO") or user
    if not (host and user and password and to_addr):
        raise RuntimeError("Faltan variables SMTP_HOST / SMTP_USER / SMTP_PASS / ALERT_EMAIL_TO.")
    port = int(os.environ.get("SMTP_PORT", "587"))

    lineas = []
    for e in equipos:
        hosp = e.get("hospital") or "sin hospital asignado"
        pm = e.get("ultimaRevisionPM") or "sin revisión registrada"
        lineas.append(f"- {e.get('modelo') or 'Equipo'} (S/N {e['sn']}) · {hosp} · última PM: {pm}")

    asunto = f"Parque de Ecógrafos: {len(equipos)} equipo(s) con revisión atrasada"
    cuerpo = "Estos equipos han pasado a tener la revisión de mantenimiento atrasada:\n\n" + "\n".join(lineas)

    msg = MIMEText(cuerpo, "plain", "utf-8")
    msg["Subject"] = asunto
    msg["From"] = user
    msg["To"] = to_addr

    with smtplib.SMTP(host, port, timeout=20) as server:
        server.starttls()
        server.login(user, password)
        server.sendmail(user, [to_addr], msg.as_string())
