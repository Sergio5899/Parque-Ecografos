# -*- coding: utf-8 -*-
"""
Lee y escribe directamente sobre "ECOS DATOS.xlsx" (una hoja por equipo).
Este módulo conoce la plantilla exacta de la hoja y solo toca las celdas
de datos, dejando intactos formato, colores, anchos de columna, etc.
"""
import io
import os
import re
import threading
from datetime import date, datetime
from copy import copy

import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_PATH = os.path.join(BASE_DIR, "ECOS DATOS.xlsx")

# Si hay credenciales de Microsoft Graph configuradas, se lee/escribe el
# Excel directamente en OneDrive en vez de en un archivo local.
USE_GRAPH = bool(os.environ.get("MS_CLIENT_ID"))
GRAPH_FILENAME = "ECOS DATOS.xlsx"
_graph_item = {"drive_id": None, "item_id": None}

_lock = threading.Lock()

COMPONENTE_LABELS = ["Power Regulator", "ACB", "PC / Computer Assembly", "Acquisition Module"]
COMPONENTE_ROWS = [10, 11, 12, 13]  # A=label(no tocar), B=version, C=maxSw, D=12nc

SONDAS_HEADER_ROW = 1
SONDAS_FIRST_ROW = 2
SONDAS_LAST_ROW = 7  # fila 8 ya es la cabecera "DICOM LOCAL"
SONDAS_COLS = {"nombre": 6, "serie": 7, "codigo12nc": 8, "pass": 9}  # F,G,H,I

DICOM_ROWS = {"ip": 9, "mask": 10, "gateway": 11, "aet": 12, "port": 13}
DICOM_VALUE_COL = 7  # G
DVDPC_ROW = 14
DVDPC_COL = 7

NOTASDICOM_ROW = 9
NOTASDICOM_COL = 9  # I

NODOS_HEADER_ROW = 19
NODOS_FIRST_ROW = 20
NODOS_LAST_ROW = 29
NODOS_COLS = {"nombre": 6, "aet": 7, "ip": 8, "puerto": 9}  # F,G,H,I

FIELD_ROWS = {
    "modelo": 1, "sn": 2, "hospital": 3, "swVersion": 4, "hwVersion": 5, "notas": 6,
}
VALUE_COL = 2  # B

FECHA_INSTALACION_ROW = 19
MANTENIMIENTO_ROWS = {
    "ultimaRevisionPM": 20, "ultimoBackup": 21, "ultimoTestGeneral": 22,
    "ultimoTestSondas": 23, "ultimoTestBaterias": 24, "borrarAvisosMtto": 25,
}

INVALID_SHEET_CHARS = re.compile(r"[\\/*?:\[\]]")


def _is_real(v):
    if v is None:
        return False
    s = str(v).strip()
    return s != "" and s != "-"


def _clean(v):
    return str(v).strip() if _is_real(v) else None


def _cell_to_date_str(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    return _clean(v)


def _date_str_to_value(s):
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        return s


def sanitize_sheet_name(sn):
    name = INVALID_SHEET_CHARS.sub("", (sn or "").strip())
    return name[:31]


def _resolve_graph_item():
    if _graph_item["drive_id"] is None:
        import graph_client
        found = graph_client.find_item_by_name(GRAPH_FILENAME)
        if not found:
            raise FileNotFoundError(f"No se encuentra '{GRAPH_FILENAME}' en el OneDrive.")
        _graph_item["drive_id"], _graph_item["item_id"], _ = found
    return _graph_item["drive_id"], _graph_item["item_id"]


def _load():
    if USE_GRAPH:
        import graph_client
        drive_id, item_id = _resolve_graph_item()
        data = graph_client.download_bytes(drive_id, item_id)
        return openpyxl.load_workbook(io.BytesIO(data))
    return openpyxl.load_workbook(EXCEL_PATH)


def _save(wb):
    if USE_GRAPH:
        import graph_client
        drive_id, item_id = _resolve_graph_item()
        buf = io.BytesIO()
        wb.save(buf)
        graph_client.upload_bytes(drive_id, item_id, buf.getvalue())
    else:
        wb.save(EXCEL_PATH)


def _parse_sheet(ws):
    sn_cell = ws.cell(row=FIELD_ROWS["sn"], column=VALUE_COL).value
    sn = _clean(sn_cell) or ws.title

    componentes = []
    for i, row in enumerate(COMPONENTE_ROWS):
        componentes.append({
            "nombre": COMPONENTE_LABELS[i],
            "version": _clean(ws.cell(row=row, column=2).value),
            "maxSw": _clean(ws.cell(row=row, column=3).value),
            "codigo12nc": _clean(ws.cell(row=row, column=4).value),
        })

    sondas = []
    for row in range(SONDAS_FIRST_ROW, SONDAS_LAST_ROW + 1):
        nombre = ws.cell(row=row, column=SONDAS_COLS["nombre"]).value
        serie = ws.cell(row=row, column=SONDAS_COLS["serie"]).value
        codigo = ws.cell(row=row, column=SONDAS_COLS["codigo12nc"]).value
        pass_ = ws.cell(row=row, column=SONDAS_COLS["pass"]).value
        if any(_is_real(v) for v in (nombre, serie, codigo, pass_)):
            sondas.append({
                "nombre": _clean(nombre), "serie": _clean(serie),
                "codigo12nc": _clean(codigo), "pass": _clean(pass_),
            })

    nodos = []
    for row in range(NODOS_FIRST_ROW, NODOS_LAST_ROW + 1):
        nombre = ws.cell(row=row, column=NODOS_COLS["nombre"]).value
        aet = ws.cell(row=row, column=NODOS_COLS["aet"]).value
        ip = ws.cell(row=row, column=NODOS_COLS["ip"]).value
        puerto = ws.cell(row=row, column=NODOS_COLS["puerto"]).value
        if any(_is_real(v) for v in (nombre, aet, ip, puerto)):
            nodos.append({
                "nombre": _clean(nombre), "aet": _clean(aet),
                "ip": _clean(ip), "puerto": _clean(puerto),
            })

    dicom = {}
    for key, row in DICOM_ROWS.items():
        dicom[key] = _clean(ws.cell(row=row, column=DICOM_VALUE_COL).value)
    dicom["dvdPc"] = _clean(ws.cell(row=DVDPC_ROW, column=DVDPC_COL).value)

    record = {
        "sn": sn,
        "modelo": _clean(ws.cell(row=FIELD_ROWS["modelo"], column=VALUE_COL).value),
        "hospital": _clean(ws.cell(row=FIELD_ROWS["hospital"], column=VALUE_COL).value),
        "swVersion": _clean(ws.cell(row=FIELD_ROWS["swVersion"], column=VALUE_COL).value),
        "hwVersion": _clean(ws.cell(row=FIELD_ROWS["hwVersion"], column=VALUE_COL).value),
        "notas": _clean(ws.cell(row=FIELD_ROWS["notas"], column=VALUE_COL).value),
        "fechaInstalacion": _cell_to_date_str(ws.cell(row=FECHA_INSTALACION_ROW, column=VALUE_COL).value),
        "notasDicom": _clean(ws.cell(row=NOTASDICOM_ROW, column=NOTASDICOM_COL).value),
        "componentes": componentes,
        "sondas": sondas,
        "nodosDicom": nodos,
        "dicom": dicom,
        "_sheet": ws.title,
    }
    for key, row in MANTENIMIENTO_ROWS.items():
        record[key] = _cell_to_date_str(ws.cell(row=row, column=VALUE_COL).value)
    return record


def list_equipos():
    with _lock:
        wb = _load()
        try:
            return [_parse_sheet(wb[name]) for name in wb.sheetnames]
        finally:
            wb.close()


def get_equipo(sn):
    with _lock:
        wb = _load()
        try:
            if sn not in wb.sheetnames:
                return None
            return _parse_sheet(wb[sn])
        finally:
            wb.close()


def _clear_and_write_data(ws, record):
    ws.cell(row=FIELD_ROWS["modelo"], column=VALUE_COL).value = record.get("modelo")
    ws.cell(row=FIELD_ROWS["sn"], column=VALUE_COL).value = record.get("sn")
    ws.cell(row=FIELD_ROWS["hospital"], column=VALUE_COL).value = record.get("hospital")
    ws.cell(row=FIELD_ROWS["swVersion"], column=VALUE_COL).value = record.get("swVersion")
    ws.cell(row=FIELD_ROWS["hwVersion"], column=VALUE_COL).value = record.get("hwVersion")
    ws.cell(row=FIELD_ROWS["notas"], column=VALUE_COL).value = record.get("notas")

    ws.cell(row=FECHA_INSTALACION_ROW, column=VALUE_COL).value = _date_str_to_value(record.get("fechaInstalacion"))
    for key, row in MANTENIMIENTO_ROWS.items():
        ws.cell(row=row, column=VALUE_COL).value = _date_str_to_value(record.get(key))

    componentes = record.get("componentes") or []
    for i, row in enumerate(COMPONENTE_ROWS):
        c = componentes[i] if i < len(componentes) else {}
        ws.cell(row=row, column=2).value = c.get("version")
        ws.cell(row=row, column=3).value = c.get("maxSw")
        ws.cell(row=row, column=4).value = c.get("codigo12nc")

    for row in range(SONDAS_FIRST_ROW, SONDAS_LAST_ROW + 1):
        for col in SONDAS_COLS.values():
            ws.cell(row=row, column=col).value = None
    sondas = (record.get("sondas") or [])[:SONDAS_LAST_ROW - SONDAS_FIRST_ROW + 1]
    for i, s in enumerate(sondas):
        row = SONDAS_FIRST_ROW + i
        ws.cell(row=row, column=SONDAS_COLS["nombre"]).value = s.get("nombre")
        ws.cell(row=row, column=SONDAS_COLS["serie"]).value = s.get("serie")
        ws.cell(row=row, column=SONDAS_COLS["codigo12nc"]).value = s.get("codigo12nc")
        ws.cell(row=row, column=SONDAS_COLS["pass"]).value = s.get("pass")

    dicom = record.get("dicom") or {}
    for key, row in DICOM_ROWS.items():
        cell = ws.cell(row=row, column=DICOM_VALUE_COL)
        cell.value = dicom.get(key)
        cell.number_format = "@"
    ws.cell(row=DVDPC_ROW, column=DVDPC_COL).value = dicom.get("dvdPc")

    ws.cell(row=NOTASDICOM_ROW, column=NOTASDICOM_COL).value = record.get("notasDicom")

    for row in range(NODOS_FIRST_ROW, NODOS_LAST_ROW + 1):
        for col in NODOS_COLS.values():
            ws.cell(row=row, column=col).value = None
    nodos = (record.get("nodosDicom") or [])[:NODOS_LAST_ROW - NODOS_FIRST_ROW + 1]
    for i, n in enumerate(nodos):
        row = NODOS_FIRST_ROW + i
        ws.cell(row=row, column=NODOS_COLS["nombre"]).value = n.get("nombre")
        ws.cell(row=row, column=NODOS_COLS["aet"]).value = n.get("aet")
        ws.cell(row=row, column=NODOS_COLS["ip"]).value = n.get("ip")
        ws.cell(row=row, column=NODOS_COLS["puerto"]).value = n.get("puerto")


def _new_sheet_from_template(wb, sheet_name):
    template = wb[wb.sheetnames[0]]
    ws = wb.copy_worksheet(template)
    ws.title = sheet_name
    for i, row in enumerate(COMPONENTE_ROWS):
        ws.cell(row=row, column=1).value = COMPONENTE_LABELS[i]
    return ws


class SheetExistsError(Exception):
    pass


class SheetNotFoundError(Exception):
    pass


def create_equipo(record):
    sn = sanitize_sheet_name(record.get("sn"))
    if not sn:
        raise ValueError("Número de serie vacío.")
    with _lock:
        wb = _load()
        try:
            if sn in wb.sheetnames:
                raise SheetExistsError(sn)
            ws = _new_sheet_from_template(wb, sn)
            record = dict(record, sn=sn)
            _clear_and_write_data(ws, record)
            _save(wb)
        finally:
            wb.close()
    return get_equipo(sn)


def update_equipo(original_sn, record):
    new_sn = sanitize_sheet_name(record.get("sn")) or original_sn
    with _lock:
        wb = _load()
        try:
            if original_sn not in wb.sheetnames:
                raise SheetNotFoundError(original_sn)
            if new_sn != original_sn and new_sn in wb.sheetnames:
                raise SheetExistsError(new_sn)
            ws = wb[original_sn]
            record = dict(record, sn=new_sn)
            _clear_and_write_data(ws, record)
            if new_sn != original_sn:
                ws.title = new_sn
            _save(wb)
        finally:
            wb.close()
    return get_equipo(new_sn)


def delete_equipo(sn):
    with _lock:
        wb = _load()
        try:
            if sn not in wb.sheetnames:
                raise SheetNotFoundError(sn)
            del wb[sn]
            _save(wb)
        finally:
            wb.close()
