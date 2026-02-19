# activos_proc.py
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from datetime import datetime, date, timedelta
import csv
import json
import re
import unicodedata

import openpyxl

from excel_com import norm_doc  # reutilizamos tu normalización de documento


_re_spaces = re.compile(r"\s+")


def _strip_accents(s: str) -> str:
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch)
    )


def norm_servicio(s) -> str:
    if s is None:
        return ""
    s = str(s)
    s = _strip_accents(s)
    s = s.upper().strip()
    s = _re_spaces.sub(" ", s)
    return s


def parse_fecha_usuario(s: str) -> date:
    s = (s or "").strip()
    if not s:
        raise ValueError("Fecha vacía")

    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    raise ValueError("Formato inválido. Usa YYYY-MM-DD o DD/MM/YYYY")


def norm_fecha_key(v) -> str:
    """
    Normaliza fechas desde Excel COM para dedupe.
    """
    if v is None or v == "":
        return ""
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()

    # Excel serial (a veces COM devuelve float)
    if isinstance(v, (int, float)):
        base = datetime(1899, 12, 30)
        try:
            d = (base + timedelta(days=float(v))).date()
            return d.isoformat()
        except Exception:
            return str(v).strip()

    s = str(v).strip()
    if not s:
        return ""
    # intenta parsear algunos formatos comunes
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            pass
    return s


@dataclass
class ActivoRow:
    rownum: int
    tipo_doc: str
    doc_raw: str
    doc_norm: str
    servicio_raw: str
    servicio_norm: str


@dataclass
class PlanRow:
    tipo_doc: str
    doc_norm: str
    fecha: date
    codigo: str               # va a H
    nombre_homologado: str    # va a K
    l_base: str               # va a L
    m_base: str               # va a M
    base_row: int             # fila base donde se extrajo L/M
    servicio_raw: str         # auditoría


def cargar_mapeo_activos(json_path: Path) -> dict[str, dict]:
    """
    Devuelve dict: entrada_norm -> {transformacion, codigo, entrada}
    """
    data = json.loads(json_path.read_text(encoding="utf-8"))
    out = {}
    for item in data:
        entrada = item.get("entrada", "")
        entrada_norm = norm_servicio(entrada)
        if not entrada_norm:
            continue
        out[entrada_norm] = {
            "entrada": entrada,
            "transformacion": (item.get("transformacion") or "").strip(),
            "codigo": (item.get("codigo") or "").strip(),
        }
    return out


def leer_activos_xlsx(xlsx_path: Path, sheet_name: str = "DETALLADO") -> list[ActivoRow]:
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"No existe hoja '{sheet_name}' en {xlsx_path.name}")
    ws = wb[sheet_name]

    rows: list[ActivoRow] = []
    # Columnas según tu archivo:
    # C=3 Tipo doc paciente, D=4 Documento, I=9 Nombre servicio
    for r in range(2, ws.max_row + 1):
        tipo = ws.cell(r, 3).value
        doc = ws.cell(r, 4).value
        serv = ws.cell(r, 9).value

        tipo_s = (str(tipo).strip() if tipo is not None else "")
        doc_s = (str(doc).strip() if doc is not None else "")
        doc_n = norm_doc(doc_s)

        serv_s = (str(serv).strip() if serv is not None else "")
        serv_n = norm_servicio(serv_s)

        if not tipo_s and not doc_s and not serv_s:
            continue  # fila vacía
        if not tipo_s or not doc_n or not serv_n:
            rows.append(
                ActivoRow(
                    rownum=r,
                    tipo_doc=tipo_s,
                    doc_raw=doc_s,
                    doc_norm=doc_n,
                    servicio_raw=serv_s,
                    servicio_norm=serv_n,
                )
            )
            continue

        rows.append(
            ActivoRow(
                rownum=r,
                tipo_doc=tipo_s,
                doc_raw=doc_s,
                doc_norm=doc_n,
                servicio_raw=serv_s,
                servicio_norm=serv_n,
            )
        )

    return rows


def construir_plan_activos(excel, activos: list[ActivoRow], mapeo: dict, fecha: date):
    """
    excel: instancia ExcelCOM ya abierta.
    Retorna: (plan: list[PlanRow], descartes: list[dict])
    """
    # 1) Index US (A tipo, B doc)
    us_keys = excel.cargar_us_keyset()

    # 2) Base ESTRUCTURA (E doc -> L y M primera ocurrencia)
    base_map = excel.cargar_estructura_base_lm()

    # 3) Dedupe actual en ESTRUCTURA: (doc|codigo|fecha)
    dupes = excel.cargar_estructura_dedupe_activos()

    fecha_key = fecha.isoformat()

    plan: list[PlanRow] = []
    descartes: list[dict] = []

    def reject(a: ActivoRow, reason: str, extra: str = ""):
        descartes.append(
            {
                "row_excel": a.rownum,
                "tipo_doc": a.tipo_doc,
                "doc_norm": a.doc_norm,
                "servicio": a.servicio_raw,
                "reason": reason,
                "extra": extra,
            }
        )

    for a in activos:
        if not a.tipo_doc or not a.doc_norm:
            reject(a, "DOC_O_TIPO_VACIO")
            continue
        if not a.servicio_norm:
            reject(a, "SERVICIO_VACIO")
            continue

        key_us = f"{a.tipo_doc}|{a.doc_norm}"
        if key_us not in us_keys:
            reject(a, "NO_EXISTE_EN_US")
            continue

        m = mapeo.get(a.servicio_norm)
        if not m:
            reject(a, "NO_MAPEO_EN_JSON", a.servicio_norm)
            continue

        nombre = (m.get("transformacion") or "").strip()
        codigo = (m.get("codigo") or "").strip()

        if not nombre:
            reject(a, "SIN_NOMBRE_HOMOLOGADO")
            continue

        if nombre.strip().upper() == "QUITAR":
            reject(a, "SERVICIO_EXCLUIDO")
            continue

        if not codigo:
            reject(a, "SIN_CODIGO_HOMOLOGADO")
            continue

        base = base_map.get(a.doc_norm)
        if not base:
            reject(a, "NO_BASE_ESTRUCTURA")
            continue

        l_base = base.get("L", "")
        m_base = base.get("M", "")
        base_row = base.get("row", 0)

        if (l_base is None or str(l_base).strip() == "") or (m_base is None or str(m_base).strip() == ""):
            reject(a, "BASE_SIN_LM", f"L='{l_base}' M='{m_base}'")
            continue

        key_dupe = f"{a.doc_norm}|{codigo}|{fecha_key}"
        if key_dupe in dupes:
            reject(a, "DUPLICADO_EN_ESTRUCTURA", key_dupe)
            continue

        plan.append(
            PlanRow(
                tipo_doc=a.tipo_doc,
                doc_norm=a.doc_norm,
                fecha=fecha,
                codigo=codigo,                    # H
                nombre_homologado=nombre,          # K
                l_base=str(l_base),
                m_base=str(m_base),
                base_row=int(base_row) if base_row else 0,
                servicio_raw=a.servicio_raw,
            )
        )

    return plan, descartes


def exportar_auditoria_csv(out_path: Path, plan: list[PlanRow], descartes: list[dict]):
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["TIPO", "tipo_doc", "doc", "fecha", "codigo", "nombre_homologado", "L", "M", "base_row", "servicio", "reason", "extra", "row_excel"])
        for p in plan:
            w.writerow(["OK", p.tipo_doc, p.doc_norm, p.fecha.isoformat(), p.codigo, p.nombre_homologado, p.l_base, p.m_base, p.base_row, p.servicio_raw, "", "", ""])
        for d in descartes:
            w.writerow(["NO", d.get("tipo_doc",""), d.get("doc_norm",""), "", "", "", "", "", "", d.get("servicio",""), d.get("reason",""), d.get("extra",""), d.get("row_excel","")])
