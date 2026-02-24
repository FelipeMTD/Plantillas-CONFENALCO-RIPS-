# activos_proc.py
from __future__ import annotations
from dataclasses import dataclass
from pathlib import Path
from datetime import datetime, date
import csv, json, re, unicodedata, openpyxl
from excel_com import norm_doc

def norm_servicio(s) -> str:
    if s is None: return ""
    s = "".join(ch for ch in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s).upper().strip()

def parse_fecha_usuario(s: str) -> date:
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try: return datetime.strptime(s, fmt).date()
        except ValueError: pass
    raise ValueError("Formato inválido. Use YYYY-MM-DD o DD/MM/YYYY")

@dataclass
class ActivoRow:
    rownum: int; tipo_doc: str; doc_raw: str; doc_norm: str; servicio_raw: str; servicio_norm: str

@dataclass
class PlanRow:
    tipo_doc: str; doc_norm: str; fecha: date; codigo: str; nombre_homologado: str; l_base: str; m_base: str; base_row: int; servicio_raw: str

# CORRECCIÓN: Se agrega sheet_name para coincidir con la llamada en main.py
def leer_activos_xlsx(xlsx_path: Path, sheet_name: str = "DETALLADO") -> list[ActivoRow]:
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"No existe hoja '{sheet_name}' en {xlsx_path.name}")
    ws = wb[sheet_name]
    
    rows = []
    for r in range(2, ws.max_row + 1):
        doc = ws.cell(r, 2).value
        serv = ws.cell(r, 5).value
        if doc or serv:
            rows.append(ActivoRow(r, "", str(doc or "").strip(), norm_doc(doc), str(serv or "").strip(), norm_servicio(serv)))
    return rows

def cargar_mapeo_activos(json_path: Path) -> dict:
    data = json.loads(json_path.read_text(encoding="utf-8"))
    return {norm_servicio(i["entrada"]): {"transformacion": (i.get("transformacion") or "").strip(), "codigo": (i.get("codigo") or "").strip()} for i in data if norm_servicio(i.get("entrada"))}

def construir_plan_activos(excel, activos, mapeo, fecha):
    doc_to_tipo = {k.split("|")[1]: k.split("|")[0] for k in excel.cargar_us_keyset() if "|" in k}
    base_map, dupes = excel.cargar_estructura_base_lm(), excel.cargar_estructura_dedupe_activos()
    plan, descartes = [], []
    for a in activos:
        tipo = doc_to_tipo.get(a.doc_norm)
        m = mapeo.get(a.servicio_norm)
        base = base_map.get(a.doc_norm)
        if not a.doc_norm or not a.servicio_norm: reason = "DOC_O_SERV_VACIO"
        elif not tipo: reason = "NO_EXISTE_EN_US"
        elif not m or m.get("transformacion") == "QUITAR": reason = "MAPEO_INVALIDO"
        elif not base or not str(base["L"]).strip() or not str(base["M"]).strip(): reason = "SIN_BASE_LM"
        elif f"{a.doc_norm}|{m['codigo']}|{fecha.isoformat()}" in dupes: reason = "DUPLICADO"
        else:
            plan.append(PlanRow(tipo, a.doc_norm, fecha, m["codigo"], m["transformacion"], str(base["L"]), str(base["M"]), base["row"], a.servicio_raw))
            continue
        descartes.append({"row_excel": a.rownum, "reason": reason, "servicio": a.servicio_raw})
    return plan, descartes

def exportar_auditoria_csv(path, plan, descartes):
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["TIPO", "doc", "codigo", "L", "M", "servicio", "reason"])
        for p in plan: w.writerow(["OK", p.doc_norm, p.codigo, p.l_base, p.m_base, p.servicio_raw, ""])
        for d in descartes: w.writerow(["NO", "", "", "", "", d["servicio"], d["reason"]])