# excel_com.py
import win32com.client as win32
from pathlib import Path
import re, math
from datetime import datetime, date, timedelta

XL_UP = -4162
CONTROL_SHEET = "__RIPS_CONTROL__"
_re_non_digits = re.compile(r"\D+")

def norm_doc(v):
    if v is None: return ""
    if isinstance(v, bool): return ""
    if isinstance(v, (int, float)):
        if isinstance(v, float) and not math.isfinite(v): return ""
        return str(int(round(v)))
    s = str(v).strip()
    return _re_non_digits.sub("", s.split('.')[0]) if s else ""

def _norm_fecha_key(v) -> str:
    if not v: return ""
    if isinstance(v, (datetime, date)): return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)):
        return (datetime(1899, 12, 30) + timedelta(days=float(v))).strftime("%Y-%m-%d")
    return str(v).strip()

class ExcelCOM:
    def __init__(self, path_xlsm: Path):
        self.path = str(path_xlsm.resolve())
        self.excel = None; self.wb = None
        self.ws_estructura = None; self.ws_us = None; self.ws_control = None; self.seen_us = set()

    def abrir(self):
        self.excel = win32.DispatchEx("Excel.Application")
        self.excel.Visible = False; self.excel.DisplayAlerts = False
        self.wb = self.excel.Workbooks.Open(self.path)
        self.ws_estructura = self.wb.Worksheets("ESTRUCTURA")
        self.ws_us = self.wb.Worksheets("US")
        self._init_control(); self._load_seen_us()

    def cerrar(self):
        if self.wb: self.wb.Save(); self.wb.Close(SaveChanges=False)
        if self.excel: self.excel.Quit()

    def _init_control(self):
        try: self.ws_control = self.wb.Worksheets(CONTROL_SHEET)
        except:
            self.ws_control = self.wb.Worksheets.Add(); self.ws_control.Name = CONTROL_SHEET
            self.ws_control.Cells(1, 1).Value = "KIND"; self.ws_control.Cells(1, 2).Value = "KEY"
        self.ws_control.Visible = 2

    def _load_seen_us(self):
        row = 2
        while self.ws_control.Cells(row, 1).Value:
            if self.ws_control.Cells(row, 1).Value == "U": self.seen_us.add(str(self.ws_control.Cells(row, 2).Value))
            row += 1

    def siguiente_fila(self, ws, col):
        # LÓGICA: Empezar en fila 3 si no hay datos después del header
        last_row = ws.Cells(ws.Rows.Count, col).End(XL_UP).Row
        return max(3, last_row + 1)

    def ultima_fila(self, ws, col):
        return max(1, int(ws.Cells(ws.Rows.Count, col).End(XL_UP).Row))

    def pegar_estructura_rango(self, filas, fila_inicio):
        if not filas: return fila_inicio
        end = fila_inicio + len(filas) - 1
        self.ws_estructura.Range(f"E{fila_inicio}:L{end}").Value = filas
        return end + 1

    def pegar_us_rango(self, filas, fila_inicio):
        nuevos = []
        for r in filas:
            doc = norm_doc(r[1])
            if doc and f"{r[0]}|{doc}" not in self.seen_us:
                r[1] = doc; nuevos.append(r[:14] + [""] * (14 - len(r))); self.seen_us.add(f"{r[0]}|{doc}")
        if not nuevos: return fila_inicio
        end = fila_inicio + len(nuevos) - 1
        self.ws_us.Range(f"A{fila_inicio}:N{end}").Value = nuevos
        start_ctrl = self.ws_control.Cells(self.ws_control.Rows.Count, 1).End(XL_UP).Row + 1
        self.ws_control.Range(f"A{start_ctrl}:B{start_ctrl + len(nuevos) - 1}").Value = [["U", f"{r[0]}|{r[1]}"] for r in nuevos]
        return end + 1

    def cargar_us_keyset(self):
        last = self.ultima_fila(self.ws_us, 2)
        if last < 2: return set()
        rng = self.ws_us.Range(f"A2:B{last}").Value
        return {f"{str(r[0]).strip()}|{norm_doc(r[1])}" for r in rng if r[0] and r[1]} if rng else set()

    def cargar_estructura_base_lm(self):
        last = self.ultima_fila(self.ws_estructura, 5)
        if last < 2: return {}
        rng = self.ws_estructura.Range(f"E2:M{last}").Value
        out = {}
        if rng:
            for i, row in enumerate(rng, 2):
                doc = norm_doc(row[0])
                if doc and doc not in out: out[doc] = {"row": i, "L": row[7], "M": row[8]}
        return out

    def cargar_estructura_dedupe_activos(self):
        last = self.ultima_fila(self.ws_estructura, 5)
        if last < 2: return set()
        rng = self.ws_estructura.Range(f"E2:H{last}").Value
        return {f"{norm_doc(r[0])}|{str(r[3]).strip()}|{_norm_fecha_key(r[1])}" for r in rng if r[0] and r[1] and r[3]} if rng else set()

    def pegar_activos_estructura(self, plan_rows, fila_inicio):
        if not plan_rows: return fila_inicio
        data = []
        for p in plan_rows:
            data.append([p.tipo_doc, p.doc_norm, p.fecha.strftime("%d/%m/%Y"), "", p.codigo, "", "", p.nombre_homologado, p.l_base, p.m_base])
        end = fila_inicio + len(data) - 1
        self.ws_estructura.Range(f"D{fila_inicio}:M{end}").Value = data
        return end + 1