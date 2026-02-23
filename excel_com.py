# excel_com.py
import win32com.client as win32
from pathlib import Path
import re
import math
from datetime import datetime, date, timedelta

XL_UP = -4162
CONTROL_SHEET = "__RIPS_CONTROL__"
_re_non_digits = re.compile(r"\D+")


def norm_doc(v):
    if v is None:
        return ""

    # Excel COM suele devolver floats (1234567.0)
    if isinstance(v, bool):
        return ""

    if isinstance(v, int):
        return str(v)

    if isinstance(v, float):
        if not math.isfinite(v):
            return ""
        # Si es entero en float -> conviértelo a int (quita el .0)
        if v.is_integer():
            return str(int(v))
        # Si por alguna razón trae decimales, redondea (documentos no deberían)
        return str(int(round(v)))

    s = str(v).strip()
    if not s:
        return ""

    # Caso string tipo "1234567.0"
    m = re.fullmatch(r"(\d+)\.0+", s)
    if m:
        return m.group(1)

    # Caso notación científica "1.050547764E9"
    try:
        f = float(s)
        if math.isfinite(f) and abs(f - round(f)) < 1e-6:
            return str(int(round(f)))
    except Exception:
        pass

    # fallback: solo dígitos
    return _re_non_digits.sub("", s)

def _norm_fecha_key(v) -> str:
    if v is None or v == "":
        return ""
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
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
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            pass
    return s


class ExcelCOM:
    def __init__(self, path_xlsm: Path):
        self.path = str(path_xlsm.resolve())
        self.excel = None
        self.wb = None
        self.ws_estructura = None
        self.ws_us = None
        self.ws_control = None
        self.seen_us = set()

    # -------------------------
    # Apertura / cierre
    # -------------------------
    def abrir(self):
        self.excel = win32.DispatchEx("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False

        self.wb = self.excel.Workbooks.Open(self.path)

        self.ws_estructura = self.wb.Worksheets("ESTRUCTURA")
        self.ws_us = self.wb.Worksheets("US")

        self._init_control()
        self._load_seen_us()

    def cerrar(self):
        if self.wb:
            self.wb.Save()
            self.wb.Close(SaveChanges=False)
        if self.excel:
            self.excel.Quit()

    # -------------------------
    # Control US
    # -------------------------
    def _init_control(self):
        try:
            self.ws_control = self.wb.Worksheets(CONTROL_SHEET)
        except Exception:
            self.ws_control = self.wb.Worksheets.Add()
            self.ws_control.Name = CONTROL_SHEET
            self.ws_control.Cells(1, 1).Value = "KIND"
            self.ws_control.Cells(1, 2).Value = "KEY"
            self.ws_control.Visible = 2

        self.ws_control.Visible = 2

    def _load_seen_us(self):
        row = 2
        while True:
            kind = self.ws_control.Cells(row, 1).Value
            key = self.ws_control.Cells(row, 2).Value
            if not kind:
                break
            if kind == "U" and key:
                self.seen_us.add(str(key))
            row += 1

    def append_us_control_batch(self, docs):
        if not docs:
            return
        start = self.ws_control.Cells(self.ws_control.Rows.Count, 1).End(XL_UP).Row + 1
        data = [["U", d] for d in docs]
        end = start + len(data) - 1
        self.ws_control.Range(f"A{start}:B{end}").Value = data

    # -------------------------
    # Utilidades
    # -------------------------
    def siguiente_fila(self, ws, col):
        return ws.Cells(ws.Rows.Count, col).End(XL_UP).Row + 1

    def ultima_fila(self, ws, col):
        last = ws.Cells(ws.Rows.Count, col).End(XL_UP).Row
        return max(1, int(last))

    # -------------------------
    # ESTRUCTURA por RANGOS (ZIP)
    # -------------------------
    def pegar_estructura_rango(self, filas, fila_inicio):
        """
        filas: lista de listas de longitud fija (columnas 5–12) => E–L
        """
        if not filas:
            return fila_inicio

        start = fila_inicio
        end = fila_inicio + len(filas) - 1
        self.ws_estructura.Range(f"E{start}:L{end}").Value = filas
        return end + 1

    # -------------------------
    # US por RANGOS A–N (DEDUP por TipoDoc+Documento)
    # -------------------------
    def pegar_us_rango(self, filas, fila_inicio):
        nuevos = []

        for row in filas:
            if len(row) < 2:
                continue

            tipo = str(row[0]).strip()
            doc_original = row[1]
            doc = norm_doc(doc_original)

            if not tipo or not doc:
                continue

            key = f"{tipo}|{doc}"

            if key in self.seen_us:
                continue

            row[1] = doc
            fila_completa = row[:14] + [""] * (14 - len(row[:14]))

            nuevos.append(fila_completa)
            self.seen_us.add(key)

        if not nuevos:
            return fila_inicio

        start = fila_inicio
        end = fila_inicio + len(nuevos) - 1

        self.ws_us.Range(f"A{start}:N{end}").Value = nuevos

        # Persistencia para dedupe entre corridas
        self.append_us_control_batch([f"{r[0]}|{r[1]}" for r in nuevos])

        return end + 1

    # -------------------------
    # INDEX: US (para validar Activos)
    # -------------------------
    def cargar_us_keyset(self):
        """
        Devuelve set de "TIPO|DOC" leyendo US A,B desde la fila 2.
        """
        last = self.ultima_fila(self.ws_us, 2)  # col B
        if last < 2:
            return set()

        rng = self.ws_us.Range(f"A2:B{last}").Value
        out = set()

        if not rng:
            return out

        # rng puede ser tuple de tuples
        for row in rng:
            tipo = (str(row[0]).strip() if row[0] is not None else "")
            doc = norm_doc(row[1])
            if tipo and doc:
                out.add(f"{tipo}|{doc}")

        return out

    # -------------------------
    # INDEX: ESTRUCTURA base L/M por doc (primera ocurrencia)
    # -------------------------
    def cargar_estructura_base_lm(self):
        """
        Busca E (doc) y devuelve dict doc_norm -> {row, L, M}
        En ESTRUCTURA: E..M = 9 cols -> idx L=7, M=8
        """
        last = self.ultima_fila(self.ws_estructura, 5)  # col E
        if last < 2:
            return {}

        rng = self.ws_estructura.Range(f"E2:M{last}").Value
        out = {}

        if not rng:
            return out

        row_idx = 2
        for row in rng:
            doc = norm_doc(row[0])  # E
            if doc and doc not in out:
                l_val = row[7]  # L
                m_val = row[8]  # M
                out[doc] = {"row": row_idx, "L": l_val, "M": m_val}
            row_idx += 1

        return out

    # -------------------------
    # INDEX: Dedupe Activos en ESTRUCTURA (doc|codigo|fecha)
    # -------------------------
    def cargar_estructura_dedupe_activos(self):
        """
        Recorre E..H y F para construir key doc|codigo|fecha
        ESTRUCTURA: E=doc, F=fecha, H=codigo
        """
        last = self.ultima_fila(self.ws_estructura, 5)  # col E
        if last < 2:
            return set()

        rng = self.ws_estructura.Range(f"E2:H{last}").Value  # E F G H
        out = set()
        if not rng:
            return out

        for row in rng:
            doc = norm_doc(row[0])  # E
            fecha_key = _norm_fecha_key(row[1])  # F
            codigo = (str(row[3]).strip() if row[3] is not None else "")  # H
            if doc and fecha_key and codigo:
                out.add(f"{doc}|{codigo}|{fecha_key}")

        return out

    # -------------------------
    # ESCRITURA: Activos en ESTRUCTURA (D..M)
    # -------------------------
    def pegar_activos_estructura(self, plan_rows, fila_inicio):
        """
        Inserta nuevas filas para Activos escribiendo D..M (10 columnas).
        Mapeo:
          D = tipo_doc
          E = doc
          F = fecha
          H = codigo
          K = nombre homologado
          L = L_base
          M = M_base
        """
        if not plan_rows:
            return fila_inicio

        data = []
        for p in plan_rows:
            row = [""] * 10  # D..M
            row[0] = p.tipo_doc            # D
            row[1] = p.doc_norm            # E
            
          
            row[2] = p.fecha.strftime("%Y-%m-%d")
            
            row[4] = p.codigo              # H
            row[7] = p.nombre_homologado   # K
            row[8] = p.l_base              # L
            row[9] = p.m_base              # M
            data.append(row)

        start = fila_inicio
        end = fila_inicio + len(data) - 1

        self.ws_estructura.Range(f"D{start}:M{end}").Value = data
        return end + 1