import win32com.client as win32
from pathlib import Path
import re

XL_UP = -4162
CONTROL_SHEET = "__RIPS_CONTROL__"
_re_non_digits = re.compile(r"\D+")


def norm_doc(v):
    if v is None:
        return ""
    return _re_non_digits.sub("", str(v))


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

    # -------------------------
    # ESTRUCTURA por RANGOS
    # -------------------------
    def pegar_estructura_rango(self, filas, fila_inicio):
        """
        filas: lista de listas de longitud fija (columnas 5–12)
        """
        if not filas:
            return fila_inicio

        start = fila_inicio
        end = fila_inicio + len(filas) - 1
        # E=5 … L=12
        self.ws_estructura.Range(f"E{start}:L{end}").Value = filas
        return end + 1

    # -------------------------
    # US por RANGOS
    # -------------------------
    # -------------------------
    # US por RANGOS A–N (DEDUP por TipoDoc+Documento)
    # -------------------------
    def pegar_us_rango(self, filas, fila_inicio):
        """
        filas: lista de listas con 14 columnas (A–N)
        Dedupe por TipoDoc + Documento
        """

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

            # Normalizamos solo documento
            row[1] = doc

            # Aseguramos 14 columnas
            fila_completa = row[:14] + [""] * (14 - len(row[:14]))

            nuevos.append(fila_completa)
            self.seen_us.add(key)

        if not nuevos:
            return fila_inicio

        start = fila_inicio
        end = fila_inicio + len(nuevos) - 1

        self.ws_us.Range(f"A{start}:N{end}").Value = nuevos

        # Guardar control
        control_data = [["U", f"{r[0]}|{r[1]}"] for r in nuevos]
        self.append_us_control_batch([f"{r[0]}|{r[1]}" for r in nuevos])

        return end + 1
