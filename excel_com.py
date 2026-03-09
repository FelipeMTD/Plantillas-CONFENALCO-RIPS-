import win32com.client as win32
from pathlib import Path
import re
import math
from datetime import datetime, date, timedelta

XL_UP = -4162
CONTROL_SHEET = "__RIPS_CONTROL__"
_re_non_digits = re.compile(r"\D+")

def norm_doc(v):
    if v is None: return ""
    if isinstance(v, bool): return ""
    if isinstance(v, int): return str(v)
    if isinstance(v, float):
        if not math.isfinite(v): return ""
        if v.is_integer(): return str(int(v))
        return str(int(round(v)))
    s = str(v).strip()
    if not s: return ""
    m = re.fullmatch(r"(\d+)\.0+", s)
    if m: return m.group(1)
    try:
        f = float(s)
        if math.isfinite(f) and abs(f - round(f)) < 1e-6:
            return str(int(round(f)))
    except Exception: pass
    return _re_non_digits.sub("", s)

def _norm_fecha_key(v) -> str:
    if v is None or v == "": return ""
    if isinstance(v, datetime): return v.date().isoformat()
    if isinstance(v, date): return v.isoformat()
    if isinstance(v, (int, float)):
        base = datetime(1899, 12, 30)
        try:
            return (base + timedelta(days=float(v))).date().isoformat()
        except Exception: return str(v).strip()
    s = str(v).strip()
    if not s: return ""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError: pass
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

    def _init_control(self):
        try: self.ws_control = self.wb.Worksheets(CONTROL_SHEET)
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
            if not kind: break
            if kind == "U" and key: self.seen_us.add(str(key))
            row += 1

    def append_us_control_batch(self, docs):
        if not docs: return
        start = self.ws_control.Cells(self.ws_control.Rows.Count, 1).End(XL_UP).Row + 1
        data = [["U", d] for d in docs]
        end = start + len(data) - 1
        self.ws_control.Range(f"A{start}:B{end}").Value = data

    def siguiente_fila(self, ws, col):
        last_row = ws.Cells(ws.Rows.Count, col).End(XL_UP).Row
        return max(3, last_row + 1)

    def ultima_fila(self, ws, col):
        last = ws.Cells(ws.Rows.Count, col).End(XL_UP).Row
        return max(1, int(last))

    # ==========================================================
    # ARRASTRAR FÓRMULAS CON R1C1
    # ==========================================================
    def arrastrar_formulas(self, sheet_name, fila_ref, fila_inicio, fila_fin, col_max=50):
        if fila_inicio > fila_fin: return
        ws = self.wb.Sheets(sheet_name)
        self.excel.ScreenUpdating = False
        try:
            for col in range(1, col_max + 1):
                celda_modelo = ws.Cells(fila_ref, col)
                if celda_modelo.HasFormula:
                    formula_relativa = celda_modelo.FormulaR1C1
                    rango_destino = ws.Range(ws.Cells(fila_inicio, col), ws.Cells(fila_fin, col))
                    rango_destino.FormulaR1C1 = formula_relativa
        except Exception as e:
            print(f"    ⚠️ Error arrastrando fórmulas col {col}: {e}")
        finally:
            self.excel.ScreenUpdating = True

    # ==========================================================
    # BARRIDO DE FECHAS COLUMNA F (CON HORA Y APÓSTROFE)
    # ==========================================================
    def arreglar_formato_fechas_final(self, sheet_name, fila_inicio, fila_fin):
        if fila_inicio > fila_fin: return
        ws = self.wb.Sheets(sheet_name)
        self.excel.ScreenUpdating = False
        try:
            rango = ws.Range(f"F{fila_inicio}:F{fila_fin}")
            valores = rango.Value
            if not valores: return
            
            if not isinstance(valores, (list, tuple)):
                valores = [[valores]]
            
            nuevos = []
            for f in valores:
                val = f[0]
                if not val:
                    nuevos.append([""])
                    continue
                
                # Si es un objeto de fecha interno (ej: pywin32 datetime)
                if hasattr(val, "strftime"):
                    formateada = val.strftime("%Y-%m-%d %H:%M")
                else:
                    s = str(val).strip()
                    
                    # Separar fecha de hora por el espacio
                    partes = s.split(" ", 1)
                    fecha_str = partes[0]
                    
                    # Extraer hora si existe, garantizando formato HH:MM
                    if len(partes) > 1:
                        hora_str = partes[1].strip()
                        # Cortamos a 5 caracteres si trae segundos ej "11:17:00" -> "11:17"
                        if len(hora_str) >= 5 and ":" in hora_str:
                            hora_str = hora_str[:5] 
                    else:
                        # Si no hay hora, ponemos 00:00 por defecto
                        hora_str = "00:00"
                    
                    if len(fecha_str) == 8 and fecha_str.isdigit():
                        fecha_formateada = f"{fecha_str[:4]}-{fecha_str[4:6]}-{fecha_str[6:8]}"
                    else:
                        fecha_formateada = fecha_str
                        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%m/%d/%Y"):
                            try:
                                fecha_formateada = datetime.strptime(fecha_str, fmt).strftime("%Y-%m-%d")
                                break
                            except ValueError: pass
                    
                    formateada = f"{fecha_formateada} {hora_str}"
                
                nuevos.append([f"'{formateada}"])
            
            rango.NumberFormat = "@" 
            rango.Value = nuevos
        except Exception as e:
            print(f"    ⚠️ Error formateando fechas: {e}")
        finally:
            self.excel.ScreenUpdating = True

    def pegar_estructura_rango(self, filas, fila_inicio):
        if not filas: return fila_inicio
        start = fila_inicio
        end = fila_inicio + len(filas) - 1
        self.ws_estructura.Range(f"E{start}:L{end}").Value = filas
        return end + 1

    def pegar_us_rango(self, filas, fila_inicio):
        nuevos = []
        for row in filas:
            if len(row) < 2: continue
            tipo, doc_original = str(row[0]).strip(), row[1]
            doc = norm_doc(doc_original)
            if not tipo or not doc: continue
            key = f"{tipo}|{doc}"
            if key in self.seen_us: continue
            row[1] = doc
            fila_completa = row[:14] + [""] * (14 - len(row[:14]))
            nuevos.append(fila_completa)
            self.seen_us.add(key)

        if not nuevos: return fila_inicio
        start, end = fila_inicio, fila_inicio + len(nuevos) - 1
        self.ws_us.Range(f"A{start}:N{end}").Value = nuevos
        self.append_us_control_batch([f"{r[0]}|{r[1]}" for r in nuevos])
        return end + 1

    def cargar_us_keyset(self):
        last = self.ultima_fila(self.ws_us, 2)
        if last < 2: return set()
        rng = self.ws_us.Range(f"A2:B{last}").Value
        out = set()
        if not rng: return out
        if not isinstance(rng[0], (list, tuple)): rng = [rng]
        for row in rng:
            if not row: continue
            tipo, doc = (str(row[0]).strip() if row[0] else ""), norm_doc(row[1])
            if tipo and doc: out.add(f"{tipo}|{doc}")
        return out

    def cargar_estructura_base_lm(self):
        last = self.ultima_fila(self.ws_estructura, 5)
        if last < 2: return {}
        rng = self.ws_estructura.Range(f"E2:M{last}").Value
        out = {}
        if not rng: return out
        if not isinstance(rng[0], (list, tuple)): rng = [rng]
        row_idx = 2
        for row in rng:
            doc = norm_doc(row[0])
            if doc and doc not in out: out[doc] = {"row": row_idx, "L": row[7], "M": row[8]}
            row_idx += 1
        return out

    def cargar_estructura_dedupe_activos(self):
        last = self.ultima_fila(self.ws_estructura, 5)
        if last < 2: return set()
        rng = self.ws_estructura.Range(f"E2:H{last}").Value
        out = set()
        if not rng: return out
        if not isinstance(rng[0], (list, tuple)): rng = [rng]
        for row in rng:
            doc, fecha_key, codigo = norm_doc(row[0]), _norm_fecha_key(row[1]), (str(row[3]).strip() if row[3] else "")
            if doc and fecha_key and codigo: out.add(f"{doc}|{codigo}|{fecha_key}")
        return out

    def pegar_activos_estructura(self, plan_rows, fila_inicio):
        if not plan_rows: return fila_inicio
        data = []
        for p in plan_rows:
            # Enviamos fecha con 00:00 predeterminado; el barrido final pondrá el apóstrofe y dejará todo limpio
            data.append([p.tipo_doc, p.doc_norm, f"{p.fecha.strftime('%Y-%m-%d')} 00:00", "", p.codigo, "", "", p.nombre_homologado, p.l_base, p.m_base])
        start, end = fila_inicio, fila_inicio + len(data) - 1
        self.ws_estructura.Range(f"D{start}:M{end}").Value = data
        return end + 1