# main.py
from pathlib import Path
import zipfile, shutil, csv
from datetime import datetime
from texto_en_col import normalizar_carpeta_csv
from excel_com import ExcelCOM
from Activos.activos_proc import cargar_mapeo_activos, leer_activos_xlsx, construir_plan_activos, exportar_auditoria_csv, parse_fecha_usuario

BASE_DIR = Path(__file__).parent
ZIP_DIR, WORK_DIR, PLANTILLA = BASE_DIR / "zip", BASE_DIR / "_work", BASE_DIR / "RIPS_COMFE_PLANTILLA.xlsm"
ACTIVOS_DIR, ACTIVOS_JSON = BASE_DIR / "Activos", BASE_DIR / "Activos" / "Activos.json"

def iter_csv(path):
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f); next(r, None)
        yield from r

def extraer_zip(zip_path: Path) -> Path:
    destino = WORK_DIR / zip_path.stem
    if destino.exists(): shutil.rmtree(destino)
    destino.mkdir(parents=True, exist_ok=True); zipfile.ZipFile(zip_path).extractall(destino)
    return destino

def procesar_activos(excel: ExcelCOM):
    xlsx = next(ACTIVOS_DIR.glob("*.xlsx"), None)
    if not xlsx or not ACTIVOS_JSON.exists(): return
    print(f"\n=== ACTIVOS FIJOS ===\n[ACTIVOS] Excel: {xlsx.name}")
    while True:
        try:
            fecha = parse_fecha_usuario(input("FECHA DE CONSULTA PARA ACTIVOS (YYYY-MM-DD o DD/MM/YYYY): ").strip()); break
        except Exception as e: print(f"[ERROR] {e}")
    
    # LLAMADA CORREGIDA: leer_activos_xlsx ahora acepta sheet_name
    activos_rows = leer_activos_xlsx(xlsx, sheet_name="DETALLADO")
    plan, descartes = construir_plan_activos(excel, activos_rows, cargar_mapeo_activos(ACTIVOS_JSON), fecha)
    exportar_auditoria_csv(BASE_DIR / f"auditoria_activos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", plan, descartes)
    
    if plan and input("¿Aplicar inserción en ESTRUCTURA? (SI para continuar): ").strip().upper() == "SI":
        fila_inicio = excel.siguiente_fila(excel.ws_estructura, 5)
        excel.pegar_activos_estructura(plan, fila_inicio)

def main():
    print("=== RIPS COM – MÁXIMA VELOCIDAD ===")
    zips = sorted(ZIP_DIR.glob("*.zip"))
    if not zips: return
    excel = ExcelCOM(PLANTILLA)
    try:
        excel.abrir()
        fila_estructura = excel.siguiente_fila(excel.ws_estructura, 5) # Empezará en fila 3
        fila_us = excel.siguiente_fila(excel.ws_us, 2) # Empezará en fila 3
        for zip_file in zips:
            carpeta = extraer_zip(zip_file); normalizar_carpeta_csv(carpeta); filas_est = []
            for tipo, mapa in {"AT": {3: 0, 4: 1, 7: 6, 11: 7}, "AP": {3: 0, 4: 1, 10: 4, 15: 6, 16: 7}, "AC": {3: 0, 4: 1, 9: 4, 17: 6, 18: 7}}.items():
                path = next(carpeta.glob(f"{tipo}*.CSV"), None)
                if path:
                    for r in iter_csv(path):
                        f = [""] * 8
                        for ic, il in mapa.items():
                            if ic < len(r): f[il] = r[ic]
                        filas_est.append(f)
            fila_estructura = excel.pegar_estructura_rango(filas_est, fila_estructura)
            filas_us = []
            path_us = next(carpeta.glob("US*.CSV"), None)
            if path_us:
                for r in iter_csv(path_us): filas_us.append((r + [""] * 14)[:14])
            fila_us = excel.pegar_us_rango(filas_us, fila_us)
        procesar_activos(excel)
    finally: excel.cerrar()

if __name__ == "__main__": main()