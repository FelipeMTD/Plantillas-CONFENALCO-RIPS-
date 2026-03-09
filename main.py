import sys
from pathlib import Path
import zipfile
import shutil
import csv
from datetime import datetime

# Importamos módulos propios
from texto_en_col import normalizar_carpeta_csv
from excel_com import ExcelCOM
from Activos.activos_proc import (
    cargar_mapeo_activos, 
    leer_activos_xlsx, 
    construir_plan_activos, 
    exportar_auditoria_csv, 
    parse_fecha_usuario
)

BASE_DIR = Path(__file__).parent
ZIP_DIR = BASE_DIR / "zip"
WORK_DIR = BASE_DIR / "_work"
PLANTILLA = BASE_DIR / "RIPS_COMFE_PLANTILLA.xlsm"

ACTIVOS_DIR = BASE_DIR / "Activos"
ACTIVOS_JSON = ACTIVOS_DIR / "Activos.json"

def iter_csv(path):
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f)
        try: next(r, None) 
        except StopIteration: pass
        yield from r

def extraer_zip(zip_path: Path) -> Path:
    destino = WORK_DIR / zip_path.stem
    if destino.exists(): shutil.rmtree(destino)
    destino.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path) as z: z.extractall(destino)
    return destino

def procesar_activos(excel: ExcelCOM):
    print("\n" + "="*50)
    print("🏥  MÓDULO DE ACTIVOS FIJOS")
    print("="*50)

    xlsx = next(ACTIVOS_DIR.glob("*.xlsx"), None)
    if not xlsx:
        print("⚠️  No se encontró archivo Excel (.xlsx) en la carpeta 'Activos'. Saltando...")
        return
    
    if not ACTIVOS_JSON.exists():
        print("⚠️  No se encontró 'Activos.json'. Saltando...")
        return

    print(f"📄  Archivo encontrado: {xlsx.name}")
    
    while True:
        try:
            fecha_str = input("📅  INGRESE FECHA DE CONSULTA (YYYY-MM-DD): ").strip()
            fecha = parse_fecha_usuario(fecha_str)
            break
        except Exception as e:
            print(f"❌  Error: {e}. Intente de nuevo.")

    print(f"\n⏳  Leyendo archivo de activos...")
    activos_rows = leer_activos_xlsx(xlsx, sheet_name="DETALLADO")
    
    print(f"🔍  Cruzando {len(activos_rows)} registros con la base de datos...")
    plan, descartes = construir_plan_activos(excel, activos_rows, cargar_mapeo_activos(ACTIVOS_JSON), fecha)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    csv_audit = BASE_DIR / f"auditoria_activos_{timestamp}.csv"
    exportar_auditoria_csv(csv_audit, plan, descartes)
    print(f"📊  Auditoría guardada en: {csv_audit.name}")
    print(f"    - Insertables: {len(plan)}")
    print(f"    - Descartados: {len(descartes)}")

    if plan:
        resp = input("\n✍️  ¿Desea insertar estos activos en la hoja ESTRUCTURA? (SI/NO): ").strip().upper()
        if resp == "SI":
            print("⏳  Insertando en Excel...")
            fila_inicio = excel.siguiente_fila(excel.ws_estructura, 5)
            excel.pegar_activos_estructura(plan, fila_inicio)
            print("✅  Inserción de activos completada (Se formateará al final).")
        else:
            print("info  Operación cancelada.")
    else:
        print("info  No hay registros válidos para insertar.")

def main():
    print("\n" + "="*60)
    print("🚀  RIPS COM – AUTOMATIZACIÓN DE ESTRUCTURA Y USUARIOS")
    print("="*60)

    zips = sorted(ZIP_DIR.glob("*.zip"))
    if not zips:
        print("❌  No se encontraron archivos .zip en la carpeta 'zip'.")
        return

    print(f"📦  Archivos ZIP encontrados: {len(zips)}")
    print("⏳  Abriendo Excel (modo oculto)...")
    excel = ExcelCOM(PLANTILLA)
    
    try:
        excel.abrir()
        print("✅  Excel abierto correctamente.")

        fila_estructura = excel.siguiente_fila(excel.ws_estructura, 5)
        fila_us = excel.siguiente_fila(excel.ws_us, 2)
        print(f"📍  Punto de partida -> Estructura: Fila {fila_estructura} | US: Fila {fila_us}")

        # ========================================================
        # 1. PEGADO MASIVO DE RIPS (ZIPS)
        # ========================================================
        for i, zip_file in enumerate(zips, 1):
            print(f"\n[{i}/{len(zips)}] 📂 Procesando ZIP: {zip_file.name}")
            carpeta = extraer_zip(zip_file)
            print("    🛠️  Normalizando CSVs...")
            normalizar_carpeta_csv(carpeta)

            filas_est = []
            mapas = {
                "AT": {3: 0, 4: 1, 7: 6, 11: 7},
                "AP": {3: 0, 4: 1, 10: 4, 15: 6, 16: 7},
                "AC": {3: 0, 4: 1, 9: 4, 17: 6, 18: 7}
            }

            for tipo, mapa in mapas.items():
                path = next(carpeta.glob(f"{tipo}*.CSV"), None)
                if path:
                    for r in iter_csv(path):
                        row_data = [""] * 8
                        for idx_csv, idx_list in mapa.items():
                            if idx_csv < len(r):
                                row_data[idx_list] = r[idx_csv]
                        filas_est.append(row_data)

            if filas_est:
                print(f"    💾  Pegando {len(filas_est)} filas en ESTRUCTURA...")
                fila_estructura = excel.pegar_estructura_rango(filas_est, fila_estructura)
            else:
                print("    ⚠️  No hay datos de estructura en este ZIP.")

            filas_us = []
            path_us = next(carpeta.glob("US*.CSV"), None)
            if path_us:
                for r in iter_csv(path_us):
                    filas_us.append((r + [""] * 14)[:14])
                
                if filas_us:
                    print(f"    👥  Procesando {len(filas_us)} usuarios...")
                    fila_us = excel.pegar_us_rango(filas_us, fila_us)
            else:
                print("    ⚠️  No hay archivo US.")

        # ========================================================
        # 2. PEGADO MASIVO DE ACTIVOS FIJOS
        # ========================================================
        procesar_activos(excel)
        
        # ========================================================
        # 3. AJUSTES FINALES (FORMATO FECHA Y FÓRMULAS)
        # ========================================================
        ultima_fila_datos = excel.ultima_fila(excel.ws_estructura, 5) # Columna E determina el fin de los datos
        
        if ultima_fila_datos >= 3:
            print("\n" + "="*50)
            print("⚙️  APLICANDO AJUSTES FINALES A ESTRUCTURA")
            print("="*50)
            print(f"   📅  Aplicando formato de fecha final (AAAA-MM-DD)...")
            excel.arreglar_formato_fechas_final("ESTRUCTURA", 3, ultima_fila_datos)
            
            print(f"   🪄  Arrastrando fórmulas de la fila 2 hasta la {ultima_fila_datos}...")
            excel.arrastrar_formulas("ESTRUCTURA", 2, 3, ultima_fila_datos)
            print("   ✅  Ajustes finalizados con éxito.")

    except Exception as e:
        print(f"\n❌  ERROR CRÍTICO: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("\n⏳  Cerrando Excel...")
        excel.cerrar()
        print("✨  ¡Proceso Finalizado!")

if __name__ == "__main__":
    main()