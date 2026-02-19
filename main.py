# main.py
from pathlib import Path
import zipfile
import shutil
import csv
from datetime import datetime

from texto_en_col import normalizar_carpeta_csv
from excel_com import ExcelCOM

from Activos.activos_proc import (
    cargar_mapeo_activos,
    leer_activos_xlsx,
    construir_plan_activos,
    exportar_auditoria_csv,
    parse_fecha_usuario,
)

BASE_DIR = Path(__file__).parent
ZIP_DIR = BASE_DIR / "zip"
WORK_DIR = BASE_DIR / "_work"
PLANTILLA = BASE_DIR / "RIPS COMFE_PLANTILLA.xlsm"

ACTIVOS_DIR = BASE_DIR / "Activos"
ACTIVOS_JSON = ACTIVOS_DIR / "Activos.json"
# toma el primer xlsx de Activos/ para no amarrarte a un nombre
def _buscar_activos_xlsx():
    if ACTIVOS_DIR.exists():
        xs = sorted(ACTIVOS_DIR.glob("*.xlsx"))
        if xs:
            return xs[0]
    return None


def iter_csv(path):
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        r = csv.reader(f)
        next(r, None)
        for row in r:
            yield row


def extraer_zip(zip_path: Path) -> Path:
    destino = WORK_DIR / zip_path.stem
    if destino.exists():
        shutil.rmtree(destino)
    destino.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path) as z:
        z.extractall(destino)
    return destino


def procesar_activos(excel: ExcelCOM):
    xlsx = _buscar_activos_xlsx()
    if not xlsx:
        print("[ACTIVOS] No encontré .xlsx en carpeta Activos/. Se omite.")
        return

    if not ACTIVOS_JSON.exists():
        print("[ACTIVOS] Falta Activos/Activos.json. Se omite.")
        return

    print(f"\n=== ACTIVOS FIJOS ===")
    print(f"[ACTIVOS] Excel: {xlsx.name}")
    print(f"[ACTIVOS] JSON:  {ACTIVOS_JSON.name}")

    # Fecha usuario
    while True:
        s = input("FECHA DE CONSULTA PARA ACTIVOS FIJO (YYYY-MM-DD o DD/MM/YYYY): ").strip()
        try:
            fecha = parse_fecha_usuario(s)
            break
        except Exception as e:
            print(f"[ERROR] {e}")

    # Carga activos + mapeo
    mapeo = cargar_mapeo_activos(ACTIVOS_JSON)
    activos_rows = leer_activos_xlsx(xlsx, sheet_name="DETALLADO")

    # Dry run (validación lógica)
    plan, descartes = construir_plan_activos(excel, activos_rows, mapeo, fecha)

    # Resumen
    print(f"[ACTIVOS] Total filas leídas: {len(activos_rows)}")
    print(f"[ACTIVOS] Aprobadas para insertar: {len(plan)}")
    print(f"[ACTIVOS] Descartadas: {len(descartes)}")

    # Top razones descarte
    conteo = {}
    for d in descartes:
        conteo[d["reason"]] = conteo.get(d["reason"], 0) + 1
    if conteo:
        print("[ACTIVOS] Descartes por razón:")
        for k in sorted(conteo, key=conteo.get, reverse=True):
            print(f"  - {k}: {conteo[k]}")

    # Auditoría CSV
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    audit_path = BASE_DIR / f"auditoria_activos_{ts}.csv"
    exportar_auditoria_csv(audit_path, plan, descartes)
    print(f"[ACTIVOS] Auditoría exportada: {audit_path.name}")

    if not plan:
        print("[ACTIVOS] No hay nada que insertar.")
        return

    # Confirmación explícita (porque escribir en ESTRUCTURA es destructivo)
    ok = input("¿Aplicar inserción en ESTRUCTURA? Escribe SI para continuar: ").strip().upper()
    if ok != "SI":
        print("[ACTIVOS] Cancelado por usuario. No se escribió nada.")
        return

    # Escribir (append al final de ESTRUCTURA)
    fila_inicio = excel.siguiente_fila(excel.ws_estructura, 5)  # col E
    fila_fin = excel.pegar_activos_estructura(plan, fila_inicio)
    print(f"[ACTIVOS] OK. Filas insertadas: {len(plan)}. Nueva fila siguiente: {fila_fin}")


def main():
    print("=== RIPS COM – MÁXIMA VELOCIDAD ===")

    zips = sorted(ZIP_DIR.glob("*.zip"))
    if not zips:
        print("No hay ZIPs")
        return

    excel = ExcelCOM(PLANTILLA)

    try:
        excel.abrir()

        fila_estructura = excel.siguiente_fila(excel.ws_estructura, 5)
        fila_us = excel.siguiente_fila(excel.ws_us, 2)

        for i, zip_file in enumerate(zips, 1):
            print(f"\n--- ({i}/{len(zips)}) {zip_file.name} ---")
            carpeta = extraer_zip(zip_file)

            normalizar_carpeta_csv(carpeta)

            filas_estructura = []

            for tipo, mapa in {
                "AT": {3: 0, 4: 1, 7: 6, 11: 7},
                "AP": {3: 0, 4: 1, 10: 4, 15: 6, 16: 7},
                "AC": {3: 0, 4: 1, 9: 4, 17: 6, 18: 7},
            }.items():
                csv_path = next(carpeta.glob(f"{tipo}*.CSV"), None)
                if not csv_path:
                    continue

                for row in iter_csv(csv_path):
                    fila = [""] * 8  # E–L
                    for idx_csv, idx_local in mapa.items():
                        if idx_csv < len(row):
                            fila[idx_local] = row[idx_csv]
                    filas_estructura.append(fila)

            fila_estructura = excel.pegar_estructura_rango(
                filas_estructura, fila_estructura
            )

            filas_us = []
            csv_us = next(carpeta.glob("US*.CSV"), None)

            if csv_us:
                for row in iter_csv(csv_us):
                    fila = [""] * 14
                    for j in range(min(14, len(row))):
                        fila[j] = row[j]
                    filas_us.append(fila)

            fila_us = excel.pegar_us_rango(filas_us, fila_us)

            print(f"[OK] filas -> ESTRUCTURA={fila_estructura} | US={fila_us}")

        # ACTIVOS: se ejecuta después de cargar ZIPs (como pediste)
        procesar_activos(excel)

    finally:
        excel.cerrar()
        print("\n[EXCEL] cerrado correctamente")

    print("\n=== FIN ===")


if __name__ == "__main__":
    main()
