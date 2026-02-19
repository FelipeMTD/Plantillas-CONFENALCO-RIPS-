from __future__ import annotations

from pathlib import Path
import zipfile
import shutil

from texto_en_col import normalizar_carpeta_csv
from cargar_rips_a_estructura import abrir_plantilla, cargar_en_hojas, PLANTILLA


BASE_DIR = Path(__file__).parent
ZIP_DIR = BASE_DIR / "zip"
WORK_DIR = BASE_DIR / "_work"


def extraer_zip(zip_path: Path) -> Path:
    destino = WORK_DIR / zip_path.stem
    if destino.exists():
        shutil.rmtree(destino)
    destino.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path) as z:
        z.extractall(destino)
    return destino


def main():
    print("=== INICIO PROCESO RIPS ===")

    if not ZIP_DIR.exists():
        raise FileNotFoundError(f"No existe la carpeta ZIP: {ZIP_DIR}")

    WORK_DIR.mkdir(parents=True, exist_ok=True)

    zips = sorted(ZIP_DIR.glob("*.zip"))
    print(f"[INFO] ZIPs detectados: {len(zips)}")
    if not zips:
        print("[INFO] No hay ZIPs para procesar. Fin.")
        return

    # OPTIMIZACIÓN: abrir una sola vez
    print(f"[EXCEL] Abriendo plantilla una sola vez: {PLANTILLA.name}")
    wb, ws, ws_us, ws_control = abrir_plantilla()

    fila_estructura = None
    fila_us = None

    for idx, zip_file in enumerate(zips, start=1):
        print(f"\n--- ({idx}/{len(zips)}) PROCESANDO ZIP: {zip_file.name} ---")

        carpeta = extraer_zip(zip_file)
        print(f"[OK] ZIP extraído en {carpeta}")

        print("[STEP] Normalizando CSVs (texto_en_col.py)")
        normalizar_carpeta_csv(carpeta)

        print("[STEP] Pegando en plantilla (US dedupe activo, ESTRUCTURA sin dedupe)")
        fila_estructura, fila_us = cargar_en_hojas(
            carpeta_csv=carpeta,
            ws_estructura=ws,
            ws_us=ws_us,
            ws_control=ws_control,
            fila_estructura=fila_estructura,
            fila_us=fila_us,
            imprimir_cada=1000,
        )

        print(f"[OK] Punteros -> ESTRUCTURA fila={fila_estructura} | US fila={fila_us}")

    print("\n[EXCEL] Guardando UNA sola vez...")
    wb.save(PLANTILLA)
    print("[EXCEL] Guardado completo.")

    print("\n=== PROCESO COMPLETO ===")
    print(f"[RESUMEN] Siguiente fila ESTRUCTURA: {fila_estructura}")
    print(f"[RESUMEN] Siguiente fila US: {fila_us}")


if __name__ == "__main__":
    main()
