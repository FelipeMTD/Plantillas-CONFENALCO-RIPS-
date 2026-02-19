from pathlib import Path
import zipfile
import shutil
import csv

from texto_en_col import normalizar_carpeta_csv
from excel_com import ExcelCOM


BASE_DIR = Path(__file__).parent
ZIP_DIR = BASE_DIR / "zip"
WORK_DIR = BASE_DIR / "_work"
PLANTILLA = BASE_DIR / "RIPS COMFE_PLANTILLA.xlsm"


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
                    for i in range(min(14, len(row))):
                        fila[i] = row[i]
                    filas_us.append(fila)

            fila_us = excel.pegar_us_rango(filas_us, fila_us)


            print(f"[OK] filas -> ESTRUCTURA={fila_estructura} | US={fila_us}")

    finally:
        excel.cerrar()
        print("\n[EXCEL] cerrado correctamente")

    print("\n=== FIN ===")


if __name__ == "__main__":
    main()
