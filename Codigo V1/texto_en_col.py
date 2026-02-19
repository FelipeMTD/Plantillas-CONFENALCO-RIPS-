from pathlib import Path
import csv

def normalizar_csv(csv_file: Path):
    with csv_file.open("r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    with csv_file.open("w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(rows)

def normalizar_carpeta_csv(carpeta):
    carpeta = Path(carpeta)
    for csv_file in carpeta.glob("*.CSV"):
        normalizar_csv(csv_file)
        print(f"[OK] {csv_file.name}")

def main():
    import sys
    normalizar_carpeta_csv(sys.argv[1])

if __name__ == "__main__":
    main()
