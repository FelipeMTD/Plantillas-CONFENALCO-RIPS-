import os
import shutil

def resetear_plantilla():
    # Definición de rutas
    ruta_maestra = r"C:\Users\FELIPE SISTEMAS\Documents\RIPS\RIPS COMFE_PLANTILLA.xlsm"
    ruta_destino = r"C:\Users\FELIPE SISTEMAS\Documents\RIPS\CODIGO\RIPS COMFE_PLANTILLA.xlsm"

    try:
        # 1. Borrar el archivo en el destino si existe
        if os.path.exists(ruta_destino):
            os.remove(ruta_destino)
            print(f"Archivo antiguo eliminado de: {ruta_destino}")
        
        # 2. Copiar la plantilla maestra al destino
        shutil.copy2(ruta_maestra, ruta_destino)
        print(f"Copia fresca creada exitosamente en: {ruta_destino}")

    except PermissionError:
        print("Error: El archivo está abierto. Ciérralo antes de ejecutar el script.")
    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")

if __name__ == "__main__":
    resetear_plantilla()