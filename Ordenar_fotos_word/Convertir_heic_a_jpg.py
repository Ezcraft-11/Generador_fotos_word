
import os
from PIL import Image
from pillow_heif import register_heif_opener

# Registrar el decodificador de HEIC
register_heif_opener()

def convertir_heic_a_jpg(carpeta_origen, carpeta_destino="Fotos_JPG"):
    if not os.path.exists(carpeta_origen):
        print(f"La carpeta '{carpeta_origen}' no existe.")
        return

    os.makedirs(carpeta_destino, exist_ok=True)
    archivos = [f for f in os.listdir(carpeta_origen) if f.lower().endswith('.heic')]

    if not archivos:
        print("No se encontraron imágenes .heic.")
        return

    for archivo in archivos:
        ruta_heic = os.path.join(carpeta_origen, archivo)
        nombre_base = os.path.splitext(archivo)[0]
        ruta_jpg = os.path.join(carpeta_destino, f"{nombre_base}.jpg")

        imagen = Image.open(ruta_heic)
        imagen.convert("RGB").save(ruta_jpg, "JPEG")
        print(f"Convertida con éxito: {archivo} -> {nombre_base}.jpg")

if __name__ == "__main__":
    # Cambia la ruta por la carpeta donde están tus fotos HEIC
    convertir_heic_a_jpg("./mis_fotos")

