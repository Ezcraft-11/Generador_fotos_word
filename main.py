
import os
from docx import Document
from docx.shared import Inches

def insertar_imagenes_en_word(carpeta_origen, archivo_salida="Imagenes_Organizadas.docx"):
    # Crear un documento nuevo
    doc = Document()
    
    # Reducir un poco los márgenes superior e inferior para asegurar que quepan las 4 imágenes
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)

    # Extensiones de imagen soportadas
    extensiones_validas = ('.png', '.jpg', '.jpeg', '.webp', '.bmp')
    
    # Filtrar y ordenar archivos de la carpeta
    archivos = [f for f in os.listdir(carpeta_origen) if f.lower().endswith(extensiones_validas)]
    archivos.sort()

    if not archivos:
        print("No se encontraron imágenes en la carpeta especificada.")
        return

    # Iterar sobre las imágenes en bloques de 4
    for i in range(0, len(archivos), 4):
        grupo = archivos[i:i+4]
        
        for nombre_foto in grupo:
            ruta_foto = os.path.join(carpeta_origen, nombre_foto)
            
            # Ajustar la altura uniforme (2.1 pulgadas permite que 4 quepan cómodamente en vertical)
            doc.add_picture(ruta_foto, height=Inches(2.1))
            
            # Reducir el espacio en blanco sobrante debajo de cada imagen
            ultimo_parrafo = doc.paragraphs[-1]
            ultimo_parrafo.paragraph_format.space_after = Inches(0.1)

        # Si quedan más imágenes pendientes, añadir un salto de página
        if i + 4 < len(archivos):
            doc.add_page_break()

    # Guardar el resultado final
    doc.save(archivo_salida)
    print(f"Proceso completado. Archivo guardado como '{archivo_salida}'.")

# --- USO DEL SCRIPT ---
# Cambia la ruta por la carpeta donde guardas tus fotos
carpeta_de_fotos = "./mis_fotos"
insertar_imagenes_en_word(carpeta_de_fotos)
