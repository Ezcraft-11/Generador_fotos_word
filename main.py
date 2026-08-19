
import os
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

def crear_word_por_cada_grupo(carpeta_origen, carpeta_destino="Archivos_Word"):
    extensiones_validas = ('.png', '.jpg', '.jpeg', '.webp', '.bmp')
    
    if not os.path.exists(carpeta_origen):
        print(f"Error: La carpeta '{carpeta_origen}' no existe.")
        return

    archivos = [f for f in os.listdir(carpeta_origen) if f.lower().endswith(extensiones_validas)]
    archivos.sort()

    if not archivos:
        print("No se encontraron imágenes en la carpeta.")
        return

    # Crear carpeta para guardar los distintos archivos Word
    os.makedirs(carpeta_destino, exist_ok=True)

    # Procesar en bloques de 4
    for num_archivo, i in enumerate(range(0, len(archivos), 4), start=1):
        grupo = archivos[i:i+4]
        doc = Document()
        
        # Configurar hoja en horizontal
        section = doc.sections[0]
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width

        # Márgenes reducidos a 0.3 pulgadas para maximizar el espacio
        section.top_margin = Inches(0.3)
        section.bottom_margin = Inches(0.3)
        section.left_margin = Inches(0.3)
        section.right_margin = Inches(0.3)

        # Crear cuadrícula 2x2
        tabla = doc.add_table(rows=2, cols=2)
        tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
        tabla.autofit = False

        posiciones = [(0, 0), (0, 1), (1, 0), (1, 1)]

        for idx, nombre_foto in enumerate(grupo):
            row_idx, col_idx = posiciones[idx]
            celda = tabla.cell(row_idx, col_idx)
            
            p = celda.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            
            ruta_foto = os.path.join(carpeta_origen, nombre_foto)
            
            # Limitar por ALTURA (2.7 pulgadas) asegura que la fila inferior no sea empujada fuera de la página
            run = p.add_run()
            run.add_picture(ruta_foto, height=Inches(2.7))

        # Guardar cada grupo de 4 imágenes en su propio archivo Word
        nombre_salida = os.path.join(carpeta_destino, f"Grupo_{num_archivo}.docx")
        doc.save(nombre_salida)
        print(f"Archivo generado: {nombre_salida}")

if __name__ == "__main__":
    carpeta_fotos = "./mis_fotos"
    crear_word_por_cada_grupo(carpeta_fotos)
