import os
import re
import sys
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import messagebox, filedialog
from PIL import Image
import io

if hasattr(sys, '_MEIPASS'):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

from PIL import Image, ImageEnhance

# --- Procesar marca de agua ---
img_path = os.path.join(base_dir, "Formatos_Página", "marca_borrador.png")
imagen = Image.open(img_path).convert("RGBA")

pixels = imagen.getdata()
nuevo_color = (180, 180, 180)  # gris suave
factor_transparencia = 0.3

nuevos_pixels = []
for r, g, b, a in pixels:
    if a > 0:
        nuevos_pixels.append((*nuevo_color, int(a * factor_transparencia)))
    else:
        nuevos_pixels.append((r, g, b, a))
imagen.putdata(nuevos_pixels)

img_byte_arr = io.BytesIO()
imagen.save(img_byte_arr, format='PNG')
marca_agua_img = img_byte_arr.getvalue()

sys.stdout.reconfigure(encoding='utf-8')

# Preguntar si aplicar marca de agua
def preguntar_borrador():
    root = tk.Tk()
    root.withdraw()
    respuesta = messagebox.askyesno("Marca de Agua", "¿Deseas aplicar la marca de agua 'BORRADOR'?")
    root.destroy()
    return respuesta

aplicar_borrador = preguntar_borrador()

# --- Función para orden natural ---
def orden_natural(texto):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', texto)]

# Carpeta de PDFs a unir
carpeta_pdfs = os.path.join(base_dir, "PDFs_Finales")
pdfs_para_unir = sorted(
    [os.path.join(carpeta_pdfs, f) for f in os.listdir(carpeta_pdfs) if f.lower().endswith('.pdf')],
    key=lambda x: orden_natural(os.path.basename(x))
)

# --- Definir rutas de salida ---
# 1. Ruta default (un nivel superior a base_dir)
output_dir_default = os.path.abspath(os.path.join(base_dir, ".."))
output_path_default = os.path.join(output_dir_default, "PDF_Unido.pdf")

# 2. Ruta manual (siempre preguntar al usuario)
root = tk.Tk()
root.withdraw()
output_path_manual = filedialog.asksaveasfilename(
    title="Guardar PDF unido como...",
    defaultextension=".pdf",
    filetypes=[("Archivos PDF", "*.pdf")],
    initialfile="PDF_Unido.pdf"
)
root.destroy()

if not output_path_manual:  # Si el usuario cancela
    print("Operación cancelada en la ruta manual. Solo se guardará en la ruta default.")

# --- Unir PDFs ---
pdf_final = fitz.open()
for pdf_file in pdfs_para_unir:
    doc = fitz.open(pdf_file)
    pdf_final.insert_pdf(doc)

total_paginas = pdf_final.page_count

# --- Numerar y aplicar marca de agua ---
for i, page in enumerate(pdf_final, start=1):
    bloques = page.search_for("INFORME DE ENSAYO N°")
    if bloques:
        bbox = bloques[0]
        y_encabezado = bbox.y0 + (bbox.y1 - bbox.y0) / 2
        texto = f"Pág. {i}/{total_paginas}"
        page.insert_text(
            (page.rect.width - 100, y_encabezado),
            texto,
            fontname="helv",
            fontsize=8,
            color=(0, 0, 0)
        )
    else:
        page.insert_text(
            (page.rect.width - 55, 80),
            f"Pág. {i}/{total_paginas}",
            fontname="helv",
            fontsize=8,
            color=(0, 0, 0)
        )

    if aplicar_borrador:
        page_width = page.rect.width
        page_height = page.rect.height
        lado_base = min(page_width, page_height)
        escala = 0.6
        ancho_img = lado_base * escala
        alto_img = lado_base * escala
        cx, cy = page_width / 2, page_height / 2
        rect_img = fitz.Rect(cx - ancho_img / 2, cy - alto_img / 2, cx + ancho_img / 2, cy + alto_img / 2)
        page.insert_image(rect_img, stream=marca_agua_img, overlay=True)

# Guardar en ambas rutas
pdf_final.save(output_path_default)
print(f"PDF guardado en ruta default: {output_path_default}")

if output_path_manual:
    pdf_final.save(output_path_manual)
    print(f"PDF guardado también en ruta manual: {output_path_manual}")

