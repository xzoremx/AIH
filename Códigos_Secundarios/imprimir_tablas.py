import os
import sys

# Determinar ruta base según si corre con PyInstaller o no
if hasattr(sys, '_MEIPASS'):
    # Si está empaquetado con PyInstaller
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

# Subir un nivel y buscar ms-playwright
ms_playwright_path= os.path.join(base_dir, "ms-playwright")
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = ms_playwright_path

from playwright.sync_api import sync_playwright

# Configuración de directorios y servidor

directorio = os.path.join(base_dir, "PDFs_Iniciales")
servidor = "http://localhost:5000"
salida_pdf = os.path.join(base_dir, "PDFs_Impresos")


os.makedirs(salida_pdf, exist_ok=True)

archivos = sorted([f for f in os.listdir(directorio) if f.endswith('.csv')])
tablas = [f[:-4] for f in archivos]

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    for tabla in tablas:
        url = f"{servidor}/?tabla={tabla}"
        print(f"Generando PDF para: {tabla}")
        page.goto(url)
        page.wait_for_selector("table")
        # Exporta solo la tabla, en Horizontal (A4), escala 50%, márgenes 0
        pdf_path = os.path.join(salida_pdf, f"{tabla}.pdf")
        page.pdf(
            path=pdf_path,
            format="A4",
            landscape=True,
            margin={"top": "0mm", "bottom": "0mm", "left": "0mm", "right": "0mm"},
            scale=0.5,  # Equivalente a 50% de la vista del navegador
            print_background=True
        )
        print(f"Guardado: {pdf_path}")
    browser.close()

print("¡Todos los PDFs han sido generados automáticamente!")
