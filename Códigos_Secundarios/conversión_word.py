import os
import comtypes.client
import sys

if hasattr(sys, '_MEIPASS'):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

sys.stdout.reconfigure(encoding='utf-8')

# Constantes de Word (mejoran la legibilidad)
wdExportFormatPDF = 17
wdExportOptimizeForPrint = 0
wdExportOptimizeForOnScreen = 1
wdExportAllDocument = 0
wdExportSelection = 1
wdExportCurrentPage = 2
wdExportFromTo = 3
wdExportDocumentContent = 0
wdExportCreateNoBookmarks = 0
wdStatisticPages = 2

def convertir_word_a_pdfs_paginas_ordenadas(carpeta_word, carpeta_destino):
    docx_files = [f for f in os.listdir(carpeta_word) if f.lower().endswith('.docx')]
    if not docx_files:
        print(" No se encontró ningún archivo .docx en la carpeta.")
        return

    docx_path_absolute = os.path.abspath(os.path.join(carpeta_word, docx_files[0]))
    print(f"Usando Word encontrado en carpeta: {docx_path_absolute}")

    carpeta_destino_absolute = os.path.abspath(carpeta_destino)

    if not os.path.exists(carpeta_destino_absolute):
        os.makedirs(carpeta_destino_absolute)
        print(f" Carpeta de destino creada: {carpeta_destino_absolute}")

    word = None
    doc = None
    try:
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False  # Ejecutar en segundo plano

        # Abrir el documento Word
        doc = word.Documents.Open(docx_path_absolute)

        # Forzar actualización de campos (ej. número de páginas)
        doc.Fields.Update()

        # Obtener el número total de páginas
        total_paginas = doc.ComputeStatistics(wdStatisticPages)
        print(f" Número total de páginas detectadas: {total_paginas}")

        if total_paginas == 0:
            print(" El documento parece estar vacío o no se pudo calcular el número de páginas.")
            return

        for i in range(1, total_paginas + 1):
            if i == total_paginas:
                if total_paginas == 1:
                    nombre_pdf = "00AAA_Unica_Pagina.pdf"
                else:
                    nombre_pdf = "ZZZ_Ultima_Pagina.pdf"
            else:
                nombre_pdf = f"00AAA_Pag_{i:02}.pdf"

            ruta_salida_absolute = os.path.join(carpeta_destino_absolute, nombre_pdf)

            print(f" Exportando página {i} a {nombre_pdf}...")

            # Exportar como PDF página por página
            doc.ExportAsFixedFormat(
                OutputFileName=ruta_salida_absolute,
                ExportFormat=wdExportFormatPDF,
                OpenAfterExport=False,
                OptimizeFor=wdExportOptimizeForPrint,
                Range=wdExportFromTo,
                From=i,
                To=i,
                Item=wdExportDocumentContent,
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=wdExportCreateNoBookmarks,
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False
            )

            print(f" Página {i} exportada como {nombre_pdf}")

        print(" Conversión completa.")

    except Exception as e:
        print(f" Ocurrió un error durante la conversión: {e}")
        import traceback
        traceback.print_exc()

    finally:
        if doc:
            doc.Close(False)
            print(" Documento cerrado.")
        if word:
            word.Quit()
            print(" Aplicación Word cerrada.")
        if 'comtypes.client' in sys.modules:
            comtypes.CoUninitialize()
            print(" Recursos COM liberados.")


if __name__ == "__main__":
    print(" Buscando en:", base_dir)

    carpeta_word = base_dir
    carpeta_destino = os.path.join(base_dir, 'PDFs_Finales')

    convertir_word_a_pdfs_paginas_ordenadas(carpeta_word, carpeta_destino)
