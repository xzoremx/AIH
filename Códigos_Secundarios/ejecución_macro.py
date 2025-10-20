import os
import time
import win32com.client
import sys

if hasattr(sys, '_MEIPASS'):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

# === Buscar archivo Excel .xlsm en la carpeta base, excluyendo temporales ===
xlsm_files = [
    f for f in os.listdir(base_dir)
    if f.lower().endswith('.xlsm') and not f.startswith('~$')
]
if not xlsm_files:
    print("No se encontró ningún archivo .xlsm válido en la carpeta raíz.")
    sys.exit(1)

ruta_excel = os.path.join(base_dir, xlsm_files[0])
print(f"Usando Excel encontrado en carpeta: {ruta_excel}")

# === Ruta del archivo .bas (módulo VBA) ===
ruta_bas = os.path.join(base_dir, 'Códigos_Secundarios', 'Código_Macro', 'Tabla_Exportar.bas')

if not os.path.exists(ruta_bas):
    print(f"No se encontró el archivo .bas en: {ruta_bas}")
    sys.exit(1)

# === Iniciar Excel ===
excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

wb = None
try:
    # === Abrir el archivo Excel ===
    wb = excel.Workbooks.Open(ruta_excel)
    vb_proj = wb.VBProject

    # === Nombre del módulo VBA a importar ===
    modulo_nombre = "MD_Exportador"

    # === Eliminar módulo si ya existe ===
    for comp in vb_proj.VBComponents:
        if comp.Name == modulo_nombre:
            vb_proj.VBComponents.Remove(comp)
            break

    # === Importar el módulo desde el archivo .bas ===
    vb_proj.VBComponents.Import(ruta_bas)

    # === Ejecutar la macro ===
    try:
        print("Ejecutando macro...")
        excel.Application.Run(f"{modulo_nombre}.ExportarTodasLasHojasAPDF_Robusto")
        print("Macro ejecutada correctamente.")
    except Exception as e:
        print(f"Error al ejecutar la macro: {e}")

    # === Guardar y cerrar ===
    print("Guardando archivo...")
    wb.Save()

    print("Esperando para liberar procesos internos de Excel...")
    time.sleep(2)

    print("Cerrando archivo...")
    wb.Close(SaveChanges=0)

except Exception as e:
    print(f"Ocurrió un error en el flujo principal: {e}")

finally:
    if excel is not None:
        print("Cerrando Excel completamente...")
        excel.Quit()
    print("¡Proceso finalizado y Excel cerrado!")



