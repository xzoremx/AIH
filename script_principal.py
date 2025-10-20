import subprocess
import sys
from pathlib import Path
import os
import shutil
import time
import tkinter as tk
from tkinter import simpledialog, messagebox

# Importar el nuevo selector de archivos (ubicado en Códigos_Secundarios)
from Códigos_Secundarios.selector_archivos import seleccionar_archivos

# Configuración de rutas
CODIGOS_DIR = "Códigos_Secundarios"
CARPETAS_VALIDACION = ["PDFs_Finales", "PDFs_Iniciales", "PDFs_Impresos"]
SCRIPTS = {
    "macro": "ejecución_macro.py",
    "tabla": "creador_tablas.py",
    "impresión": "imprimir_tablas.py",
    "word": "conversión_word.py",
    "formato": "aplicación_formato.py",
    "union": "unión&enumeración.py",
    "preprocesamiento_csv": "preprocesamiento_csv.py",
    "config_inicial": "config_inicial.py"
}

# Variables globales para rutas elegidas
WORD_PATH = None
EXCEL_PATH = None

def eliminar_archivos_word_excel():
    """Elimina todos los archivos Word y Excel que estén en la carpeta del script principal."""
    raiz = Path(__file__).resolve().parent  # Carpeta donde está script_principal.py
    eliminados = []

    for ext in ["*.doc", "*.docx", "*.xls", "*.xlsx", "*.xlsm"]:
        for archivo in raiz.glob(ext):
            try:
                archivo.unlink()
                eliminados.append(archivo.name)
            except Exception as e:
                print(f" Error al eliminar {archivo}: {e}")

    if eliminados:
        print(f" Archivos eliminados en carpeta raíz: {', '.join(eliminados)}")
    else:
        print(" No había archivos Word/Excel para eliminar en carpeta raíz.")


def preguntar_tamano_fuente():
    root = tk.Tk()
    root.withdraw()
    valor = simpledialog.askinteger(
        "Tamaño de letra",
        "Ingrese el tamaño de fuente para la tabla (solo número, ej: 10, 16, 20, 30):",
        initialvalue=14,
        minvalue=6,
        maxvalue=30
    )
    root.destroy()
    return f"{valor}pt" if valor else None


def confirmar_config_inicial():
    root = tk.Tk()
    root.withdraw()
    resp = messagebox.askyesno(
        "Confirmación",
        "¿Continuar con los requisitos predefinidos para la tabla de resultados?\n"
    )
    root.destroy()
    return bool(resp)


def vaciar_carpetas():
    for carpeta in CARPETAS_VALIDACION:
        try:
            Path(carpeta).mkdir(exist_ok=True)
            for item in Path(carpeta).iterdir():
                try:
                    if item.is_file():
                        item.unlink()
                    elif item.is_dir():
                        shutil.rmtree(item)
                except Exception as e:
                    print(f" Error al eliminar {item}: {e}")
                    return False
            print(f" Carpeta {carpeta} vaciada correctamente")
        except Exception as e:
            print(f" Error crítico al vaciar {carpeta}: {e}")
            return False
    return True


def ejecutar_script(nombre_script, wait=True, args=None):
    ruta_script = Path(CODIGOS_DIR) / SCRIPTS[nombre_script]
    if not ruta_script.exists():
        print(f" Error: No se encontró el script {ruta_script}")
        return False if wait else None

    comando = [sys.executable, str(ruta_script)]
    if args:
        comando.extend(args)

    # Pasar rutas de Word y Excel como variables de entorno
    env = os.environ.copy()
    if WORD_PATH:
        env["WORD_PATH"] = WORD_PATH
    if EXCEL_PATH:
        env["EXCEL_PATH"] = EXCEL_PATH

    print(f"\n Ejecutando {SCRIPTS[nombre_script]}{' (en background)' if not wait else ''}...")
    try:
        if wait:
            resultado = subprocess.run(
                comando,
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding='utf-8',
                errors='replace',
                env=env
            )
            if resultado.stdout:
                print(" Salida del script:")
                print(resultado.stdout)
            return True
        else:
            return subprocess.Popen(comando, env=env)
    except subprocess.CalledProcessError as e:
        print(f" Error en {SCRIPTS[nombre_script]}:")
        print(f"Código de error: {e.returncode}")
        if e.stderr:
            print(" Mensaje de error:")
            print(e.stderr)
        return False if wait else None
    except Exception as e:
        print(f" Error inesperado en {SCRIPTS[nombre_script]}: {e}")
        return False if wait else None


def preprocesamiento_csvs_en_pdfs_iniciales():
    carpeta = Path("PDFs_Iniciales")
    archivos_csv = list(carpeta.glob("*.csv"))

    if not archivos_csv:
        print(" No se encontraron archivos CSV para preprocesar.")
        return

    print(f"\n Preprocesando {len(archivos_csv)} archivos CSV...")
    for archivo in archivos_csv:
        resultado = ejecutar_script("preprocesamiento_csv", wait=True, args=[archivo.name])
        if not resultado:
            print(f"Falló el preprocesamiento de {archivo.name}")
        else:
            print(f"Preprocesamiento completo: {archivo.name}")


def main():
    global WORD_PATH, EXCEL_PATH

    print("═══════════════════════════════════════")
    print("  INICIANDO FLUJO DE PROCESAMIENTO")
    print("═══════════════════════════════════════")

    # Paso 0: Eliminar Word y Excel previos
    print("\n Eliminando archivos Word y Excel en carpeta raíz...")
    eliminar_archivos_word_excel()

    # Paso 1: Seleccionar Word y Excel
    archivos = seleccionar_archivos()
    WORD_PATH = archivos["word"]
    EXCEL_PATH = archivos["excel"]

    if not WORD_PATH or not EXCEL_PATH:
        print("\n No se seleccionaron archivos. Abortando flujo...")
        sys.exit(0)

    # Paso 2: Tamaño de fuente
    font_size = preguntar_tamano_fuente()
    if not font_size:
        print("\n No se seleccionó tamaño de fuente. Abortando flujo...")
        sys.exit(0)

    # Paso 3: Preparar carpetas
    print("\n Verificando y vaciando carpetas requeridas:")
    if not vaciar_carpetas():
        print("\n No se pudo preparar el entorno. Abortando flujo...")
        sys.exit(1)

    # Paso 4: Ejecutar macro en Excel
    print("\n═════════ EJECUTANDO MACRO ═════════")
    if not ejecutar_script("macro"):
        print("\n Falló ejecución_macro.py. Abortando flujo...")
        sys.exit(1)
    print(" ejecución_macro.py completado exitosamente")

    

    # Paso 5: Ajuste primario
    print("\n═════════ AJUSTE PRIMARIO: COLUMNAS 'INCERTIDUMBRE' ═════════")
    if confirmar_config_inicial():
        print(" Opción 'Sí' seleccionada: ejecutando config_inicial.py...")
        if not ejecutar_script("config_inicial"):
            print("\n Falló config_inicial.py. Abortando flujo...")
            sys.exit(1)
        print(" config_inicial.py completado exitosamente")
    else:
        print(" Opción 'No' seleccionada: se omite config_inicial.py.")

    # Paso 5.5: Ajuste adicional
    print("\n═════════ PREPROCESAMIENTO DE CSVS ═════════")
    preprocesamiento_csvs_en_pdfs_iniciales()

    # Paso 6: Servidor Flask + impresión
    print("\n═════════ EJECUTANDO SERVIDOR FLASK + IMPRESIÓN (EN PARALELO) ═════════")
    proc_tabla = ejecutar_script("tabla", wait=False, args=[font_size])
    time.sleep(5)
    proc_impresion = ejecutar_script("impresión", wait=True)
    print(" impresión completada exitosamente")
    if proc_tabla is not None:
        print(" Finalizando servidor Flask...")
        proc_tabla.terminate()
        proc_tabla.wait()
        print(" Servidor Flask detenido.")

    # Paso 7: Conversión Word
    print("\n═════════ EJECUTANDO CONVERSIÓN WORD ═════════")
    if not ejecutar_script("word"):
        print("\n Falló conversión_word.py. Abortando flujo...")
        sys.exit(1)
    print(" conversión_word.py completado exitosamente")

    # Paso 8: Formato
    print("\n═════════ EJECUTANDO FORMATO ═════════")
    if not ejecutar_script("formato"):
        print("\n Falló aplicación_formato.py. Abortando flujo...")
        sys.exit(1)
    print(" aplicación_formato.py completado exitosamente")

    # Paso 9: Unión y enumeración
    print("\n═════════ EJECUTANDO UNIÓN Y ENUMERACIÓN ═════════")
    if not ejecutar_script("union"):
        print("\n Flujo completado con errores en la unión/enumeración")
        sys.exit(1)
    print(" unión&enumeración.py completado exitosamente")

    print("\n═══════════════════════════════════════")
    print("  FLUJO COMPLETADO EXITOSAMENTE")
    print("═══════════════════════════════════════")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        print("\n Error inesperado durante el procesamiento:")
        traceback.print_exc()

   



