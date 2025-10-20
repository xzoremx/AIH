from flask import Flask, render_template_string, request
import pandas as pd
import os
import sys
import tkinter as tk
from tkinter import simpledialog



if len(sys.argv) > 1:
    font_size_global = sys.argv[1]
else:
    print("No se proporcionó tamaño de fuente. Abortando servidor.")
    sys.exit(0)


if hasattr(sys, '_MEIPASS'):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))




app = Flask(__name__)

def aplicar_italic_especial(valor: str) -> str:
    """
    Aplica cursiva a nombres científicos salvo en excepciones
    Funcionan en cualquier parte del texto.
    """
    excepciones = {"sp.", "cf.", "(Grupo", "spp.", "aff."}
    palabras = valor.strip().split()

    resultado = []
    for palabra in palabras:
        if palabra in excepciones:
            resultado.append(palabra)  # normal
        else:
            resultado.append(f"<i>{palabra}</i>")  # cursiva

    return " ".join(resultado)



def leer_mapeo_combinadas(mapa_path):
    # Devuelve una lista de merges tipo: (row, col, rowspan, colspan)
    merges = []
    with open(mapa_path, encoding='utf-8') as f:
        for line in f:
            parts = line.strip().split(';')
            fila = int(parts[0].split('=')[1]) - 1  
            col = int(parts[1].split('=')[1]) - 1
            rowspan = int(parts[2].split('=')[1])
            colspan = int(parts[3].split('=')[1])
            merges.append((fila, col, rowspan, colspan))
    return merges


def leer_formato_celdas(formato_path):
    """Devuelve un dict: (fila, col) -> {'italic': bool, 'bold': bool, 'fmt': str, 'decimales': int|None}"""
    formatos = {}
    with open(formato_path, encoding='utf-8') as f:
        for line in f:
            # Ejemplo: FILA=2;COL=3;ITALIC=Verdadero;BOLD=Falso;FMT=General;DECIMALES=2
            partes = {kv.split('=')[0]: kv.split('=')[1] for kv in line.strip().split(';') if '=' in kv}
            fila = int(partes.get('FILA')) - 1  # cero-based
            col = int(partes.get('COL')) - 1
            italic = partes.get('ITALIC', 'Falso') == 'Verdadero'
            bold = partes.get('BOLD', 'Falso') == 'Verdadero'
            fmt = partes.get('FMT', '')
            decimales = partes.get('DECIMALES', '')
            try:
                decimales = int(decimales)
            except:
                decimales = None
            formatos[(fila, col)] = {
                'italic': italic,
                'bold': bold,
                'fmt': fmt,
                'decimales': decimales
            }
    return formatos


def formatea_valor(valor, formato):
    import re
    try:
        v = str(valor).replace(" ","").replace(",","")
        if v.replace('.','',1).isdigit():
            valor = float(v)
            # Primero trata de usar 'decimales' del formato, luego el patrón en 'fmt', luego 2 por defecto
            decimales = formato.get('decimales')
            if decimales is None:
                m = re.search(r'0[.,](0+)', formato.get('fmt', ''))
                decimales = len(m.group(1)) if m else 2
            # Formatea con espacio en miles y coma en decimales
            texto = f"{valor:,.{decimales}f}".replace(",", "X").replace(".", ",").replace("X", " ")
            valor = texto
        else:
            valor = str(valor)
    except Exception:
        valor = str(valor)

    if formato.get('italic'):
        # Aplica cursiva especial por defecto
        valor = aplicar_italic_especial(valor)

        # Excepciones globales: remover cursiva si el valor contiene estas palabras
        excepciones_no_cursiva = {"Nd", "Genero", "Género"}
        for exc in excepciones_no_cursiva:
            if exc.lower() in str(valor).lower():
                # Remover todas las etiquetas <i>…</i> si contiene la excepción
                valor = valor.replace("<i>", "").replace("</i>", "")
                break

    if formato.get('bold'):
        valor = f"<b>{valor}</b>"
    return valor



def construir_tabla_html(csv_path, mapa_path, formato_path, font_size="14pt"):
    import pandas as pd
    df = pd.read_csv(csv_path, encoding='cp1252', header=None)
    df = df.fillna("")
    merges = leer_mapeo_combinadas(mapa_path)
    formatos = leer_formato_celdas(formato_path)

    nrows, ncols = df.shape
    ocupado = [[False]*ncols for _ in range(nrows)]
    html = [f'<table border="1" style="border-collapse:collapse; font-size:{font_size};">']

    for r in range(nrows):
        html.append('<tr>')
        for c in range(ncols):
            if ocupado[r][c]:
                continue
            merge = next(((fr, fc, rs, cs) for (fr, fc, rs, cs) in merges if fr == r and fc == c), None)
            formato = formatos.get((r, c), {'italic': False, 'bold': False, 'fmt': ''})
            valor = df.iat[r,c]
            try:
                if isinstance(valor,str) and valor.replace('.','',1).replace(',','',1).isdigit():
                    valor = float(valor.replace(',','.'))
            except Exception:
                pass
            valor = formatea_valor(valor,formato)
            attr=""
            if merge:
                _, _, rowspan, colspan = merge
                for dr in range(rowspan):
                    for dc in range(colspan):
                        if not (dr == 0 and dc == 0):
                            ocupado[r+dr][c+dc] = True
                if rowspan > 1:
                    attr += f' rowspan="{rowspan}"'
                if colspan > 1:
                    attr += f' colspan="{colspan}"'
            html.append(f'<td{attr}>{valor}</td>')
        html.append('</tr>')
    html.append('</table>')
    return ''.join(html)


@app.route('/')
def mostrar_tabla():
    # Por defecto muestra la primera tabla que encuentre en el directorio
    directorio = os.path.join(base_dir, "PDFs_Iniciales")
    archivos = sorted([f for f in os.listdir(directorio) if f.endswith('.csv')])
    if not archivos:
        return "<h2>No se encontraron archivos CSV en el directorio</h2>"

    # Permite elegir la tabla vía parámetro
    tabla = request.args.get('tabla', archivos[0][:-4])
    csv_path = os.path.join(directorio, f"{tabla}.csv")
    mapa_path = os.path.join(directorio, f"{tabla}_mapa.txt")
    formato_path = os.path.join(directorio, f"{tabla}_formato.txt")
    if not os.path.exists(csv_path) or not os.path.exists(mapa_path):
        return f"<h2>No se encontró el CSV o el mapeo para: {tabla}</h2>"

    tabla_html = construir_tabla_html(csv_path, mapa_path, formato_path, font_size=font_size_global)

    # Listado de todas las tablas para navegación
    links = [f'<a href="/?tabla={f[:-4]}">{f[:-4]}</a>' for f in archivos]

    return render_template_string("""
    <html>
    <head>
    <title>Visualizador de Tablas (con celdas combinadas)</title>
    <style>
    @media print {
        body * { visibility: hidden; }
        table, table * { visibility: visible; }
        table { position: absolute; left: 0; top: 0; width: 100vw !important; }
    }
    </style>
    <style>
    body { font-family: Arial, sans-serif; margin: 10px; }
    table { width: 100%; max-width: none; }
    th, td {
        padding: 3px 1px;
        text-align: center;
        word-break: normal;
        white-space: normal;
        word-wrap: break-word;
        overflow-wrap: break-word;
    }
    </style>
    </head>
    <body>
    <h2>Tablas disponibles: {{ links|safe }}</h2>
    <hr>
    {{ tabla_html|safe }}
    <hr>
    </body>
    </html>
    """, links=" | ".join(links), tabla_html=tabla_html)



if __name__ == '__main__':
    app.run(port=5000, debug=True)




    

