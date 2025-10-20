import argparse
import csv
import os
import re
import shutil
from glob import glob
from tempfile import NamedTemporaryFile
import unicodedata
import sys

# ================== Paths base (compatible PyInstaller) ==================
if hasattr(sys, '_MEIPASS'):
    base_dir = sys._MEIPASS
else:
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

# Carpeta por defecto donde están los CSV exportados
carpeta_csv = os.path.join(base_dir, "PDFs_Iniciales")

# ================== Utilidades ==================

def normalize_text(s: str) -> str:
    """Normaliza solo para detección (NO usar para escribir contenido)."""
    if s is None:
        return ""
    s = str(s).replace("\u00a0", " ").strip()  # NBSP -> espacio
    # remover tildes
    s = "".join(c for c in unicodedata.normalize("NFD", s)
                if unicodedata.category(c) != "Mn")
    return s.lower()

def detect_encoding(csv_path: str) -> str:
    """
    Detecta la codificación del CSV: intenta UTF-8; si falla, usa cp1252.
    Devuelve "utf-8" o "cp1252".
    """
    try:
        with open(csv_path, "r", encoding="utf-8", newline="") as f:
            f.read(2048)
        return "utf-8"
    except UnicodeDecodeError:
        return "cp1252"

def sniff_delimiter(csv_path: str, encoding: str) -> str:
    """Intenta detectar el delimitador del CSV en base a una muestra."""
    with open(csv_path, "r", encoding=encoding, errors="ignore", newline="") as f:
        sample = f.read(4096)
        sniffer = csv.Sniffer()
        try:
            dialect = sniffer.sniff(sample, delimiters=[",",";","\t","|"])
            return dialect.delimiter
        except csv.Error:
            # fallback común en ES
            return ";"

def count_less_than(arr_sorted, k):
    # arr_sorted: enteros 1-based, orden ascendente
    return sum(1 for x in arr_sorted if x < k)

def contiguous_ranges(nums):
    """Agrupa enteros en segmentos contiguos [(a,b), ...]."""
    if not nums:
        return []
    nums = sorted(nums)
    out = []
    a = b = nums[0]
    for x in nums[1:]:
        if x == b + 1:
            b = x
        else:
            out.append((a, b))
            a = b = x
    out.append((a, b))
    return out

# ================== Detección de columnas "Incertidumbre" ==================

def find_incertidumbre_cols(csv_path, header_scan_rows=15):
    """
    Escanea las N primeras filas del CSV.
    Si encuentra una celda cuyo texto (normalizado) COMIENZA con 'incertidumbre',
    marca esa columna y también la siguiente (merge de 2 columnas).
    Devuelve set (1-based) de columnas a eliminar.
    """
    cols_to_remove = set()
    if not os.path.exists(csv_path):
        return cols_to_remove

    enc = detect_encoding(csv_path)
    delim = sniff_delimiter(csv_path, enc)

    with open(csv_path, "r", encoding=enc, errors="ignore", newline="") as f:
        reader = csv.reader(f, delimiter=delim)
        for r, row in enumerate(reader, start=1):
            if r > header_scan_rows:
                break
            for j, val in enumerate(row, start=1):
                if not val:
                    continue
                nv = normalize_text(val)
                if nv.startswith("incertidumbre"):
                    cols_to_remove.add(j)
                    cols_to_remove.add(j + 1)  # asumimos merge de 2 columnas

    return set(sorted(x for x in cols_to_remove if x >= 1))

# ================== CSV: eliminar columnas ==================

def drop_columns_from_csv(csv_path, cols_to_remove):
    if not cols_to_remove or not os.path.exists(csv_path):
        return

    enc = detect_encoding(csv_path)
    delim = sniff_delimiter(csv_path, enc)
    drop_idx = sorted((c - 1 for c in cols_to_remove), reverse=True)  # 0-based desc

    tmp = NamedTemporaryFile("w", delete=False, newline="", encoding=enc)
    tmp_path = tmp.name
    with open(csv_path, "r", encoding=enc, errors="ignore", newline="") as fin, tmp:
        reader = csv.reader(fin, delimiter=delim)
        writer = csv.writer(tmp, delimiter=delim)
        for row in reader:
            new_row = list(row)
            for idx in drop_idx:
                if 0 <= idx < len(new_row):
                    del new_row[idx]
            writer.writerow(new_row)
    shutil.move(tmp_path, csv_path)

# ================== _mapa.txt: ajustar merges ==================

MAPA_RE = re.compile(
    r"FILA=(?P<FILA>\d+);COL=(?P<COL>\d+);ROWS=(?P<ROWS>\d+);COLS=(?P<COLS>\d+)",
    re.I
)

def adjust_mapa(mapa_path, cols_to_remove):
    """
    Para cada merge (COL..COL+COLS-1):
      - Quita columnas de cols_to_remove.
      - Si queda vacío, descarta.
      - Reindexa columnas.
      - **Une** sub-intervalos que tras reindexar quedan contiguos,
        emitiendo un único merge (evita 'cortes' visuales).
    """
    if not os.path.exists(mapa_path):
        return

    cols_to_remove = sorted(set(cols_to_remove))  # 1-based asc
    tmp_path = mapa_path + ".tmp"

    with open(mapa_path, "r", encoding="utf-8", errors="ignore") as fin, \
         open(tmp_path, "w", encoding="utf-8", newline="") as fout:

        for line in fin:
            m = MAPA_RE.search(line)
            if not m:
                fout.write(line)
                continue

            FILA = int(m.group("FILA"))
            COL  = int(m.group("COL"))
            ROWS = int(m.group("ROWS"))
            COLS = int(m.group("COLS"))

            start = COL
            end   = COL + COLS - 1

            # columnas originales del merge
            orig_cols = list(range(start, end + 1))
            # columnas que permanecen (quitamos las borradas)
            remain = [c for c in orig_cols if c not in cols_to_remove]
            if not remain:
                continue

            # 1) dividir en bloques contiguos en el espacio original
            parts = contiguous_ranges(remain)  # [(a,b), ...]

            # 2) transformar cada bloque al nuevo índice (reindexar)
            transformed = []
            for a, b in parts:
                width = b - a + 1
                shift = count_less_than(cols_to_remove, a)   # #borradas < a
                new_col = a - shift
                transformed.append((new_col, width))

            # 3) **coalesce**: unir bloques contiguos tras reindexar
            transformed.sort(key=lambda t: t[0])
            merged = []
            for c0, w0 in transformed:
                if not merged:
                    merged.append((c0, w0))
                else:
                    pc, pw = merged[-1]
                    if c0 == pc + pw:  # justo pegado al final → unir
                        merged[-1] = (pc, pw + w0)
                    else:
                        merged.append((c0, w0))

            # 4) emitir líneas finales
            for c0, w0 in merged:
                fout.write(f"FILA={FILA};COL={c0};ROWS={ROWS};COLS={w0}\n")

    shutil.move(tmp_path, mapa_path)
    
# ================== _formato.txt: ajustar columnas ==================

def adjust_formato(fmt_path, cols_to_remove):
    if not os.path.exists(fmt_path):
        return
    cols_to_remove = sorted(set(cols_to_remove))  # 1-based asc
    tmp_path = fmt_path + ".tmp"

    with open(fmt_path, "r", encoding="utf-8", errors="ignore") as fin, \
         open(tmp_path, "w", encoding="utf-8", newline="") as fout:

        for raw in fin:
            if "COL=" not in raw:
                fout.write(raw)
                continue
            try:
                col = int(re.search(r"COL=(\d+)", raw, flags=re.I).group(1))
            except Exception:
                # línea no parseable → re-emite
                fout.write(raw)
                continue

            if col in cols_to_remove:
                # descartar entradas de columnas eliminadas
                continue

            shift = count_less_than(cols_to_remove, col)
            new_col = col - shift
            new = re.sub(r"COL=\d+", f"COL={new_col}", raw, flags=re.I)
            fout.write(new)

    shutil.move(tmp_path, fmt_path)

# ================== Pipeline por prefix ==================

def process_prefix(prefix_path, header_scan_rows=15, verbose=True):
    csv_path  = prefix_path + ".csv"
    mapa_path = prefix_path + "_mapa.txt"
    fmt_path  = prefix_path + "_formato.txt"

    # 1) detectar columnas a eliminar en CSV
    cols_to_remove = find_incertidumbre_cols(csv_path, header_scan_rows=header_scan_rows)
    if not cols_to_remove:
        if verbose:
            print(f"[SKIP] No se detectó 'Incertidumbre' en: {os.path.basename(csv_path)}")
        return

    if verbose:
        # usar ASCII para evitar UnicodeEncodeError en consolas CP1252
        print(f"[OK]  {os.path.basename(prefix_path)} -> eliminar columnas: {sorted(cols_to_remove)}")

    # 2) aplicar a CSV
    drop_columns_from_csv(csv_path, cols_to_remove)

    # 3) ajustar mapa y formato
    adjust_mapa(mapa_path, cols_to_remove)
    adjust_formato(fmt_path, cols_to_remove)

# ================== CLI ==================

def main():
    # Evitar errores de consola por codificación en Windows
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

    ap = argparse.ArgumentParser(
        description="Elimina columnas 'Incertidumbre' (2 cols) detectadas en CSV y ajusta *_mapa/_formato."
    )
    ap.add_argument("--root", default=carpeta_csv, help="Carpeta raíz con los archivos exportados")
    ap.add_argument("--pattern", default="*_mapa.txt", help="Patrón para localizar mapas y derivar prefix")
    ap.add_argument("--scan-rows", type=int, default=15, help="Filas superiores del CSV donde buscar 'Incertidumbre'")
    ap.add_argument("--quiet", action="store_true", help="Ocultar logs por archivo")
    args = ap.parse_args()

    root = args.root if os.path.isdir(args.root) else carpeta_csv
    mapas = glob(os.path.join(root, args.pattern))
    if not mapas:
        print(f"[INFO] No se encontraron mapas con patrón: {args.pattern} en {root}")
        return

    for mapa in mapas:
        prefix = mapa[:-len("_mapa.txt")]
        process_prefix(prefix, header_scan_rows=args.scan_rows, verbose=not args.quiet)

if __name__ == "__main__":
    main()
