import argparse
import csv
import os
import re
import sys
import unicodedata
from glob import glob

# ------------------- Paths base (compatible PyInstaller) -------------------
if hasattr(sys, '_MEIPASS'):
    base_dir = sys._MEIPASS  # type: ignore[attr-defined]
else:
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

carpeta_csv = os.path.join(base_dir, "PDFs_Iniciales")

# ------------------- Utilidades -------------------------------------------
def normalize_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s).replace("\u00a0", " ").strip()
    s = "".join(c for c in unicodedata.normalize("NFD", s)
                if unicodedata.category(c) != "Mn")
    return s.lower()

def detect_encoding(csv_path: str) -> str:
    try:
        with open(csv_path, "r", encoding="utf-8", newline="") as f:
            f.read(2048)
        return "utf-8"
    except UnicodeDecodeError:
        return "cp1252"

def sniff_delimiter(csv_path: str, encoding: str) -> str:
    with open(csv_path, "r", encoding=encoding, errors="ignore", newline="") as f:
        sample = f.read(4096)
        sniffer = csv.Sniffer()
        try:
            dialect = sniffer.sniff(sample, delimiters=[",", ";", "\t", "|"])
            return dialect.delimiter
        except csv.Error:
            return ";"

# ------------------- Preprocesamiento: limpiar CSV ------------------------
def limpiar_csv(ruta_csv: str):
    if not os.path.isfile(ruta_csv):
        print(f"No se encontró el archivo: {ruta_csv}")
        return

    enc = detect_encoding(ruta_csv)
    with open(ruta_csv, "r", encoding=enc, errors="ignore") as f:
        lineas = f.readlines()

    # Eliminar SOLO filas vacías al final
    idx = len(lineas) - 1
    while idx >= 0:
        linea = lineas[idx].strip().replace(",", "")
        if linea != "":
            break
        idx -= 1

    nuevas = lineas[:idx + 1]
    eliminadas = len(lineas) - len(nuevas)

    with open(ruta_csv, "w", encoding=enc, newline="") as f:
        f.writelines(nuevas)

    print(f"Archivo limpiado: {ruta_csv} (se eliminaron {eliminadas} filas vacías al final)")

# ------------------- Observaciones Nd / --- / ND -------------------------------
def insertar_observaciones(csv_path: str, mapa_path: str,
                           formato_path: str, decimales_path: str):
    """
    Detecta si en el CSV existen 'Nd', 'ND' o '---' y agrega observaciones
    correspondientes al final, actualizando también mapa, formato y decimales.
    """

    if not os.path.exists(csv_path) or not os.path.exists(mapa_path):
        print(f"[SKIP] Faltan archivos base para {csv_path}")
        return

    # --- Detectar encoding del CSV ---
    enc = detect_encoding(csv_path)

    # --- Leer CSV con encoding correcto ---
    with open(csv_path, "r", encoding=enc, newline="") as f:
        rows = [row for row in csv.reader(f, delimiter=",")]

    # Detectar existencia exacta de Nd / ND / ---
    has_nd  = any(cell.strip() == "Nd" for row in rows for cell in row if cell)
    has_ND  = any(cell.strip() == "ND" for row in rows for cell in row if cell)
    has_dash = any(cell.strip() == "---" for row in rows for cell in row if cell)

    if not has_nd and not has_ND and not has_dash:
        print(f"[NO_OBS] {os.path.basename(csv_path)}")
        return

    # --- Leer mapa.txt y ubicar última fila/cols ---
    with open(mapa_path, "r", encoding="utf-8") as f:
        mapa_lines = f.read().strip().splitlines()

    last_line = mapa_lines[-1]
    parts = dict(p.split("=") for p in last_line.split(";"))
    last_row = int(parts["FILA"])
    col_count = int(parts["COLS"])

    new_rows = []
    if has_nd:
        new_rows.append("Nd: No determinado")
    if has_ND:
        new_rows.append("ND: No Detectable")
    if has_dash:
        new_rows.append("[---]: No se encontraron organismos, no aplica incertidumbre.")

    # --- Agregar nuevas filas al CSV ---
    for obs in new_rows:
        rows.append([obs] + [""] * (col_count - 1))

    with open(csv_path, "w", encoding=enc, newline="") as f:
        writer = csv.writer(f, delimiter=",", lineterminator="\n")
        writer.writerows(rows)

    # --- Actualizar mapa.txt ---
    for i in range(len(new_rows)):
        fila = last_row + i + 1
        mapa_lines.append(f"FILA={fila};COL=1;ROWS=1;COLS={col_count}")

    with open(mapa_path, "w", encoding="utf-8") as f:
        f.write("\n".join(mapa_lines) + "\n")

    # --- Actualizar formato.txt ---
    if os.path.exists(formato_path):
        with open(formato_path, "r", encoding="utf-8") as f:
            formato_lines = f.read().strip().splitlines()
    else:
        formato_lines = []

    for i in range(len(new_rows)):
        fila = last_row + i + 1
        for c in range(1, col_count + 1):
            formato_lines.append(
                f"FILA={fila};COL={c};ITALIC=Falso;BOLD=Falso;FMT=General;DECIMALES="
            )

    with open(formato_path, "w", encoding="utf-8") as f:
        f.write("\n".join(formato_lines) + "\n")

    # --- Actualizar decimales.txt ---
    if os.path.exists(decimales_path):
        with open(decimales_path, "r", encoding="utf-8") as f:
            dec_lines = f.read().strip().splitlines()
    else:
        dec_lines = []

    for i in range(len(new_rows)):
        fila = last_row + i + 1
        for c in range(1, col_count + 1):
            dec_lines.append(f"FILA={fila};COL={c};DECIMALES=")

    with open(decimales_path, "w", encoding="utf-8") as f:
        f.write("\n".join(dec_lines) + "\n")

    print(f"[OBS_OK] {os.path.basename(csv_path)} → se agregaron {len(new_rows)} observaciones.")



# ------------------- Lógica Macrozoobentos -------------------------------
TITLE_NEEDLE = "determinacion de macrozoobentos"
HEADER_TOKENS = {"clase", "orden", "familia", "genero", "especie"}

CATEG_MAP = {
    1: "Filo",
    2: "Subclase",
    3: "Infraclase",
    4: "Suborden",
    5: "Superfamilia",
    6: "Subphylum",
    7: "Subfamilia",
    8: "Superorden",
}
NUM_PAREN_RE = re.compile(r"\((\d{1,2})\)")

def has_macrozoobentos_title(rows, title_rows=10) -> bool:
    upto = min(title_rows, len(rows))
    for r in range(upto):
        for val in rows[r]:
            if TITLE_NEEDLE in normalize_text(val):
                return True
    return False

def find_header_row(rows, min_hits=3):
    for r, row in enumerate(rows):
        hits = 0
        for val in row:
            nv = normalize_text(val)
            for token in HEADER_TOKENS:
                if token in nv:
                    hits += 1
        if hits >= min_hits:
            return r
    return None

def find_categoria_label_cells(rows):
    labels = []
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            nv = normalize_text(val)
            if "categoria" in nv and "taxonom" in nv:
                labels.append((r, c))
    return labels

def extract_markers_from_rows(rows, start_row):
    found = set()
    for r in range(start_row, len(rows)):
        for val in rows[r]:
            if not val:
                continue
            for m in NUM_PAREN_RE.findall(val):
                try:
                    n = int(m)
                    if n in CATEG_MAP:
                        found.add(n)
                except ValueError:
                    pass
    return sorted(found)

def update_categorias_taxonomicas(csv_path: str,
                                  title_rows: int = 10,
                                  min_header_hits: int = 3,
                                  scan_offset: int = 2,
                                  sep: str = "; ",
                                  prefix: str = "Categorías taxonómicas: ",
                                  update_all_labels: bool = False) -> str:
    if not os.path.exists(csv_path):
        return "skip_missing"

    enc = detect_encoding(csv_path)
    delim = sniff_delimiter(csv_path, enc)

    with open(csv_path, "r", encoding=enc, errors="ignore", newline="") as f:
        rows = [row for row in csv.reader(f, delimiter=delim)]

    if not has_macrozoobentos_title(rows, title_rows=title_rows):
        return "skip_title"

    header_row = find_header_row(rows, min_hits=min_header_hits)
    if header_row is None:
        return "skip_header"

    start_row = header_row + scan_offset
    markers = extract_markers_from_rows(rows, start_row)

    labels = find_categoria_label_cells(rows)
    if not labels:
        return "skip_label"

    if markers:
      final_text = prefix + sep.join(f"({n}) {CATEG_MAP[n]}" for n in markers)
    else:
      final_text = prefix + "—"


    changed = False
    targets = labels if update_all_labels else labels[:1]
    for (r, c) in targets:
        if rows[r][c] != final_text:
            rows[r][c] = final_text
            changed = True

    if not changed:
        return "no_change"

    with open(csv_path, "w", encoding=enc, newline="") as f:
        csv.writer(f, delimiter=delim).writerows(rows)

    return "ok"

# ------------------- CLI ---------------------------------------------------
def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

    ap = argparse.ArgumentParser(
        description=("Preprocesa CSVs: limpia, inserta observaciones por 'Nd'/'---' "
                     "y actualiza 'Categorías taxonómicas' en tablas de DETERMINACIÓN DE MACROZOOBENTOS.")
    )
    ap.add_argument("--csv", action="append",
                    help="Ruta a un CSV (puedes repetir la bandera para varios).")
    ap.add_argument("--root", default=carpeta_csv,
                    help="Carpeta base si no se pasa --csv.")
    ap.add_argument("--pattern", default="*.csv",
                    help="Patrón glob para seleccionar CSVs en --root.")
    ap.add_argument("--title-rows", type=int, default=10)
    ap.add_argument("--min-header-hits", type=int, default=3)
    ap.add_argument("--scan-offset", type=int, default=2)
    ap.add_argument("--newline-sep", action="store_true")
    ap.add_argument("--update-all", action="store_true")
    ap.add_argument("--quiet", action="store_true")
    args, unknown = ap.parse_known_args()

    sep = "\n" if args.newline_sep else "; "

    # Resolver targets
    targets = []
    if args.csv:
        targets.extend(args.csv)

    if unknown:  # Argumentos sueltos (nombres con espacios no citados)
        joined = " ".join(unknown).strip()
        if joined:
            targets.append(joined)

    if not targets:
        targets = sorted(glob(os.path.join(args.root, args.pattern)))

    if not targets:
        print("[INFO] No se encontraron CSVs.")
        return

    for p in targets:
        # Resolver ruta relativa en carpeta base
        if not os.path.isabs(p):
            candidate = os.path.join(carpeta_csv, p)
            p = candidate if os.path.exists(candidate) else p

        # 1) Limpiar
        limpiar_csv(p)

        # 2) Insertar observaciones (Nd / ---)
        base_name = os.path.splitext(p)[0]
        mapa_path = base_name + "_mapa.txt"
        formato_path = base_name + "_formato.txt"
        decimales_path = base_name + "_decimales.txt"
        insertar_observaciones(p, mapa_path, formato_path, decimales_path)

        # 3) Actualizar categorías (si aplica a macrozoobentos)
        status = update_categorias_taxonomicas(
            p,
            title_rows=args.title_rows,
            min_header_hits=args.min_header_hits,
            scan_offset=args.scan_offset,
            sep=sep,
            update_all_labels=args.update_all,
        )
        if not args.quiet:
            print(f"[{status.upper()}] {os.path.basename(p)}")


if __name__ == "__main__":
    main()







