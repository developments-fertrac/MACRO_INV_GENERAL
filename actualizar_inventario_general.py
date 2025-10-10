# actualizar_inventario_integral.py
# -*- coding: utf-8 -*-

from __future__ import annotations
import io, re, os, contextlib
from pathlib import Path
from datetime import date, datetime
import pandas as pd
import numpy as np
import msoffcrypto
from unidecode import unidecode
import tempfile

# ==== CONFIG ====
# BASE_PATH = Path(r"C:\Users\jperez\Desktop\Tecnologia\Inventario General")
# BASE_PATH = Path(r"C:\MACRO_INVENTARIO_GENERAL")
BASE_PATH = Path(__file__).resolve().parent  

PASS_INV = "Compras2025"
PASSWORDS_TRY = ["Compras2025"]  # añade más si algún archivo usa otra clave

OUTPUT_BASENAME = "$2025 INVENTARIO GENERAL ACTUALIZADO"  # sin extensión
APPLY_PASSWORD_TO_OUTPUT = True  # aplica contraseña al archivo nuevo

# Prefijos para ubicar archivos descargados del ERP (tolerantes a sufijos/fechas)
PFX_INV_ACTUALIZADO = "INVENTARIO GENERAL ACTUALIZADO"        # “… (MAÑANA)”
PFX_VAL_GENERAL     = "VALORIZADO GENERAL"
PFX_VAL_FALT_IMPO   = "VALORIZADO FALTANTES IMPO"
PFX_VAL_FALT        = "VALORIZADO FALTANTES"
PFX_VAL_TOBERIN     = "VALORIZADO TOBERIN"

FN_INV_PLANTILLA = "$2025 INVENTARIO GENERAL.xlsx"
SHEET_INV_ORIG   = "INVENTARIO"
SHEET_INV_LISTA  = "INV LISTA PRECIOS"

HEADER_ROW_INV         = 2   # fila visible 2 (indexada desde 1) en hoja INVENTARIO
HEADER_ROW_INV_LISTA   = 1   # fila visible 1 en INV LISTA PRECIOS
HEADER_ROW_VAL         = 9   # fila visible 9 en archivos VALORIZADO

# Columnas a limpiar en INVENTARIO (original)
COLS_A_LIMPIAR = [
    "REFERENCIA", "NOMBRE LISTA", "NOMBRE ODOO", "NOMBRE MYR",
    "MARCA copia", "INV BODEGA", "EXISTENCIA AGO 26", "COSTO PROMEDIO",
    "LINEA COPIA", "SUB-LINEA COPIA", "LIDER LINEA", "CLASIFICACION",
    "Marca sistema", "Linea sistema", "Sub- linea sistema"
]
# Paso 12: columnas a traer desde la hoja copia INVENTARIO
COLS_DESDE_COPIA = ["MARCA copia", "INV BODEGA GERENCIA", "LINEA COPIA", "SUB-LINEA COPIA", "LIDER LINEA", "CLASIFICACION"]

# ==== DEPENDENCIAS (COM) ====
try:
    import win32com.client as win32
    HAS_COM = True
except Exception:
    HAS_COM = False

# ==== UTILS BÁSICAS ====
def log(msg): print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

def _norm(s: str) -> str:
    t = unidecode(str(s)).lower()
    t = re.sub(r"[^a-z0-9 ]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t

def month_abbr_es(dt: date) -> str:
    abrs = ["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"]
    return abrs[dt.month-1]

def exist_col_title_for_today() -> str:
    today = date.today()
    return f"EXISTENCIA {month_abbr_es(today)} {today.day:02d}"

def to_num_str(x):
    """Convierte a referencia numérica segura (string sin .0)."""
    if pd.isna(x): return ""
    s = str(x).strip()
    with contextlib.suppress(Exception):
        f = float(s.replace(",",""))
        if abs(f - int(f)) < 1e-9:
            return str(int(f))
        return str(f)
    return s

# ==== ARCHIVOS / LECTURA PANDASeable ====
def _strip_dol_tmp(name: str) -> str:
    base = Path(name).stem.replace("~$", "")
    base = re.sub(r"^\$+", "", base)
    return base

def find_by_prefix(base_dir: Path, prefix: str, exts=(".xlsx",".xlsm",".xls",".csv")) -> Path:
    """Busca por prefijo normalizado, elige el más reciente."""
    pref = _norm(prefix)
    cands = []
    for f in base_dir.iterdir():
        if not (f.is_file() and f.suffix.lower() in exts):
            continue
        nn = _norm(_strip_dol_tmp(f.name))
        if nn.startswith(pref) or pref in nn:
            cands.append(f); continue
        tokens = pref.split()
        if all(t in nn for t in tokens):
            cands.append(f)
    if not cands:
        raise FileNotFoundError(f"No encontré archivos que coincidan con '{prefix}' en {base_dir}")
    cands.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return cands[0]

def decrypt_to_stream(xlsx_path: Path, password: str) -> io.BytesIO:
    bio = io.BytesIO()
    with open(xlsx_path, "rb") as f:
        office = msoffcrypto.OfficeFile(f)
        office.load_key(password=password)
        office.decrypt(bio)
    bio.seek(0)
    return bio

def is_encrypted_xlsx(path: Path) -> bool:
    try:
        with open(path, "rb") as f:
            of = msoffcrypto.OfficeFile(f)
            return bool(getattr(of, "is_encrypted", True))
    except Exception:
        return False

def save_bytesio_to_temp(bio: io.BytesIO, stem: str) -> Path:
    tmp = Path(tempfile.gettempdir()) / f"~dec_{stem}_{datetime.now().strftime('%H%M%S')}.xlsx"
    with open(tmp, "wb") as out:
        out.write(bio.getvalue())
    return tmp

def com_convert_to_xlsx(path: Path, passwords: list[str] | None = None) -> Path:
    """Convierte silenciosamente a .xlsx usando COM."""
    passwords = passwords or []
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.Interactive = False
    excel.EnableEvents = False
    try: excel.AskToUpdateLinks = False
    except Exception: pass
    try: excel.AutomationSecurity = 3  # deshabilita macros
    except Exception: pass
    try: excel.ScreenUpdating = False
    except Exception: pass

    def _force_invisible(app):
        try:
            app.Visible = False
            for i in range(1, app.Windows.Count + 1):
                try:
                    app.Windows(i).Visible = False
                except Exception:
                    pass
        except Exception:
            pass
    _force_invisible(excel)

    # Detectar cifrado
    encrypted = False
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        try:
            with open(path, "rb") as f:
                of = msoffcrypto.OfficeFile(f)
                encrypted = bool(getattr(of, "is_encrypted", False))
        except Exception:
            encrypted = False

    wb = None
    last_err = None
    pw_attempts = (passwords if encrypted else [None] + passwords)

    for pw in pw_attempts:
        try:
            if pw:
                wb = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True, Password=pw)
            else:
                wb = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True)
            break
        except Exception as e:
            last_err = e
            continue

    if wb is None:
        excel.Quit()
        msg = "archivo cifrado sin contraseña válida" if encrypted else "no pude abrir el archivo"
        raise RuntimeError(f"COM no pudo abrir '{path.name}': {msg}. Detalle: {last_err}")

    tmp = Path(tempfile.gettempdir()) / f"~conv_{path.stem}_{datetime.now().strftime('%H%M%S')}.xlsx"
    wb.SaveAs(str(tmp), FileFormat=51)  # 51 = xlOpenXMLWorkbook (.xlsx)
    wb.Close(SaveChanges=False)
    excel.Quit()
    return tmp

def open_as_excel_source(path: Path, passwords: list[str] | None = None):
    """
    Devuelve un 'source' para pandas: Path, BytesIO desencriptado, o .xlsx temporal convertido por COM.
    Sin prompts visibles.
    """
    passwords = passwords or []
    if path.suffix.lower() == ".csv":
        return path
    try:
        with pd.ExcelFile(path, engine="openpyxl"):
            return path
    except Exception as e1:
        err = str(e1).lower()
        if any(k in err for k in ("password", "encrypt", "badzipfile", "not a zip")):
            for pw in passwords:
                try:
                    bio = decrypt_to_stream(path, pw)
                    with pd.ExcelFile(bio, engine="openpyxl"):
                        pass
                    return bio
                except Exception:
                    continue
        if HAS_COM:
            return com_convert_to_xlsx(path, passwords)
        raise

def find_sheet_name_flexible_pd(src, targets=("INVENTARIO","INVENTARIO GENERAL","INV","Sheet1","Sheet 1","Hoja1")) -> str:
    """Elige la mejor hoja, tolerante a sinónimos; si no, la primera."""
    xf = pd.ExcelFile(src, engine="openpyxl")
    names = xf.sheet_names
    if not names:
        raise ValueError("El libro no tiene hojas.")
    norm_map = {_norm(n): n for n in names}
    for t in targets:
        tn = _norm(t)
        if tn in norm_map:
            return norm_map[tn]
    for t in targets:
        tn = _norm(t)
        for kn, real in norm_map.items():
            if tn in kn:
                return real
    return names[0]

def read_excel_header_at(path: Path, sheet: str | int, header_row_visible: int) -> pd.DataFrame:
    """Lee una hoja con header en 'header_row_visible' (1-based) de forma robusta."""
    src = open_as_excel_source(path, PASSWORDS_TRY)
    hdr_idx0 = header_row_visible - 1
    chosen = find_sheet_name_flexible_pd(src, targets=(sheet, "INVENTARIO", "INVENTARIO GENERAL", "INV", "Sheet1", "Sheet 1", "Hoja1")) \
             if isinstance(sheet, str) else sheet
    df = pd.read_excel(src, sheet_name=chosen, engine="openpyxl", header=hdr_idx0)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")].copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

# ==== LECTURA DE INSUMOS ====
def cargar_inventario_actualizado(base_dir: Path) -> pd.DataFrame:
    """ERP preferido; si no hay, cae en PLANTILLA."""
    # ERP
    try:
        p = find_by_prefix(base_dir, PFX_INV_ACTUALIZADO)
        log(f"Abriendo inventario actualizado (ERP): {p.name}")
        df = read_excel_header_at(p, sheet="Sheet 1", header_row_visible=1)
        idx = {_norm(c): c for c in df.columns}

        ref_col = (
            idx.get("referencia") or idx.get("referencia interna") or idx.get("ref")
            or idx.get("codigo") or idx.get("código")
            or next((real for kn, real in idx.items() if "referenc" in kn or "codigo" in kn or kn.endswith("ref")), None)
        )
        if not ref_col:
            raise KeyError(f"{p.name}: no encuentro columna de Referencia. Encabezados: {list(df.columns)}")

        df = df[~df[ref_col].isna() & (df[ref_col].astype(str).str.strip() != "")].copy()
        df["__REFERENCIA__"] = df[ref_col].apply(to_num_str)

        nom_col     = idx.get("nombre") or "Nombre"
        marca_col   = next((real for kn, real in idx.items()
                            if ("marca/ nombre a mostrar" in kn) or ("marca nombre a mostrar" in kn) or (kn == "marca")), None) \
                      or next((real for kn, real in idx.items() if "marca" in kn and "mostrar" in kn), None)
        linea_col   = next((real for kn, real in idx.items()
                            if ("linea/ nombre a mostrar" in kn) or ("línea/ nombre a mostrar" in kn)), None) \
                      or next((real for kn, real in idx.items() if "linea" in kn and "mostrar" in kn), None)
        sublinea_col = next((real for kn, real in idx.items() if "sub" in kn and "linea" in kn and "mostrar" in kn), None)
        costo_col   = idx.get("costo") or "Costo"

        rename = {}
        if nom_col      in df.columns: rename[nom_col]      = "__NOMBRE__"
        if marca_col    in df.columns: rename[marca_col]    = "__MARCA_SYS__"
        if linea_col    in df.columns: rename[linea_col]    = "__LINEA_SYS__"
        if sublinea_col in df.columns: rename[sublinea_col] = "__SUBLINEA_SYS__"
        if costo_col    in df.columns: rename[costo_col]    = "__COSTO__"
        return df.rename(columns=rename)
    except FileNotFoundError:
        pass

    # PLANTILLA
    p_pl = base_dir / FN_INV_PLANTILLA
    if p_pl.exists():
        p = p_pl
    else:
        for pref in ["$2025 INVENTARIO GENERAL", "2025 INVENTARIO GENERAL", "INVENTARIO GENERAL"]:
            try:
                p = find_by_prefix(base_dir, pref)
                break
            except Exception:
                p = None
        if p is None:
            raise FileNotFoundError(
                f"No encontré ni '{PFX_INV_ACTUALIZADO}' ni una variante de '$2025 INVENTARIO GENERAL' en {base_dir}"
            )

    log(f"[Fallback] Abriendo plantilla de inventario: {p.name}")
    df = read_excel_header_at(p, sheet=SHEET_INV_ORIG, header_row_visible=HEADER_ROW_INV)
    idx = {_norm(c): c for c in df.columns}

    ref_col = (
        idx.get("referencia") or idx.get("referencia fertrac") or idx.get("referencia interna")
        or idx.get("ref") or idx.get("código") or idx.get("codigo")
        or next((real for kn, real in idx.items() if "referenc" in kn or "codigo" in kn or kn.endswith("ref")), None)
    )
    if not ref_col:
        raise KeyError(f"{p.name}: no encuentro columna 'REFERENCIA'. Encabezados: {list(df.columns)}")

    df = df[~df[ref_col].isna() & (df[ref_col].astype(str).str.strip() != "")].copy()
    df["__REFERENCIA__"] = df[ref_col].apply(to_num_str)

    nombre_odoo = idx.get("nombre odoo") or idx.get("nombre")
    marca_sys   = idx.get("marca sistema")
    linea_sys   = idx.get("linea sistema") or idx.get("línea sistema")
    sub_sys     = idx.get("sub- linea sistema") or idx.get("sub-linea sistema") or idx.get("sub linea sistema")
    costo_prom  = idx.get("costo promedio") or idx.get("costo prom")

    rename = {}
    if nombre_odoo in df.columns: rename[nombre_odoo] = "__NOMBRE__"
    if marca_sys   in df.columns: rename[marca_sys]   = "__MARCA_SYS__"
    if linea_sys   in df.columns: rename[linea_sys]   = "__LINEA_SYS__"
    if sub_sys     in df.columns: rename[sub_sys]     = "__SUBLINEA_SYS__"
    if costo_prom  in df.columns: rename[costo_prom]  = "__COSTO__"

    return df.rename(columns=rename)

def cargar_valorizado(base_dir: Path, prefix: str) -> pd.DataFrame:
    """Lee VALORIZADO* (header visible en fila 9)."""
    p = find_by_prefix(base_dir, prefix)
    log(f"Abrir: {p.name}")
    src = open_as_excel_source(p, PASSWORDS_TRY)

    if p.suffix.lower() == ".csv":
        df_all = pd.read_csv(src, header=None, dtype=str)
    else:
        df_all = pd.read_excel(src, sheet_name=0, engine="openpyxl", header=None)

    hdr_row0 = HEADER_ROW_VAL - 1
    if hdr_row0 >= len(df_all):
        raise ValueError(f"{p.name}: HEADER_ROW_VAL={HEADER_ROW_VAL} supera el número de filas.")

    df = df_all.iloc[hdr_row0:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.iloc[0]]
    df = df.iloc[1:].reset_index(drop=True)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

    idx = {_norm(c): c for c in df.columns}
    refc = idx.get("referencia interna") or idx.get("referencia") or idx.get("ref") \
           or next((real for kn, real in idx.items() if "referenc" in kn), None)
    cant = idx.get("cantidad") or next((real for kn, real in idx.items() if kn.startswith("cant")), None)

    if not refc: raise KeyError(f"{p.name}: no encuentro 'Referencia interna'. Encabezados: {list(df.columns)}")
    if not cant: raise KeyError(f"{p.name}: no encuentro 'Cantidad'. Encabezados: {list(df.columns)}")

    out = pd.DataFrame()
    out["__REF_INT__"] = df[refc].apply(to_num_str)
    out["__CANT__"]    = pd.to_numeric(df[cant], errors="coerce").fillna(0.0)
    return out

# ==== EXCEL COM (edición del libro PLANTILLA) ====
def excel_open(path: Path, password: str | None = None):
    """Abre con COM en modo silencioso; si está cifrado, desencripta a temporal y re-aplica password al guardar."""
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.Interactive = False
    excel.EnableEvents = False
    try: excel.AskToUpdateLinks = False
    except Exception: pass
    try: excel.AutomationSecurity = 3
    except Exception: pass
    try: excel.ScreenUpdating = False
    except Exception: pass

    def _force_invisible(app):
        try:
            app.Visible = False
            for i in range(1, app.Windows.Count + 1):
                try:
                    app.Windows(i).Visible = False
                except Exception:
                    pass
        except Exception:
            pass
    _force_invisible(excel)

    info = {"tmp_path": None, "target_path": str(path), "reapply_password": None}

    # Detectar cifrado para no disparar prompts
    encrypted = False
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        try:
            with open(path, "rb") as f:
                of = msoffcrypto.OfficeFile(f)
                encrypted = bool(getattr(of, "is_encrypted", False))
        except Exception:
            encrypted = False

    src_path = path
    if encrypted:
        ok = False
        for pw in PASSWORDS_TRY:
            try:
                bio = decrypt_to_stream(path, pw)
                tmp = save_bytesio_to_temp(bio, Path(path).stem)
                src_path = tmp
                info["tmp_path"] = str(tmp)
                info["reapply_password"] = pw
                ok = True
                break
            except Exception:
                continue
        if not ok:
            excel.Quit()
            raise RuntimeError(f"El libro '{path.name}' está cifrado y ninguna contraseña funcionó.")

    try:
        wb = excel.Workbooks.Open(str(src_path), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True)
        _force_invisible(excel)
        return excel, wb, info
    except Exception as e:
        excel.Quit()
        raise RuntimeError(f"No pude abrir el libro {path.name} de forma silenciosa.") from e

def excel_close(excel, wb, save=True):
    try:
        wb.Close(SaveChanges=save)
    finally:
        excel.Quit()

def ws_headers(ws, header_row_visible: int) -> tuple[dict, dict]:
    """Devuelve (mapa header→col_idx, mapa normalizado→col_idx)"""
    used_cols = ws.UsedRange.Columns.Count
    hdr = {}
    for c in range(1, used_cols+1):
        v = ws.Cells(header_row_visible, c).Value
        if v is None: continue
        s = str(v).strip()
        if s and s != "None":
            hdr[s] = c
    hdrn = {_norm(k): v for k, v in hdr.items()}
    return hdr, hdrn

# ==== AJUSTES PIVOT-SAFE (helpers) ====
def ws_first_pivot_row(ws) -> int | None:
    """Fila superior de la primera PivotTable, o None si no hay."""
    try:
        pts = ws.PivotTables()
        count = int(getattr(pts, "Count", 0))
        if count == 0:
            return None
        first = None
        for i in range(1, count + 1):
            try:
                r = pts(i).TableRange2.Row
                if first is None or r < first:
                    first = r
            except Exception:
                pass
        return first
    except Exception:
        return None

def ws_pivot_blocks(ws):
    """Lista de bloques de pivots [(r1, r2, c1, c2), ...] usando TableRange2."""
    blocks = []
    try:
        pts = ws.PivotTables()
        count = int(getattr(pts, "Count", 0))
        for i in range(1, count + 1):
            try:
                tr = pts(i).TableRange2
                r1, c1 = tr.Row, tr.Column
                r2 = r1 + tr.Rows.Count - 1
                c2 = c1 + tr.Columns.Count - 1
                blocks.append((int(r1), int(r2), int(c1), int(c2)))
            except Exception:
                pass
    except Exception:
        pass
    return blocks

def _ranges_without_pivots_for_column(col_idx: int, start_row: int, end_row: int, pivot_blocks):
    """Devuelve sub-rangos [a,b] dentro de [start_row,end_row] que NO cruzan pivots en la columna dada."""
    if end_row < start_row:
        return []
    holes = []
    for (r1, r2, c1, c2) in pivot_blocks:
        if c1 <= col_idx <= c2:
            holes.append((max(r1, start_row), min(r2, end_row)))
    holes.sort()
    segments = []
    cur = start_row
    for (h1, h2) in holes:
        if h2 < cur or h1 > end_row:
            continue
        if h1 > cur:
            segments.append((cur, h1 - 1))
        cur = max(cur, h2 + 1)
    if cur <= end_row:
        segments.append((cur, end_row))
    return [(a, b) for (a, b) in segments if b >= a]

# ==== OTRAS WS UTILS ====
def ws_last_row(ws, key_col_idx: int, header_row_visible: int):
    """Última fila con datos considerando la columna clave indicada."""
    last = ws.Cells(ws.Rows.Count, key_col_idx).End(-4162).Row  # xlUp
    return max(last, header_row_visible)

def ws_fill_column_values(ws, col_idx: int, start_row: int, values: list):
    """
    Escribe valores en una columna saltando los bloques de PivotTable.
    Si parte del rango cae dentro de una pivot, esa porción se omite.
    """
    n = len(values)
    if n == 0:
        return

    end_row = start_row + n - 1
    pivots = ws_pivot_blocks(ws)
    safe_segments = _ranges_without_pivots_for_column(col_idx, start_row, end_row, pivots)

    offset = 0  # índice en 'values'
    for (a, b) in safe_segments:
        if offset >= n:
            break
        seg_len = min(b - a + 1, n - offset)
        if seg_len <= 0:
            continue

        # Normaliza Nones/NaN a cadena vacía para evitar errores
        chunk = values[offset: offset + seg_len]
        chunk = [("" if (v is None or (isinstance(v, float) and np.isnan(v))) else v) for v in chunk]

        rng = ws.Range(ws.Cells(a, col_idx), ws.Cells(a + seg_len - 1, col_idx))
        rng.Value = [[v] for v in chunk]
        offset += seg_len


def ws_clear_column(ws, col_idx: int, start_row: int, end_row: int):
    """
    Limpia una columna por tramos, evitando los bloques de cualquier PivotTable.
    # AJUSTE: ahora es pivot-safe
    """
    if end_row < start_row:
        return
    pivots = ws_pivot_blocks(ws)
    safe_segments = _ranges_without_pivots_for_column(col_idx, start_row, end_row, pivots)
    for (a, b) in safe_segments:
        rng = ws.Range(ws.Cells(a, col_idx), ws.Cells(b, col_idx))
        rng.ClearContents()

def ws_copy_down_formula(ws, col_idx: int, start_row: int, end_row: int):
    """Copia la fórmula desde start_row hasta end_row (si hay)."""
    if end_row < start_row: return
    fml = ws.Cells(start_row, col_idx).Formula
    if fml:
        ws.Range(ws.Cells(start_row, col_idx), ws.Cells(end_row, col_idx)).Formula = fml

def ws_headers_smart(ws, expected_row: int, preferred_labels: list[str] | None = None):
    """
    Detecta de forma robusta la fila de encabezado (expected_row y luego 1..10),
    y valida alguna etiqueta preferida si se suministra.
    Devuelve: (header_row_used, hdr_map, hdr_norm_map)
    """
    preferred_norm = [_norm(x) for x in (preferred_labels or [])]
    tried = [expected_row] + [r for r in range(1, 11) if r != expected_row]
    for hr in tried:
        hdr, hdrn = ws_headers(ws, hr)
        if not hdrn:
            continue
        if not preferred_norm or any(lbl in hdrn for lbl in preferred_norm):
            return hr, hdr, hdrn
    # Fallback: primera fila del UsedRange
    try:
        first_row = ws.UsedRange.Row
        hdr, hdrn = ws_headers(ws, first_row)
        if hdrn:
            return first_row, hdr, hdrn
    except Exception:
        pass
    return expected_row, {}, {}

def find_reference_col_idx(hdrn: dict, ws, header_row_used: int) -> int:
    """
    Encuentra índice de columna para REFERENCIA (o sinónimos).
    Si no halla, devuelve la primera columna con datos debajo del header.
    """
    for name in ["REFERENCIA", "REFERENCIA FERTRAC", "REFERENCIA INTERNA", "REF", "CÓDIGO", "CODIGO", "SKU"]:
        cidx = hdrn.get(_norm(name))
        if cidx:
            return cidx
    for k, v in hdrn.items():
        if "referenc" in k or "codigo" in k or k.endswith("ref"):
            return v
    used_cols = ws.UsedRange.Columns.Count
    for c in range(1, used_cols + 1):
        for r in range(header_row_used + 1, header_row_used + 15):
            val = ws.Cells(r, c).Value
            if val not in (None, "", "None"):
                return c
    return 1

def ws_ensure_existencia_header(ws, header_row_visible: int) -> int:
    """Devuelve col_idx del encabezado EXISTENCIA {MES DD}; si existe uno previo, lo renombra."""
    target = exist_col_title_for_today()
    hdr, hdrn = ws_headers(ws, header_row_visible)
    target_col = None
    for name, col in hdr.items():
        if _norm(name).startswith("existencia "):
            target_col = col
            ws.Cells(header_row_visible, target_col).Value = target
            break
    if target_col is None:
        used_cols = ws.UsedRange.Columns.Count
        target_col = used_cols + 1
        ws.Cells(header_row_visible, target_col).Value = target
    return target_col

# ==== PROCESO ====
def main():
    if not HAS_COM:
        raise RuntimeError("Este script requiere Excel COM (win32com). Instálalo y ejecuta en Windows con Excel.")

    log("== Inicio actualización de inventario ==")

    # 1) Inventario actualizado (ERP o plantilla)
    df_src = cargar_inventario_actualizado(BASE_PATH)

    # 14–17) Valorizados
    df_val_gen   = cargar_valorizado(BASE_PATH, PFX_VAL_GENERAL)
    df_val_impo  = cargar_valorizado(BASE_PATH, PFX_VAL_FALT_IMPO)
    df_val_falt  = cargar_valorizado(BASE_PATH, PFX_VAL_FALT)
    df_val_tob   = cargar_valorizado(BASE_PATH, PFX_VAL_TOBERIN)

    # Join de cantidades (por referencia interna)
    val_map_impo = df_val_impo.set_index("__REF_INT__")["__CANT__"]
    val_map_falt = df_val_falt.set_index("__REF_INT__")["__CANT__"]
    val_map_tob  = df_val_tob.set_index("__REF_INT__")["__CANT__"]

    # En VALORIZADO GENERAL agregamos columnas calculadas (solo en DF)
    df_val_gen = df_val_gen.copy()
    df_val_gen["__IMPO_MATCH__"] = df_val_gen["__REF_INT__"].isin(val_map_impo.index)
    df_val_gen["__IMPO_CANT__"]  = df_val_gen["__REF_INT__"].map(val_map_impo).fillna(0.0)
    df_val_gen["__IMPO_DIF__"]   = df_val_gen["__CANT__"] - df_val_gen["__IMPO_CANT__"]

    df_val_gen["__FALT_MATCH__"] = df_val_gen["__REF_INT__"].isin(val_map_falt.index)
    df_val_gen["__FALT_CANT__"]  = df_val_gen["__REF_INT__"].map(val_map_falt).fillna(0.0)
    df_val_gen["__FALT_DIF__"]   = df_val_gen["__CANT__"] - df_val_gen["__FALT_CANT__"]

    df_val_gen["__TOB_MATCH__"]  = df_val_gen["__REF_INT__"].isin(val_map_tob.index)
    df_val_gen["__TOB_CANT__"]   = df_val_gen["__REF_INT__"].map(val_map_tob).fillna(0.0)
    df_val_gen["__TOB_DIF__"]    = df_val_gen["__CANT__"] - df_val_gen["__TOB_CANT__"]

    # 18) Consolidado EXISTENCIA_CALC
    df_val_gen["__EXIST_CALC__"] = (
        df_val_gen["__IMPO_CANT__"] + df_val_gen["__FALT_CANT__"] + df_val_gen["__TOB_CANT__"]
    )
    exist_map = df_val_gen.set_index("__REF_INT__")["__EXIST_CALC__"]

    # 2) Abrir libro PLANTILLA protegido con COM
    p_inv = BASE_PATH / FN_INV_PLANTILLA
    log(f"Abrir libro: {p_inv}")
    excel, wb, saveinfo = excel_open(p_inv, password=PASS_INV)

    try:
        ws_inv = wb.Worksheets(SHEET_INV_ORIG)
    except Exception:
        ws_inv = wb.Worksheets(1)

    # 2a) Crear copia de la hoja INVENTARIO
    log("Creando copia de la hoja INVENTARIO…")
    try:
        excel.DisplayAlerts = False
        for i in range(1, wb.Worksheets.Count + 1):
            try:
                if wb.Worksheets(i).Name.strip() == "INVENTARIO (COPIA)":
                    wb.Worksheets(i).Delete()
                    break
            except:
                pass
    except Exception as e:
        log(f"Error al eliminar hoja existente: {e}")
    was_protected = False
    try:
        if ws_inv.ProtectContents or ws_inv.ProtectDrawingObjects or ws_inv.ProtectScenarios:
            was_protected = True
            ws_inv.Unprotect(Password=PASS_INV)

    except Exception as e:
        log(f"Aviso al desproteger: {e}")
    try:
        ws_inv_copia = wb.Worksheets.Add(After=ws_inv)
        ws_inv_copia.Name = "INVENTARIO (COPIA)"
        ws_inv.UsedRange.Copy(Destination=ws_inv_copia.Range("A1"))
        try:
            for col in range(1, ws_inv.UsedRange.Columns.Count + 1):
                ws_inv_copia.Columns(col).ColumnWidth = ws_inv.Columns(col).ColumnWidth
        except Exception as e:
            log(f"Aviso: no se pudo copiar anchos de columna: {e}")
        
    except Exception as e:
        log(f"ERROR al crear copia manual: {e}")
        raise RuntimeError(f"No se pudo crear copia de la hoja INVENTARIO: {e}")
    if was_protected:
        try:
            ws_inv.Protect(Password=PASS_INV, DrawingObjects=True, Contents=True, Scenarios=True)
            log("Hoja INVENTARIO re-protegida")
        except Exception as e:
            log(f"Aviso al re-proteger: {e}")


    # 2b) Detectar encabezado robusto en INVENTARIO
    header_row_used, hdr_inv, hdrn_inv = ws_headers_smart(
        ws_inv,
        expected_row=HEADER_ROW_INV,
        preferred_labels=["REFERENCIA", "REFERENCIA FERTRAC"]
    )

    # 2c) Limpiar columnas solicitadas en hoja INVENTARIO
    log(f"Encabezados detectados en fila {header_row_used}. Limpiando columnas en INVENTARIO…")
    ref_col_idx = find_reference_col_idx(hdrn_inv, ws_inv, header_row_used)
    last_row = ws_last_row(ws_inv, ref_col_idx, header_row_used)
    start_data_row = header_row_used + 1

    # AJUSTE: recorte defensivo antes de la primera Pivot
    pivot_top = ws_first_pivot_row(ws_inv)
    if pivot_top and pivot_top > header_row_used:
        last_row = min(last_row, pivot_top - 1)

    for colname in COLS_A_LIMPIAR:
        cidx = hdrn_inv.get(_norm(colname))
        if cidx:
            ws_clear_column(ws_inv, cidx, start_data_row, last_row)

    # 2d) INV LISTA PRECIOS: borrar REFERENCIA FERTRAC (header esperado en fila 1, pero robusto)
    log("Limpiando REFERENCIA FERTRAC en INV LISTA PRECIOS…")
    try:
        # Buscar la hoja de forma flexible
        ws_lp = None
        target_norm = _norm(SHEET_INV_LISTA)
        
        for i in range(1, wb.Worksheets.Count + 1):
            sheet_name = wb.Worksheets(i).Name
            if _norm(sheet_name) == target_norm or target_norm in _norm(sheet_name):
                ws_lp = wb.Worksheets(i)
                log(f"Hoja encontrada: '{sheet_name}'")
                break
        
        if ws_lp is None:
            # Intento alternativo: buscar por palabras clave
            for i in range(1, wb.Worksheets.Count + 1):
                sheet_name_norm = _norm(wb.Worksheets(i).Name)
                if "inv" in sheet_name_norm and "lista" in sheet_name_norm and "precio" in sheet_name_norm:
                    ws_lp = wb.Worksheets(i)
                    log(f"Hoja encontrada (por palabras clave): '{wb.Worksheets(i).Name}'")
                    break
        
        if ws_lp:
            hr_lp, hdr_lp, hdrn_lp = ws_headers_smart(ws_lp, HEADER_ROW_INV_LISTA, ["REFERENCIA FERTRAC"])
            cidx = hdrn_lp.get(_norm("REFERENCIA FERTRAC"))
            if cidx:
                last_row_lp = ws_last_row(ws_lp, cidx, hr_lp)
                # AJUSTE: recorte defensivo por Pivot en hoja LP
                pivot_top_lp = ws_first_pivot_row(ws_lp)
                if pivot_top_lp and pivot_top_lp > hr_lp:
                    last_row_lp = min(last_row_lp, pivot_top_lp - 1)
                ws_clear_column(ws_lp, cidx, hr_lp + 1, last_row_lp)
                log("REFERENCIA FERTRAC limpiada exitosamente")
            else:
                log("Columna REFERENCIA FERTRAC no encontrada en la hoja")
        else:
            raise Exception("Hoja no encontrada después de búsqueda flexible")
                
    except Exception as e:
        log(f"No se pudo procesar 'INV LISTA PRECIOS': {e}")

    # 3–7, 13) Pegar columnas desde INVENTARIO ACTUALIZADO (ERP) hacia INVENTARIO
    log("Pegando columnas desde Inventario actualizado…")
    ref_values   = df_src["__REFERENCIA__"].tolist()
    nombre_odoo  = df_src.get("__NOMBRE__",       pd.Series([""]*len(ref_values))).tolist()
    marca_sys    = df_src.get("__MARCA_SYS__",    pd.Series([""]*len(ref_values))).tolist()
    linea_sys    = df_src.get("__LINEA_SYS__",    pd.Series([""]*len(ref_values))).tolist()
    sublinea_sys = df_src.get("__SUBLINEA_SYS__", pd.Series([""]*len(ref_values))).tolist()
    costo_prom   = df_src.get("__COSTO__",        pd.Series([np.nan]*len(ref_values))).tolist()

    def paste_if_exists(col_name, values, number_format=None):
        cidx = hdrn_inv.get(_norm(col_name))
        if not cidx:
            return
        ws_fill_column_values(ws_inv, cidx, start_data_row, values)
        if number_format:
            ws_inv.Columns(cidx).NumberFormat = number_format

    paste_if_exists("REFERENCIA", ref_values, number_format="0")
    paste_if_exists("NOMBRE ODOO", nombre_odoo)
    paste_if_exists("Marca sistema", marca_sys)
    paste_if_exists("Linea sistema", linea_sys)
    paste_if_exists("Sub- linea sistema", sublinea_sys)
    paste_if_exists("COSTO PROMEDIO", costo_prom)

    last_row = max(last_row, start_data_row + len(ref_values) - 1)

    # 8) Arrastrar fórmulas: Dif marca, Dif linea, Dif sub-linea
    for colname in ["Dif marca", "Dif linea", "Dif sub-linea"]:
        cidx = hdrn_inv.get(_norm(colname))
        if cidx:
            ws_copy_down_formula(ws_inv, cidx, start_data_row, last_row)

    # 9) TOTAL INV + EXISTENCIA {MES DD}
    col_total_inv = hdrn_inv.get(_norm("TOTAL INV"))
    if col_total_inv:
        ws_copy_down_formula(ws_inv, col_total_inv, start_data_row, last_row)

    col_exist = ws_ensure_existencia_header(ws_inv, header_row_used)
    ws_copy_down_formula(ws_inv, col_exist, start_data_row, last_row)

    # 10–11) Arrastrar fórmulas de NOMBRE LISTA y NOMBRE MYR
    for colname in ["NOMBRE LISTA", "NOMBRE MYR"]:
        cidx = hdrn_inv.get(_norm(colname))
        if cidx:
            ws_copy_down_formula(ws_inv, cidx, start_data_row, last_row)

    # 12) Traer columnas desde hoja copia (relación por REFERENCIA)
    log("Trayendo columnas desde hoja 'INVENTARIO (COPIA)' por REFERENCIA…")
    try:
        ws_cp = wb.Worksheets("INVENTARIO (COPIA)")
        hr_cp, hdr_cp, hdrn_cp = ws_headers_smart(ws_cp, HEADER_ROW_INV, ["REFERENCIA"])
        ref_idx_cp = hdrn_cp.get(_norm("REFERENCIA")) or find_reference_col_idx(hdrn_cp, ws_cp, hr_cp)
        if ref_idx_cp:
            last_cp = ws_last_row(ws_cp, ref_idx_cp, hr_cp)
            refs_cp = [
                str(ws_cp.Cells(r, ref_idx_cp).Value).strip() if ws_cp.Cells(r, ref_idx_cp).Value is not None else ""
                for r in range(hr_cp + 1, last_cp + 1)
            ]
            maps = {}
            for colname in COLS_DESDE_COPIA:
                cidx = hdrn_cp.get(_norm(colname))
                if not cidx:
                    continue
                vals = [ws_cp.Cells(r, cidx).Value for r in range(hr_cp + 1, last_cp + 1)]
                maps[colname] = dict(zip(refs_cp, vals))

            ref_idx_inv = ref_col_idx
            for r in range(start_data_row, last_row + 1):
                ref = str(ws_inv.Cells(r, ref_idx_inv).Value).strip() if ws_inv.Cells(r, ref_idx_inv).Value is not None else ""
                if not ref:
                    continue
                for colname in COLS_DESDE_COPIA:
                    tgt_idx = hdrn_inv.get(_norm(colname))
                    if not tgt_idx or colname not in maps:
                        continue
                    val = maps[colname].get(ref, None)
                    if val is not None:
                        ws_inv.Cells(r, tgt_idx).Value = val
    except Exception as e:
        log(f"Aviso: no se pudo completar paso 12: {e}")

    # 18) Llevar EXISTENCIA_CALC (desde valorizados) → EXISTENCIA {MES DD}
    log("Escribiendo EXISTENCIA consolidada desde VALORIZADOS…")
    ref_idx_inv = ref_col_idx
    if ref_idx_inv:
        for r in range(start_data_row, last_row + 1):
            raw = ws_inv.Cells(r, ref_idx_inv).Value
            key = to_num_str(raw)
            if not key:
                continue
            val = exist_map.get(key)
            if pd.notna(val):
                ws_inv.Cells(r, col_exist).Value = float(val)

    # 19) Ajustes con filtros (LINEA COPIA == #N/D y Línea sistema en set dado)
    log("Aplicando ajustes básicos de #N/D para líneas seleccionadas…")
    marcas_ok = {"clevite", "cummins", "dana", "fersa", "meritor", "zf"}
    idx_linea_copia = hdrn_inv.get(_norm("LINEA COPIA"))
    idx_linea_sys   = hdrn_inv.get(_norm("Linea sistema"))
    idx_marca_copia = hdrn_inv.get(_norm("MARCA copia"))
    idx_marca_sys   = hdrn_inv.get(_norm("Marca sistema"))
    idx_sub_copia   = hdrn_inv.get(_norm("SUB-LINEA COPIA"))
    idx_sub_sys     = hdrn_inv.get(_norm("Sub- linea sistema"))

    if idx_linea_copia and idx_linea_sys:
        for r in range(start_data_row, last_row + 1):
            val_linea_copia = str(ws_inv.Cells(r, idx_linea_copia).Value).strip() if ws_inv.Cells(r, idx_linea_copia).Value is not None else ""
            val_linea_sys   = str(ws_inv.Cells(r, idx_linea_sys).Value).strip()   if ws_inv.Cells(r, idx_linea_sys).Value   is not None else ""
            val_marca_sys   = str(ws_inv.Cells(r, idx_marca_sys).Value).strip()   if (idx_marca_sys and ws_inv.Cells(r, idx_marca_sys).Value is not None) else ""
            if val_linea_copia in ("#N/D", "#N/A", "N/D", "ND", ""):
                if _norm(val_linea_sys) and any(m in _norm(val_linea_sys) for m in marcas_ok):
                    if idx_marca_copia and idx_marca_sys:
                        ws_inv.Cells(r, idx_marca_copia).Value = val_marca_sys
                    if idx_linea_copia:
                        ws_inv.Cells(r, idx_linea_copia).Value = val_linea_sys
                    if idx_sub_copia and idx_sub_sys:
                        ws_inv.Cells(r, idx_sub_copia).Value = (
                            str(ws_inv.Cells(r, idx_sub_sys).Value) if ws_inv.Cells(r, idx_sub_sys).Value is not None else ""
                        )

   # === ELIMINAR HOJA TEMPORAL ===
    log("Eliminando hoja temporal INVENTARIO (COPIA)...")
    try:
        ws_copia_name = ws_inv_copia.Name
        # Asegurarse de que no esté activa la hoja que vamos a borrar
        ws_inv.Activate()
        # Desactivar alertas para evitar confirmación de eliminación
        excel.DisplayAlerts = False
        wb.Worksheets(ws_copia_name).Delete()
        excel.DisplayAlerts = False  # mantener desactivadas
        log("Hoja temporal eliminada")
    except Exception as e:
        log(f"Aviso: no se pudo eliminar hoja temporal: {e}")    


    # === LLENAR REFERENCIA FERTRAC EN INV LISTA PRECIOS ===
    log("Llenando REFERENCIA FERTRAC en INV LISTA PRECIOS desde INVENTARIO...")
    try:
        ws_lp = None
        target_norm = _norm(SHEET_INV_LISTA)
        
        for i in range(1, wb.Worksheets.Count + 1):
            sheet_name = wb.Worksheets(i).Name
            if _norm(sheet_name) == target_norm or target_norm in _norm(sheet_name):
                ws_lp = wb.Worksheets(i)
                break
        
        if ws_lp is None:
            for i in range(1, wb.Worksheets.Count + 1):
                sheet_name_norm = _norm(wb.Worksheets(i).Name)
                if "inv" in sheet_name_norm and "lista" in sheet_name_norm and "precio" in sheet_name_norm:
                    ws_lp = wb.Worksheets(i)
                    break
        
        if ws_lp:
            hr_lp, hdr_lp, hdrn_lp = ws_headers_smart(ws_lp, HEADER_ROW_INV_LISTA, ["REFERENCIA FERTRAC"])
            ref_fertrac_idx = hdrn_lp.get(_norm("REFERENCIA FERTRAC"))
            
            if ref_fertrac_idx:
                ref_idx_inv = ref_col_idx  # columna REFERENCIA de INVENTARIO
                referencias_inv = []
                
                for r in range(start_data_row, last_row + 1):
                    val = ws_inv.Cells(r, ref_idx_inv).Value
                    if val is not None and str(val).strip():
                        referencias_inv.append(val)
                last_row_lp = hr_lp + len(referencias_inv)
                ws_fill_column_values(ws_lp, ref_fertrac_idx, hr_lp + 1, referencias_inv)
                try:
                    rng = ws_lp.Range(ws_lp.Cells(hr_lp + 1, ref_fertrac_idx), 
                                     ws_lp.Cells(last_row_lp, ref_fertrac_idx))
                    rng.NumberFormat = "0"
                except Exception as e:
                    log(f"Aviso: no se pudo aplicar formato numérico: {e}")
                
                log(f"{len(referencias_inv)} referencias copiadas a REFERENCIA FERTRAC")
            else:
                log("No se encontró columna REFERENCIA FERTRAC en INV LISTA PRECIOS")
        else:
            log("No se encontró la hoja INV LISTA PRECIOS")
            
    except Exception as e:
        log(f"Error al llenar REFERENCIA FERTRAC: {e}")  

    # === GUARDADO COMO ARCHIVO NUEVO ===
    log("Preparando guardado...")

    # 1) Asegurar que haya una hoja visible y activa
    try:
        ws_count = int(wb.Worksheets.Count)
        has_visible = False
        for i in range(1, ws_count + 1):
            try:
                if int(wb.Worksheets(i).Visible) == -1:  # xlSheetVisible
                    has_visible = True
                    break
            except Exception:
                pass
        if not has_visible and ws_count >= 1:
            wb.Worksheets(1).Visible = -1  # visible
            wb.Worksheets(1).Activate()
    except Exception:
        pass

    # 2) Evitar que el libro se guarde como Add-In oculto
    with contextlib.suppress(Exception):
        wb.IsAddin = False
    with contextlib.suppress(Exception):
        # que la ventana del libro esté marcada visible (no persiste siempre, pero ayuda)
        wb.Windows(1).Visible = True

    # 3) Nombre de salida con timestamp
    try:
        base_name = OUTPUT_BASENAME
    except NameError:
        base_name = f"{Path(FN_INV_PLANTILLA).stem} (ACTUALIZADO)"
    out_name = f"{base_name} {datetime.now():%Y%m%d_%H%M}.xlsx"
    out_path = BASE_PATH / out_name

    # 4) Decide si aplicar contraseña
    #    - si el original estaba cifrado reusamos esa
    #    - o aplica PASS_INV si quieres; para testear, pon apply_pw = None
    apply_pw = saveinfo.get("reapply_password")
    # >>> si quieres probar SIN contraseña para descartar el problema, descomenta la línea siguiente:
    # apply_pw = None
    if apply_pw:
        wb.SaveAs(str(out_path), FileFormat=51, Password=apply_pw)
    else:
        wb.SaveAs(str(out_path), FileFormat=51)

    log(f"✅ Archivo NUEVO creado: {out_path}")

    # Cerrar sin re-guardar el original
    excel_close(excel, wb, save=False)

    # limpiar temporal si existe
    tmp = saveinfo.get("tmp_path")
    if tmp and os.path.exists(tmp):
        with contextlib.suppress(Exception):
            os.remove(tmp)



if __name__ == "__main__":
    main()
