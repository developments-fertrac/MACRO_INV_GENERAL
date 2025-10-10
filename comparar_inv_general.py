# comparar_inv_general_gui.py
# -*- coding: utf-8 -*-

from __future__ import annotations
from pathlib import Path
import io, re, contextlib, sys, tempfile, os
from datetime import datetime
import pandas as pd
import numpy as np

# GUI: 2 diálogos, forzados topmost
import tkinter as tk
from tkinter import filedialog

# ====== CONFIG ======
SHEET_CANDIDATES = ["INVENTARIO", "INVENTARIO GENERAL", "INV", "Sheet1", "Sheet 1", "Hoja1"]
SEARCH_HEADER_MAXROW = 10
COLUMNAS_A_COMPARAR = None    # p.ej.: ["NOMBRE ODOO","Marca sistema","COSTO PROMEDIO"]; None = todas comunes
CAMPOS_IGNORAR = {"Unnamed: 0"}

PASSWORDS_TRY = ["Compras2025"]  # agrega más si usas otras
ALLOW_COM_CONVERSION = True      # usa Excel COM para .xls / casos raros

VERBOSE = True
def log(msg):
    if VERBOSE:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

# COM opcional (para .xls)
try:
    import win32com.client as win32
    HAS_COM = True
except Exception:
    HAS_COM = False

# msoffcrypto para desencriptar xlsx/xlsm si fuese necesario
try:
    import msoffcrypto
    HAS_CRYPTO = True
except Exception:
    HAS_CRYPTO = False

# ====== UTILIDADES ======
def _norm(s: str) -> str:
    try:
        import unidecode
        t = unidecode.unidecode(str(s))
    except Exception:
        t = str(s)
    t = t.lower()
    t = re.sub(r"[^a-z0-9 ]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t

def _make_unique_columns(cols):
    """Convierte nombres duplicados en únicos: 'Col' -> 'Col', 'Col__2', 'Col__3', ..."""
    out, seen = [], {}
    for c in map(lambda x: str(x).strip(), cols):
        if c not in seen:
            seen[c] = 1
            out.append(c)
        else:
            seen[c] += 1
            out.append(f"{c}__{seen[c]}")
    return out

def to_num_str(x):
    if pd.isna(x): return ""
    s = str(x).strip()
    with contextlib.suppress(Exception):
        f = float(s.replace(",", ""))
        if abs(f - int(f)) < 1e-9:
            return str(int(f))
        return str(f)
    return s

def find_header_row(df: pd.DataFrame) -> int | None:
    for r in range(min(SEARCH_HEADER_MAXROW, len(df))):
        for v in df.iloc[r].tolist():
            if _norm(v) in {"referencia","referencia interna","ref","codigo","código","sku"}:
                return r
    return None

def pick_sheet(xls: pd.ExcelFile) -> str:
    names = xls.sheet_names
    nmap = {_norm(s): s for s in names}
    for cand in SHEET_CANDIDATES:
        key = _norm(cand)
        if key in nmap:
            return nmap[key]
    for cand in SHEET_CANDIDATES:
        for nm in names:
            if _norm(cand) in _norm(nm):
                return nm
    return names[0]

def decrypt_to_stream(xlsx_path: Path, password: str) -> io.BytesIO:
    bio = io.BytesIO()
    with open(xlsx_path, "rb") as f:
        of = msoffcrypto.OfficeFile(f)
        of.load_key(password=password)
        of.decrypt(bio)
    bio.seek(0)
    return bio

def com_convert_to_xlsx(path: Path) -> Path:
    log(f"COM: convirtiendo '{path.name}' → .xlsx temporal")
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try: excel.AskToUpdateLinks = False
    except Exception: pass
    try: excel.AutomationSecurity = 3
    except Exception: pass

    tmp = Path(tempfile.gettempdir()) / f"~conv_{path.stem}_{datetime.now():%H%M%S}.xlsx"
    wb = None
    try:
        wb = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True)
        wb.SaveAs(str(tmp), FileFormat=51)  # .xlsx
    finally:
        try:
            if wb: wb.Close(SaveChanges=False)
        finally:
            excel.Quit()
    return tmp

def open_as_excel_source(path: Path):
    suf = path.suffix.lower()
    if suf == ".csv":
        log(f"Leyendo CSV: {path.name}")
        return path

    # intenta openpyxl directo
    try:
        log(f"Intento openpyxl: {path.name}")
        with pd.ExcelFile(path, engine="openpyxl"):
            return path
    except Exception as e1:
        err = str(e1).lower()
        log(f"openpyxl falló: {e1}")

        # cifrado / zip roto → msoffcrypto
        if HAS_CRYPTO and any(k in err for k in ("password", "encrypt", "badzipfile", "not a zip")):
            for pw in PASSWORDS_TRY:
                try:
                    log(f"Intento desencriptar con clave '{pw}'…")
                    bio = decrypt_to_stream(path, pw)
                    with pd.ExcelFile(bio, engine="openpyxl"):
                        pass
                    log("Desencriptado OK")
                    return bio
                except Exception as e2:
                    log(f"Clave no válida: {e2}")

        # .xls / casos raros → COM
        if ALLOW_COM_CONVERSION and HAS_COM:
            try:
                return com_convert_to_xlsx(path)
            except Exception as e3:
                log(f"COM conversión falló: {e3}")

        # sin salida posible
        raise

def read_inventory(path: Path) -> pd.DataFrame:
    src = open_as_excel_source(path)

    if isinstance(src, Path) and src.suffix.lower() == ".csv":
        raw = pd.read_csv(src, header=None, dtype=object, encoding="utf-8", engine="python")
        hdr = find_header_row(raw)
        if hdr is None: hdr = 0
        df = pd.read_csv(src, header=hdr, dtype=object, encoding="utf-8", engine="python")
    else:
        x = pd.ExcelFile(src, engine="openpyxl")
        sh = pick_sheet(x)
        raw = pd.read_excel(src, sheet_name=sh, engine="openpyxl", header=None, dtype=object)
        hdr = find_header_row(raw)
        if hdr is None: hdr = 0
        df = pd.read_excel(src, sheet_name=sh, engine="openpyxl", header=hdr, dtype=object)

    # limpiar y hacer únicos los encabezados
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")].copy()
    df.columns = _make_unique_columns(df.columns)

    # localizar REFERENCIA
    idx = {_norm(c): c for c in df.columns}
    ref_col = (
        idx.get("referencia") or idx.get("referencia interna") or idx.get("ref")
        or idx.get("codigo") or idx.get("código") or idx.get("sku")
        or next((real for kn, real in idx.items() if "referenc" in kn or "codigo" in kn or kn.endswith("ref")), None)
    )
    if not ref_col:
        raise RuntimeError(f"{Path(path).name}: no se encontró columna 'REFERENCIA'.")

    # filtrar vacíos y normalizar clave
    df = df[~df[ref_col].isna() & (df[ref_col].astype(str).str.strip() != "")].copy()
    df["__REFERENCIA__"] = df[ref_col].apply(to_num_str)

    # quitar duplicados de referencia (conserva la última)
    dup_count = int(df["__REFERENCIA__"].duplicated().sum())
    if dup_count:
        log(f"[WARN] {Path(path).name}: {dup_count} referencias duplicadas -> se conserva la última.")
    df = df.drop_duplicates(subset="__REFERENCIA__", keep="last")

    keep_cols = [c for c in df.columns if c not in CAMPOS_IGNORAR]
    return df[keep_cols].copy()

def norm_for_compare(v):
    # aplana si llega una Serie (pasa cuando aún hay columnas repetidas en origen)
    if isinstance(v, pd.Series):
        v = v.iloc[0] if not v.empty else ""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    if isinstance(v, str):
        return v.strip()
    return v

def comparar_en_una_hoja(path_manual: Path, path_auto: Path, out_path: Path,
                         columnas_objetivo: list[str] | None = None):
    log(f"Leyendo MANUAL: {path_manual}")
    df_m = read_inventory(path_manual)
    log(f"Leyendo AUTO  : {path_auto}")
    df_a = read_inventory(path_auto)

    key = "__REFERENCIA__"
    cols_m = set(df_m.columns) - {key}
    cols_a = set(df_a.columns) - {key}
    if columnas_objetivo:
        comunes = [c for c in columnas_objetivo if (c in cols_m and c in cols_a)]
    else:
        comunes = sorted(list(cols_m.intersection(cols_a)))

    # asegurar índice sin duplicados
    dm = df_m.drop_duplicates(key, keep="last").set_index(key, drop=False)
    da = df_a.drop_duplicates(key, keep="last").set_index(key, drop=False)
    refs_inter = sorted(set(dm.index).intersection(set(da.index)))

    log(f"Comparando {len(refs_inter)} referencias comunes…")
    diffs = []
    for ref in refs_inter:
        row_m = dm.loc[ref]
        row_a = da.loc[ref]
        for col in comunes:
            vm = norm_for_compare(row_m.get(col, ""))
            va = norm_for_compare(row_a.get(col, ""))

            iguales = (vm == va)
            if not iguales:
                # comparación numérica tolerante
                try:
                    fm = float(str(vm).replace(",", ""))
                    fa = float(str(va).replace(",", ""))
                    if abs(fm - fa) < 1e-9:
                        iguales = True
                except Exception:
                    pass

            if not iguales:
                diffs.append({
                    "REFERENCIA": ref,
                    "COLUMNA": col,
                    "VALOR_MANUAL": vm,
                    "VALOR_AUTO": va
                })

    df_diffs = pd.DataFrame(diffs, columns=["REFERENCIA","COLUMNA","VALOR_MANUAL","VALOR_AUTO"])

    out_path.parent.mkdir(parents=True, exist_ok=True)
    log(f"Escribiendo resultado: {out_path}")
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        sheet = "Diferencias"
        df_diffs.to_excel(w, index=False, sheet_name=sheet)
        ws = w.sheets[sheet]
        for j, col in enumerate(df_diffs.columns, start=1):
            try:
                maxlen = max(10, min(60, int(df_diffs[col].astype(str).map(len).fillna(0).max()) + 2))
                ws.column_dimensions[ws.cell(row=1, column=j).column_letter].width = maxlen
            except Exception:
                pass
        try:
            ws.auto_filter.ref = ws.dimensions
        except Exception:
            pass
    log("✅ Listo.")

def main():
    # raíz oculta y diálogos siempre al frente
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    root.update()

    # 1) MANUAL
    path_manual = filedialog.askopenfilename(
        title="Selecciona el INVENTARIO MANUAL",
        filetypes=[("Excel / CSV", "*.xlsx;*.xlsm;*.xls;*.csv"), ("Todos", "*.*")],
        parent=root
    )
    if not path_manual:
        sys.exit(0)

    # fuerza topmost otra vez por si Windows lo pierde
    root.attributes('-topmost', True)
    root.update()

    # 2) AUTO
    path_auto = filedialog.askopenfilename(
        title="Selecciona el INVENTARIO ACTUALIZADO (código)",
        filetypes=[("Excel / CSV", "*.xlsx;*.xlsm;*.xls;*.csv"), ("Todos", "*.*")],
        parent=root
    )
    if not path_auto:
        sys.exit(0)

    out_name = f"DIFERENCIAS INVENTARIO {datetime.now():%Y%m%d_%H%M}.xlsx"
    out_path = Path(path_manual).parent / out_name

    try:
        comparar_en_una_hoja(Path(path_manual), Path(path_auto), out_path,
                             columnas_objetivo=COLUMNAS_A_COMPARAR)
        print(f"\nArchivo generado:\n{out_path}\n")
    except Exception as e:
        print(f"\nERROR: {e}\n")
        raise

if __name__ == "__main__":
    main()
