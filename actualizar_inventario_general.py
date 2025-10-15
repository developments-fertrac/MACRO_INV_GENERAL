# actualizar_inventario_integral_optimizado.py
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
BASE_PATH = Path(__file__).resolve().parent  

PASS_INV = "Compras2025"
PASSWORDS_TRY = ["Compras2025", "Compras2026"]

OUTPUT_BASENAME = "$2025 INVENTARIO GENERAL ACTUALIZADO"
APPLY_PASSWORD_TO_OUTPUT = True

# Prefijos para ubicar archivos descargados del ERP
PFX_INV_ACTUALIZADO = "INVENTARIO GENERAL ACTUALIZADO"
PFX_VAL_GENERAL     = "VALORIZADO GENERAL"
PFX_VAL_FALT_IMPO   = "VALORIZADO FALTANTES IMPO"
PFX_VAL_FALT        = "VALORIZADO FALTANTES"
PFX_VAL_TOBERIN     = "VALORIZADO TOBERIN"
PFX_MARCAS          = "MARCAS"
PFX_DISTRIBUCION    = "DISTRIBUCION DE MATRICES"

# Matriz USD
PFX_MATRIZ_USD = "$2025 MATRIZ USD"
SHEET_MATRIZ_2025 = "2025"

FN_INV_PLANTILLA = "$2025 INVENTARIO GENERAL.xlsx"
SHEET_INV_ORIG   = "INVENTARIO"
SHEET_INV_COPIA  = "INVENTARIO COPIA"
SHEET_INV_LISTA  = "INV LISTA PRECIOS"

HEADER_ROW_INV         = 2
HEADER_ROW_INV_LISTA   = 1
HEADER_ROW_VAL         = 9
HEADER_ROW_MATRIZ      = 1

# Columnas a limpiar en INVENTARIO COPIA
COLS_A_LIMPIAR = [
    "REFERENCIA", "NOMBRE LISTA", "NOMBRE ODOO", "NOMBRE MYR",
    "MARCA copia", "INV BODEGA", "EXISTENCIA AGO 26", "COSTO PROMEDIO",
    "LINEA COPIA", "SUB-LINEA COPIA", "LIDER LINEA", "CLASIFICACION",
    "Marca sistema", "Linea sistema", "Sub- linea sistema"
]

# Columnas a traer desde INVENTARIO original
COLS_DESDE_ORIGINAL = ["MARCA copia", "INV BODEGA GERENCIA", "LINEA COPIA", "SUB-LINEA COPIA", "LIDER LINEA", "CLASIFICACION"]

# ==== DEPENDENCIAS (COM) ====
try:
    import win32com.client as win32
    HAS_COM = True
except Exception:
    HAS_COM = False

# ==== UTILS BÁSICAS ====
def log(msg): print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

# Cache para normalización
_NORM_CACHE = {}
def _norm(s: str) -> str:
    if s in _NORM_CACHE:
        return _NORM_CACHE[s]
    t = unidecode(str(s)).lower()
    t = re.sub(r"[^a-z0-9 ]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    _NORM_CACHE[s] = t
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

# ==== ARCHIVOS / LECTURA ====
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
    excel.ScreenUpdating = False
    
    try: excel.AskToUpdateLinks = False
    except Exception: pass
    try: excel.AutomationSecurity = 3
    except Exception: pass

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
    wb.SaveAs(str(tmp), FileFormat=51)
    wb.Close(SaveChanges=False)
    excel.Quit()
    return tmp

def open_as_excel_source(path: Path, passwords: list[str] | None = None):
    """Devuelve un 'source' para pandas."""
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
    """Elige la mejor hoja."""
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
    """Lee una hoja con header en 'header_row_visible' (1-based)."""
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

def cargar_matriz_usd(base_dir: Path) -> pd.DataFrame:
    """
    Carga el archivo MATRIZ USD, hoja 2025.
    """
    try:
        p = find_by_prefix(base_dir, PFX_MATRIZ_USD)
        log(f"Abriendo Matriz USD: {p.name}")
        
        src = open_as_excel_source(p, PASSWORDS_TRY)
        
        xf = pd.ExcelFile(src, engine="openpyxl")
        sheet_found = None
        for sn in xf.sheet_names:
            if "2025" in sn or _norm(sn) == "2025":
                sheet_found = sn
                break
        if not sheet_found:
            sheet_found = xf.sheet_names[0]
        
        df_raw = pd.read_excel(src, sheet_name=sheet_found, engine="openpyxl", header=None)
        
        header_row_idx = None
        for idx in range(min(20, len(df_raw))):
            row_values = df_raw.iloc[idx].astype(str).str.lower()
            has_ref = any("referencia" in str(v).lower() and "fertrac" in str(v).lower() for v in df_raw.iloc[idx])
            has_desc = any("descripcion" in str(v).lower() and "lista" in str(v).lower() for v in df_raw.iloc[idx])
            
            if has_ref or has_desc:
                header_row_idx = idx
                break
        
        if header_row_idx is None:
            max_non_empty = 0
            for idx in range(min(10, len(df_raw))):
                non_empty = df_raw.iloc[idx].notna().sum()
                if non_empty > max_non_empty:
                    max_non_empty = non_empty
                    header_row_idx = idx
            log(f"  Usando fila {header_row_idx + 1} como encabezado (más valores no vacíos)")
        
        df = pd.read_excel(src, sheet_name=sheet_found, engine="openpyxl", header=header_row_idx)
        
        df.columns = [str(c).strip() if not str(c).startswith("Unnamed") and str(c) != "nan" else f"_COL_{i}" 
                      for i, c in enumerate(df.columns)]        
      
        idx = {_norm(c): c for c in df.columns}
        
        ref_col = None
        for col_name in df.columns:
            col_norm = _norm(col_name)
            if "referencia" in col_norm and ("fertrac" in col_norm or "inventario" in col_norm):
                ref_col = col_name
                break
        
        if not ref_col:
            for col_name in df.columns[:5]:
                non_null = df[col_name].notna().sum()
                if non_null > 10:
                    sample = df[col_name].dropna().astype(str).head(5)
                    if any("FP-" in str(v) or str(v).replace("-", "").isdigit() for v in sample):
                        ref_col = col_name
                        log(f"  Usando columna '{col_name}' como REFERENCIA (detectada por patrón)")
                        break
        
        desc_col = None
        for col_name in df.columns:
            col_norm = _norm(col_name)
            if "descripcion" in col_norm and "lista" in col_norm and "precio" in col_norm:
                desc_col = col_name
                break
        
        if not desc_col:
            for col_name in df.columns:
                if col_name == ref_col:
                    continue
                non_null = df[col_name].notna().sum()
                if non_null > 10:
                    sample = df[col_name].dropna().astype(str).head(5)
                    avg_len = sum(len(str(v)) for v in sample) / len(sample) if len(sample) > 0 else 0
                    if avg_len > 15:
                        desc_col = col_name
                        log(f"  Usando columna '{col_name}' como DESCRIPCION (detectada por longitud)")
                        break
        
        if not ref_col:
            raise KeyError(f"No encontré columna 'REFERENCIA INVENTARIO FERTRAC' en {p.name}. Columnas: {list(df.columns)}")
        if not desc_col:
            raise KeyError(f"No encontré columna 'DESCRIPCION LISTA PRECIOS' en {p.name}. Columnas: {list(df.columns)}")
        
       
        df = df[~df[ref_col].isna() & (df[ref_col].astype(str).str.strip() != "")].copy()
        
        out = pd.DataFrame()
        out["__REF_MATRIZ__"] = df[ref_col].apply(to_num_str)
        out["__DESC_LISTA__"] = df[desc_col].fillna("")
        
        out = out.drop_duplicates(subset=["__REF_MATRIZ__"], keep="first")
        
        return out
        
    except FileNotFoundError:
        log(f"⚠ ADVERTENCIA: No se encontró el archivo '{PFX_MATRIZ_USD}'. Se continuará sin actualizar NOMBRE LISTA desde Matriz USD.")
        return pd.DataFrame(columns=["__REF_MATRIZ__", "__DESC_LISTA__"])
    except Exception as e:
        log(f"⚠ ERROR al cargar Matriz USD: {e}")
        import traceback
        log(traceback.format_exc())
        log(f"  Se continuará sin actualizar NOMBRE LISTA desde Matriz USD.")
        return pd.DataFrame(columns=["__REF_MATRIZ__", "__DESC_LISTA__"])
    
def cargar_marcas(base_dir: Path) -> set:
    """
    Carga el archivo MARCAS y retorna un set con las marcas propias.
    """
    try:
        p = find_by_prefix(base_dir, PFX_MARCAS)
        log(f"Abriendo archivo MARCAS: {p.name}")
        
        src = open_as_excel_source(p, PASSWORDS_TRY)        
        df = pd.read_excel(src, sheet_name=0, engine="openpyxl", header=None)      
        marcas_propias = set()
        
        for col_idx in range(min(3, len(df.columns))):
            for val in df[col_idx].dropna():
                val_str = str(val).strip().upper()
                if val_str and val_str not in ("", "NONE", "NAN", "MARCA", "MARCAS"):
                    if any(c.isalpha() for c in val_str):
                        marcas_propias.add(val_str)
        
        log(f"{len(marcas_propias)} marcas propias cargadas")
        return marcas_propias
        
    except FileNotFoundError:
        log(f"⚠ ADVERTENCIA: No se encontró el archivo '{PFX_MARCAS}'")
        return set()
    except Exception as e:
        log(f"⚠ ERROR al cargar MARCAS: {e}")
        import traceback
        log(traceback.format_exc())
        return set()


def cargar_distribucion(base_dir: Path) -> dict:
    """
    Carga el archivo DISTRIBUCIÓN DE MATRICES.
    """
    try:
        p = find_by_prefix(base_dir, PFX_DISTRIBUCION)
        log(f"Abriendo archivo DISTRIBUCIÓN: {p.name}")
        
        src = open_as_excel_source(p, PASSWORDS_TRY)
        xf = pd.ExcelFile(src, engine="openpyxl")
        sheet_name = None
        for sn in xf.sheet_names:
            if "DISTRIBUCION" in _norm(sn) or "MATRICES" in _norm(sn):
                sheet_name = sn
                break
        
        if not sheet_name:
            sheet_name = xf.sheet_names[0]
        
        df_raw = pd.read_excel(src, sheet_name=sheet_name, engine="openpyxl", header=None)
        header_row = None
        for idx in range(min(10, len(df_raw))):
            row_str = ' '.join([str(v).upper() for v in df_raw.iloc[idx] if pd.notna(v)])
            if "LINEA" in row_str and "GESTOR" in row_str:
                header_row = idx
                break
        
        if header_row is None:
            header_row = 2         
        df = pd.read_excel(src, sheet_name=sheet_name, engine="openpyxl", header=header_row)
        df.columns = [str(c).strip() for c in df.columns]
        
        idx = {_norm(c): c for c in df.columns}
        linea_col = (
            idx.get("linea") or idx.get("línea") or idx.get("marca")
            or next((real for kn, real in idx.items() if "linea" in kn or "línea" in kn), None)
        )

        gestor_col = (
            idx.get("gestor") or idx.get("lider") or idx.get("líder") 
            or next((real for kn, real in idx.items() if "gestor" in kn or "lider" in kn), None)
        )
        
        clasif_col = (
            idx.get("categoria") or idx.get("categoría") or idx.get("clasificacion") 
            or idx.get("clasificación") or idx.get("tipo")
            or next((real for kn, real in idx.items() if "categ" in kn or "clasificac" in kn), None)
        )
        
        if not linea_col:
            log(f"  ⚠ No se encontró columna de LINEA/MARCA")
            return {'gestor': {}, 'clasificacion': {}}
        
        gestor_map = {}
        clasif_map = {}

        for idx_row, row in df.iterrows():
            linea_val = row[linea_col] if linea_col in row.index else None
            if pd.isna(linea_val) or str(linea_val).strip() == "":
                continue
            
            linea_str = str(linea_val).strip().upper()
            
            linea_key = re.sub(r'\s*\([^)]*\)\s*', '', linea_str).strip()
            
            if not linea_key:
                continue
        
            if gestor_col and gestor_col in row.index:
                gestor_val = row[gestor_col]
                if not pd.isna(gestor_val) and str(gestor_val).strip():
                    gestor_map[linea_key] = str(gestor_val).strip()
            
            if clasif_col and clasif_col in row.index:
                clasif_val = row[clasif_col]
                if not pd.isna(clasif_val) and str(clasif_val).strip():
                    clasif_map[linea_key] = str(clasif_val).strip()
        
        log(f"{len(gestor_map)} gestores cargados")
        log(f"{len(clasif_map)} clasificaciones cargadas")
        
        return {
            'gestor': gestor_map,
            'clasificacion': clasif_map
        }
        
    except FileNotFoundError:
        log(f"⚠ ADVERTENCIA: No se encontró el archivo '{PFX_DISTRIBUCION}'")
        return {'gestor': {}, 'clasificacion': {}}
    except Exception as e:
        log(f"⚠ ERROR al cargar DISTRIBUCIÓN: {e}")
        import traceback
        log(traceback.format_exc())
        return {'gestor': {}, 'clasificacion': {}}


def aplicar_reglas_marcas_propias(ws_inv_copia, start_data_row: int, last_row: int, 
                                   ref_col_idx: int, hdrn_copia: dict, 
                                   marcas_propias: set, distribucion: dict):
    """
    Aplica las reglas de negocio para referencias de marcas propias:
    """
    try:
        log("Aplicando reglas para marcas propias...")
        
        col_linea_copia = hdrn_copia.get(_norm("LINEA COPIA"))
        col_marca_sistema = hdrn_copia.get(_norm("MARCA SISTEMA")) or hdrn_copia.get(_norm("Marca sistema"))
        col_marca_copia = hdrn_copia.get(_norm("MARCA COPIA")) or hdrn_copia.get(_norm("MARCA copia"))
        col_inv_bodega_ger = hdrn_copia.get(_norm("INV BODEGA GERENCIA"))
        col_sublinea_copia = hdrn_copia.get(_norm("SUB-LINEA COPIA"))
        col_lider_linea = hdrn_copia.get(_norm("LIDER LINEA"))
        col_clasificacion = hdrn_copia.get(_norm("CLASIFICACION"))
        
        if not col_linea_copia or not col_marca_sistema:
            log("  ⚠ No se encontraron columnas necesarias (LINEA COPIA o MARCA SISTEMA)")
            log(f"    LINEA COPIA: {'✓' if col_linea_copia else '✗'}")
            log(f"    MARCA SISTEMA: {'✓' if col_marca_sistema else '✗'}")
            return
        
        
        cols_to_read = [ref_col_idx, col_linea_copia, col_marca_sistema]
        data = read_multiple_columns_optimized(ws_inv_copia, start_data_row, last_row, cols_to_read)
        
        referencias = data[ref_col_idx]
        lineas_copia = data[col_linea_copia]
        marcas_sistema = data[col_marca_sistema]
        
        filas_a_procesar = []
        filas_a_eliminar = []
       
        filtro1_rechazados = 0
        filtro2_rechazados = 0

        log(f"  Analizando {len(referencias)} registros...")
        
        for i in range(len(referencias)):
            ref = str(referencias[i]).strip() if referencias[i] not in (None, "", "None") else ""
            linea = str(lineas_copia[i]).strip() if lineas_copia[i] not in (None, "", "None") else ""
            marca = str(marcas_sistema[i]).strip() if marcas_sistema[i] not in (None, "", "None") else ""
            
            if ref and re.match(r'^\d{4}[A-Za-z]$', ref):
                filas_a_eliminar.append((i, ref))
                continue
            
            linea_upper = linea.upper()
            marca_upper = marca.upper()
            
            if linea or linea_upper in ("INDETERMINADO", "#N/D", "#N/A", "N/A", "NA", "NONE"):
                continue
            
            if marca_upper not in marcas_propias:
                continue
            
            filas_a_procesar.append((i, ref, marca))
        

        # Eliminar filas tipo "0041R" (en orden inverso para no afectar índices)
        if filas_a_eliminar:
            log(f"  Eliminando referencias tipo '####L':")
            for idx, ref in filas_a_eliminar[:5]:  
                log(f"    - {ref}")
            if len(filas_a_eliminar) > 5:
                log(f"    ... y {len(filas_a_eliminar) - 5} más")
            
            for idx, ref in sorted(filas_a_eliminar, reverse=True):
                fila_excel = start_data_row + idx
                try:
                    ws_inv_copia.Rows(fila_excel).Delete()
                except Exception as e:
                    log(f"    ⚠ Error al eliminar fila {fila_excel} ({ref}): {e}")
            
            log(f"{len(filas_a_eliminar)} filas eliminadas")
            
            # Recalcular last_row después de eliminar
            last_row = last_row - len(filas_a_eliminar)
            log(f"  Nuevo rango de datos: hasta fila {last_row}")
        
        if not filas_a_procesar:
            log("  ℹ No hay registros para procesar después de aplicar filtros")
            return
        
        log(f"  Procesando {len(filas_a_procesar)} registros ...")
        
        gestor_map = distribucion.get('gestor', {})
        clasif_map = distribucion.get('clasificacion', {})
        
        updates = {}
        for idx, ref, marca in filas_a_procesar:
            fila_excel = start_data_row + idx
            marca_upper = marca.upper().strip()
            
            updates[fila_excel] = {
                'marca': marca,
                'inv_bodega': "0",
                'linea': marca,
                'sublinea': marca,
                'lider': gestor_map.get(marca_upper, ""),
                'clasificacion': clasif_map.get(marca_upper, "")
            }
        
        filas_ordenadas = sorted(updates.keys())
        
        def agrupar_contiguos(filas):
            """Agrupa filas contiguas para actualización por rangos"""
            if not filas:
                return []
            grupos = []
            inicio = filas[0]
            fin = filas[0]
            for fila in filas[1:]:
                if fila == fin + 1:
                    fin = fila
                else:
                    grupos.append((inicio, fin))
                    inicio = fin = fila
            grupos.append((inicio, fin))
            return grupos
        
        grupos_contiguos = agrupar_contiguos(filas_ordenadas)
        
        columnas_actualizadas = 0
        lideres_encontrados = sum(1 for v in updates.values() if v['lider'])
        clasif_encontradas = sum(1 for v in updates.values() if v['clasificacion'])
        
        if col_marca_copia:
            for inicio, fin in grupos_contiguos:
                valores = [updates[f]['marca'] for f in range(inicio, fin + 1)]
                rng = ws_inv_copia.Range(
                    ws_inv_copia.Cells(inicio, col_marca_copia),
                    ws_inv_copia.Cells(fin, col_marca_copia)
                )
                rng.Value = [[v] for v in valores]
            columnas_actualizadas += 1
        
        if col_inv_bodega_ger:
            for inicio, fin in grupos_contiguos:
                valores = [updates[f]['inv_bodega'] for f in range(inicio, fin + 1)]
                rng = ws_inv_copia.Range(
                    ws_inv_copia.Cells(inicio, col_inv_bodega_ger),
                    ws_inv_copia.Cells(fin, col_inv_bodega_ger)
                )
                rng.Value = [[v] for v in valores]
            columnas_actualizadas += 1
        
        if col_linea_copia:
            for inicio, fin in grupos_contiguos:
                valores = [updates[f]['linea'] for f in range(inicio, fin + 1)]
                rng = ws_inv_copia.Range(
                    ws_inv_copia.Cells(inicio, col_linea_copia),
                    ws_inv_copia.Cells(fin, col_linea_copia)
                )
                rng.Value = [[v] for v in valores]
            columnas_actualizadas += 1
        
        if col_sublinea_copia:
            for inicio, fin in grupos_contiguos:
                valores = [updates[f]['sublinea'] for f in range(inicio, fin + 1)]
                rng = ws_inv_copia.Range(
                    ws_inv_copia.Cells(inicio, col_sublinea_copia),
                    ws_inv_copia.Cells(fin, col_sublinea_copia)
                )
                rng.Value = [[v] for v in valores]
            columnas_actualizadas += 1
        
        if col_lider_linea:
            for inicio, fin in grupos_contiguos:
                valores = [updates[f]['lider'] for f in range(inicio, fin + 1)]
                rng = ws_inv_copia.Range(
                    ws_inv_copia.Cells(inicio, col_lider_linea),
                    ws_inv_copia.Cells(fin, col_lider_linea)
                )
                rng.Value = [[v] for v in valores]
            columnas_actualizadas += 1
        
        if col_clasificacion:
            for inicio, fin in grupos_contiguos:
                valores = [updates[f]['clasificacion'] for f in range(inicio, fin + 1)]
                rng = ws_inv_copia.Range(
                    ws_inv_copia.Cells(inicio, col_clasificacion),
                    ws_inv_copia.Cells(fin, col_clasificacion)
                )
                rng.Value = [[v] for v in valores]
            columnas_actualizadas += 1
        
        log(f" Proceso completado:")
    except Exception as e:
        log(f"  ⚠ ERROR al aplicar reglas de marcas propias: {e}")
        import traceback
        log(traceback.format_exc())

# ==== EXCEL COM ====
def excel_open(path: Path, password: str | None = None):
    """Abre con COM en modo silencioso y OPTIMIZADO."""
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.Interactive = False
    excel.EnableEvents = False
    excel.ScreenUpdating = False
    
    try: excel.AskToUpdateLinks = False
    except Exception: pass
    try: excel.AutomationSecurity = 3
    except Exception: pass

    info = {"tmp_path": None, "target_path": str(path), "reapply_password": None}

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
        try:
            excel.Calculation = -4135  
        except Exception as e:
            log(f"Aviso: no se pudo establecer cálculo manual: {e}")
        return excel, wb, info
    except Exception as e:
        excel.Quit()
        raise RuntimeError(f"No pude abrir el libro {path.name} de forma silenciosa.") from e

def excel_close(excel, wb, save=True):
    try:
        if save:
            excel.Calculation = -4105  
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

# ==== AJUSTES PIVOT-SAFE ====
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
    """Lista de bloques de pivots [(r1, r2, c1, c2), ...]."""
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

def ws_ensure_range(ws, start_row: int, expected_rows: int, header_row: int) -> int:
    """
    Asegura que el rango detectado incluya todas las filas esperadas.
    """
    calculated_last = start_row + expected_rows - 1
    
    pivot_top = ws_first_pivot_row(ws)
    if pivot_top and pivot_top > header_row:
        if calculated_last >= pivot_top:
            log(f"⚠ Límite por pivot: reduciendo de {calculated_last} a {pivot_top - 1}")
            return pivot_top - 1
    
    return calculated_last


def ws_apply_borders_to_range(ws, start_row: int, end_row: int, start_col: int, end_col: int):
    """
    Aplica bordes a un rango completo de celdas.
    """
    try:
        
        full_range = ws.Range(ws.Cells(start_row, start_col), ws.Cells(end_row, end_col))
        
        for border_id in [7, 8, 9, 10, 11, 12]:
            try:
                full_range.Borders(border_id).LineStyle = 1      
                full_range.Borders(border_id).Weight = 2         
                full_range.Borders(border_id).ColorIndex = -4105 
            except Exception:
                continue
            
    except Exception as e:
        log(f"  ⚠ Error al aplicar bordes: {e}")


def ws_remove_formatting_from_range(ws, start_row: int, end_row: int, start_col: int, end_col: int):
    """
    Elimina formato de negrita y color de fondo de un rango.
    Preserva los formatos de número (General, Contabilidad, etc.) de cada columna.
    """
    try:
        
        for col in range(start_col, end_col + 1):
            try:
                original_number_format = ws.Cells(start_row, col).NumberFormat

                col_range = ws.Range(ws.Cells(start_row, col), ws.Cells(end_row, col))
                
                try:
                    col_range.Font.Bold = False
                except Exception:
                    pass
                
                try:
                    col_range.Interior.ColorIndex = 0  
                except Exception:
                    pass
                
                try:
                    col_range.NumberFormat = original_number_format
                except Exception:
                    pass
                    
            except Exception:
                continue
        
       
    except Exception as e:
        log(f"  ⚠ Error al limpiar formato: {e}")


def ws_update_subtotal_formula(ws, formula_row: int, last_data_row: int):
    """Actualiza la fórmula de subtotal en la fila 1 para que abarque todo el rango."""
    try:
        
        used_cols = ws.UsedRange.Columns.Count
        updated_count = 0
        
        for col in range(1, used_cols + 1):
            try:
                cell_formula = ws.Cells(formula_row, col).Formula
                
                if cell_formula and "SUBTOTAL" in str(cell_formula).upper():
                    import re
                    match = re.search(r'SUBTOTAL\((\d+),', str(cell_formula))
                    
                    if match:
                        func_num = match.group(1)
                        col_letter = _col_num_to_letter(col)
                        new_formula = f"=SUBTOTAL({func_num},{col_letter}3:{col_letter}{last_data_row})"
                        ws.Cells(formula_row, col).Formula = new_formula
                        updated_count += 1
                        
            except Exception:
                continue      
        
    except Exception as e:
        log(f"  ⚠ Error al actualizar fórmulas de subtotal: {e}")


def ws_add_final_subtotals(ws, last_data_row: int, header_row: int, hdrn: dict):
    """
    Agrega subtotales al final de todos los registros para EXISTENCIA y TOTAL INV.
    """
    try:
        log(f"Agregando subtotales finales en fila {last_data_row + 1}...")
        
        subtotal_row = last_data_row + 1
        
        # Buscar columna EXISTENCIA
        exist_col = None
        for name, col in hdrn.items():
            if name.startswith("existencia "):
                exist_col = col
                break
        

        total_inv_col = hdrn.get(_norm("TOTAL INV"))      
        header_color = None
        try:
            header_color = ws.Cells(header_row, 1).Interior.Color
        except Exception:
            header_color = 15849925  
        
        subtotals_added = 0

        if exist_col:
            try:
                col_letter = _col_num_to_letter(exist_col)
                formula = f"=SUBTOTAL(9,{col_letter}{header_row + 1}:{col_letter}{last_data_row})"
                
                cell = ws.Cells(subtotal_row, exist_col)               
                cell.Formula = formula
                
                try:
                    cell.Calculate()
                    valor = cell.Value
                    
                    if valor is not None:
                        cell.Value = valor

                except Exception as e:
                    log(f"    ⚠ No se pudo calcular valor: {e}")
                
                cell.NumberFormat = "#.##0"
                cell.Font.Bold = True
                cell.Interior.Color = header_color
                
                try:
                    for border_id in [7, 8, 9, 10]:
                        cell.Borders(border_id).LineStyle = 1
                        cell.Borders(border_id).Weight = 2
                        cell.Borders(border_id).ColorIndex = -4105
                except Exception:
                    pass
                
                formato_final = cell.NumberFormat
                valor_mostrado = cell.Text

                subtotals_added += 1                            
                
            except Exception as e:
                log(f"    ⚠ Error al agregar subtotal EXISTENCIA: {e}")
                import traceback
                log(traceback.format_exc())
        
        # Subtotal TOTAL INV
        if total_inv_col:
            try:
                col_letter = _col_num_to_letter(total_inv_col)
                formula = f"=SUBTOTAL(109,{col_letter}{header_row + 1}:{col_letter}{last_data_row})"
                
                cell = ws.Cells(subtotal_row, total_inv_col)               
                cell.Formula = formula
                cell.Font.Bold = True
                cell.Interior.Color = header_color

                try:
                    original_format = ws.Cells(last_data_row, total_inv_col).NumberFormat
                    cell.NumberFormat = original_format
                except Exception as e:
                    log(f"    ⚠ No se pudo copiar formato original: {e}")

                try:
                    for border_id in [7, 8, 9, 10]:
                        cell.Borders(border_id).LineStyle = 1
                        cell.Borders(border_id).Weight = 2
                        cell.Borders(border_id).ColorIndex = -4105
                except Exception:
                    pass
                
                subtotals_added += 1
                
            except Exception as e:
                log(f"    ⚠ Error al agregar subtotal TOTAL INV: {e}")
        
        if subtotals_added > 0:
            log(f"{subtotals_added} subtotales agregados exitosamente con formato")
        else:
            log(f"  ⚠ No se pudieron agregar subtotales")
        
    except Exception as e:
        log(f"  ⚠ Error al agregar subtotales finales: {e}")

def _col_num_to_letter(col_num):
    """Convierte número de columna a letra."""
    letter = ''
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        letter = chr(65 + remainder) + letter
    return letter


def _ranges_without_pivots_for_column(col_idx: int, start_row: int, end_row: int, pivot_blocks):
    """Devuelve sub-rangos [a,b] dentro de [start_row,end_row] que NO cruzan pivots."""
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

# ==== WS UTILS  ====
def ws_last_row(ws, key_col_idx: int, header_row_visible: int):
    """Última fila con datos."""
    last = ws.Cells(ws.Rows.Count, key_col_idx).End(-4162).Row
    return max(last, header_row_visible)

def ws_fill_column_values(ws, col_idx: int, start_row: int, values: list):
    """Escribe valores en una columna saltando pivots."""
    n = len(values)
    if n == 0:
        return

    end_row = start_row + n - 1
    pivots = ws_pivot_blocks(ws)
    safe_segments = _ranges_without_pivots_for_column(col_idx, start_row, end_row, pivots)

    offset = 0
    for (a, b) in safe_segments:
        if offset >= n:
            break
        seg_len = min(b - a + 1, n - offset)
        if seg_len <= 0:
            continue

        chunk = values[offset: offset + seg_len]
        chunk = [("" if (v is None or (isinstance(v, float) and np.isnan(v))) else v) for v in chunk]

        rng = ws.Range(ws.Cells(a, col_idx), ws.Cells(a + seg_len - 1, col_idx))
        rng.Value = [[v] for v in chunk]
        offset += seg_len

def ws_clear_column(ws, col_idx: int, start_row: int, end_row: int):
    """Limpia una columna por tramos, evitando pivots."""
    if end_row < start_row:
        return
    pivots = ws_pivot_blocks(ws)
    safe_segments = _ranges_without_pivots_for_column(col_idx, start_row, end_row, pivots)
    for (a, b) in safe_segments:
        rng = ws.Range(ws.Cells(a, col_idx), ws.Cells(b, col_idx))
        rng.ClearContents()

def ws_copy_down_formula(ws, col_idx: int, start_row: int, end_row: int):
    """Copia la fórmula desde start_row hasta end_row."""
    if end_row < start_row: return
    fml = ws.Cells(start_row, col_idx).Formula
    if fml:
        rng = ws.Range(ws.Cells(start_row, col_idx), ws.Cells(end_row, col_idx))
        rng.Formula = fml

def ws_headers_smart(ws, expected_row: int, preferred_labels: list[str] | None = None):
    """Detecta de forma robusta la fila de encabezado."""
    preferred_norm = [_norm(x) for x in (preferred_labels or [])]
    tried = [expected_row] + [r for r in range(1, 11) if r != expected_row]
    for hr in tried:
        hdr, hdrn = ws_headers(ws, hr)
        if not hdrn:
            continue
        if not preferred_norm or any(lbl in hdrn for lbl in preferred_norm):
            return hr, hdr, hdrn
    try:
        first_row = ws.UsedRange.Row
        hdr, hdrn = ws_headers(ws, first_row)
        if hdrn:
            return first_row, hdr, hdrn
    except Exception:
        pass
    return expected_row, {}, {}

def find_reference_col_idx(hdrn: dict, ws, header_row_used: int) -> int:
    """Encuentra índice de columna para REFERENCIA."""
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
    """Devuelve col_idx del encabezado EXISTENCIA {MES DD}."""
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

def normalize_sheet_name(wb, target_name: str) -> str:
    """Normaliza el nombre de una hoja eliminando espacios extras."""
    target_norm = _norm(target_name)
    
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        sheet_name = ws.Name
        sheet_norm = _norm(sheet_name)
        
        if sheet_norm == target_norm or target_norm in sheet_norm:
            clean_name = sheet_name.strip()
            if clean_name != sheet_name:
                try:
                    ws.Name = clean_name
                    log(f"Nombre de hoja normalizado: '{sheet_name}' → '{clean_name}'")
                    return clean_name
                except Exception as e:
                    log(f"No se pudo renombrar hoja: {e}")
                    return sheet_name
            return clean_name
    
    return target_name

def read_range_as_array(ws, start_row: int, end_row: int, col_idx: int):
    """Lee un rango completo en una sola operación."""
    if end_row < start_row:
        return []
    rng = ws.Range(ws.Cells(start_row, col_idx), ws.Cells(end_row, col_idx))
    values = rng.Value
    if values is None:
        return [""] * (end_row - start_row + 1)
    if not isinstance(values, (list, tuple)):
        return [values]
    return [row[0] if isinstance(row, (list, tuple)) else row for row in values]

def write_range_as_array(ws, start_row: int, col_idx: int, values: list):
    """Escribe un rango completo en una sola operación."""
    if not values:
        return
    end_row = start_row + len(values) - 1
    rng = ws.Range(ws.Cells(start_row, col_idx), ws.Cells(end_row, col_idx))
    rng.Value = [[v] for v in values]

def read_multiple_columns_optimized(ws, start_row: int, end_row: int, col_indices: list[int]) -> dict:
    """Lee múltiples columnas en UNA SOLA operación - OPTIMIZACIÓN CRÍTICA."""
    if end_row < start_row or not col_indices:
        return {col: [] for col in col_indices}
    
    min_col = min(col_indices)
    max_col = max(col_indices)
    
    rng = ws.Range(ws.Cells(start_row, min_col), ws.Cells(end_row, max_col))
    values = rng.Value
    
    if values is None:
        return {col: [""] * (end_row - start_row + 1) for col in col_indices}
    
    if not isinstance(values[0], (list, tuple)):
        values = [values]
    
    result = {}
    for col_idx in col_indices:
        offset = col_idx - min_col
        result[col_idx] = [row[offset] if isinstance(row, (list, tuple)) and len(row) > offset else "" 
                          for row in values]
    
    return result

# ==== PROCESO PRINCIPAL ====
def main():
    if not HAS_COM:
        raise RuntimeError("Este script requiere Excel COM (win32com). Instálalo y ejecuta en Windows con Excel.")

    log("== Inicio actualización de inventario ==")

    # 1) Cargar datos externos
    log("Cargando datos externos...")
    df_src = cargar_inventario_actualizado(BASE_PATH)

    # Valorizados
    df_val_gen   = cargar_valorizado(BASE_PATH, PFX_VAL_GENERAL)
    df_val_impo  = cargar_valorizado(BASE_PATH, PFX_VAL_FALT_IMPO)
    df_val_falt  = cargar_valorizado(BASE_PATH, PFX_VAL_FALT)
    df_val_tob   = cargar_valorizado(BASE_PATH, PFX_VAL_TOBERIN)

    # Cargar Matriz USD
    df_matriz_usd = cargar_matriz_usd(BASE_PATH)
    matriz_map = df_matriz_usd.set_index("__REF_MATRIZ__")["__DESC_LISTA__"].to_dict() if len(df_matriz_usd) > 0 else {}
    log(f"Matriz USD: {len(matriz_map)} referencias disponibles para actualizar NOMBRE LISTA")

    # Cargar archivos auxiliares para marcas propias
    marcas_propias = cargar_marcas(BASE_PATH)
    log(f"Marcas propias: {len(marcas_propias)} marcas cargadas")
    
    distribucion = cargar_distribucion(BASE_PATH)
    log(f"Distribución: {len(distribucion['gestor'])} gestores, {len(distribucion['clasificacion'])} clasificaciones")

    # Join de cantidades
    val_map_impo = df_val_impo.set_index("__REF_INT__")["__CANT__"]
    val_map_falt = df_val_falt.set_index("__REF_INT__")["__CANT__"]
    val_map_tob  = df_val_tob.set_index("__REF_INT__")["__CANT__"]

    # Calcular columnas en VALORIZADO GENERAL
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

    # Consolidado EXISTENCIA_CALC
    # FÓRMULA: VALORIZADO GENERAL - FALTANTES IMPO - FALTANTES - TOBERÍN
    df_val_gen["__EXIST_CALC__"] = (
        df_val_gen["__CANT__"] 
        - df_val_gen["__IMPO_CANT__"] 
        - df_val_gen["__FALT_CANT__"] 
        - df_val_gen["__TOB_CANT__"]
    )
    exist_map = df_val_gen.set_index("__REF_INT__")["__EXIST_CALC__"]

    # 2) Abrir libro PLANTILLA
    p_inv = BASE_PATH / FN_INV_PLANTILLA
    log(f"Abriendo libro Excel: {p_inv}")
    excel, wb, saveinfo = excel_open(p_inv, password=PASS_INV)

    # 3) NORMALIZAR nombre de hoja INVENTARIO
    log("Normalizando nombre de hoja INVENTARIO...")
    normalized_inv_name = normalize_sheet_name(wb, SHEET_INV_ORIG)
    
    try:
        ws_inv_orig = wb.Worksheets(normalized_inv_name)
    except Exception:
        ws_inv_orig = wb.Worksheets(1)
        normalized_inv_name = ws_inv_orig.Name

    # 4) ELIMINAR hoja INVENTARIO COPIA si existe
    try:
        excel.DisplayAlerts = False
        for i in range(1, wb.Worksheets.Count + 1):
            try:
                sheet_name = wb.Worksheets(i).Name
                if _norm(sheet_name) == _norm(SHEET_INV_COPIA):
                    wb.Worksheets(i).Delete()
                    log(f"Hoja existente eliminada: {sheet_name}")
                    break
            except:
                pass
    except Exception as e:
        log(f"Error al eliminar hoja existente: {e}")

    # 5) CREAR nueva hoja INVENTARIO COPIA
    log("Creando nueva hoja INVENTARIO COPIA...")
    was_protected = False
    try:
        if ws_inv_orig.ProtectContents or ws_inv_orig.ProtectDrawingObjects or ws_inv_orig.ProtectScenarios:
            was_protected = True
            ws_inv_orig.Unprotect(Password=PASS_INV)
    except Exception as e:
        log(f"Aviso al desproteger: {e}")

    try:
        ws_inv_copia = wb.Worksheets.Add(After=ws_inv_orig)
        ws_inv_copia.Name = SHEET_INV_COPIA
        
        ws_inv_orig.UsedRange.Copy(Destination=ws_inv_copia.Range("A1"))
        
        try:
            for col in range(1, ws_inv_orig.UsedRange.Columns.Count + 1):
                ws_inv_copia.Columns(col).ColumnWidth = ws_inv_orig.Columns(col).ColumnWidth
        except Exception as e:
            log(f"Aviso: no se pudo copiar anchos de columna: {e}")
        
        log(f"Hoja '{SHEET_INV_COPIA}' creada exitosamente")
        
    except Exception as e:
        log(f"ERROR al crear copia: {e}")
        raise RuntimeError(f"No se pudo crear copia de la hoja INVENTARIO: {e}")

    if was_protected:
        try:
            ws_inv_orig.Protect(Password=PASS_INV, DrawingObjects=True, Contents=True, Scenarios=True)
            log("Hoja INVENTARIO original re-protegida")
        except Exception as e:
            log(f"Aviso al re-proteger: {e}")

    # 6) TRABAJAR EN INVENTARIO COPIA
    log("Trabajando en hoja INVENTARIO COPIA...")
    
    header_row_used, hdr_copia, hdrn_copia = ws_headers_smart(
        ws_inv_copia,
        expected_row=HEADER_ROW_INV,
        preferred_labels=["REFERENCIA", "REFERENCIA FERTRAC"]
    )
    log(f"Encabezados detectados en fila {header_row_used} de INVENTARIO COPIA")

    ref_col_idx = find_reference_col_idx(hdrn_copia, ws_inv_copia, header_row_used)
    start_data_row = header_row_used + 1

    # Detectar rango inicial solo para referencia
    initial_last_row = ws_last_row(ws_inv_copia, ref_col_idx, header_row_used)
    log(f"Rango inicial detectado: {initial_last_row - start_data_row + 1} filas")

    # El last_row real se calculará después de pegar los datos
    last_row = initial_last_row

    # 7) LIMPIAR columnas en INVENTARIO COPIA
    # Calcular el rango máximo esperado ANTES de limpiar
    log("Calculando rango esperado para limpieza...")
    expected_rows = len(df_src["__REFERENCIA__"])
    max_last_row = ws_ensure_range(ws_inv_copia, start_data_row, expected_rows, header_row_used)

    log(f"Limpiando columnas en INVENTARIO COPIA (hasta fila {max_last_row})...")
    for colname in COLS_A_LIMPIAR:
        cidx = hdrn_copia.get(_norm(colname))
        if cidx:
            ws_clear_column(ws_inv_copia, cidx, start_data_row, max_last_row)


    # 8) Limpiar REFERENCIA FERTRAC en INV LISTA PRECIOS
    log("Limpiando REFERENCIA FERTRAC en INV LISTA PRECIOS...")
    try:
        ws_lp = None
        target_norm = _norm(SHEET_INV_LISTA)
        
        for i in range(1, wb.Worksheets.Count + 1):
            sheet_name = wb.Worksheets(i).Name
            if _norm(sheet_name) == target_norm or target_norm in _norm(sheet_name):
                ws_lp = wb.Worksheets(i)
                log(f"Hoja encontrada: '{sheet_name}'")
                break
        
        if ws_lp is None:
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
                pivot_top_lp = ws_first_pivot_row(ws_lp)
                if pivot_top_lp and pivot_top_lp > hr_lp:
                    last_row_lp = min(last_row_lp, pivot_top_lp - 1)
                ws_clear_column(ws_lp, cidx, hr_lp + 1, last_row_lp)
                log("REFERENCIA FERTRAC limpiada exitosamente")
            else:
                log("Columna REFERENCIA FERTRAC no encontrada")
        else:
            log("Hoja INV LISTA PRECIOS no encontrada")
                
    except Exception as e:
        log(f"No se pudo procesar 'INV LISTA PRECIOS': {e}")

    # 9) PEGAR columnas desde datos externos en INVENTARIO COPIA
    log("Pegando columnas desde Inventario actualizado en INVENTARIO COPIA...")
    ref_values   = df_src["__REFERENCIA__"].tolist()
    nombre_odoo  = df_src.get("__NOMBRE__",       pd.Series([""]*len(ref_values))).tolist()
    marca_sys    = df_src.get("__MARCA_SYS__",    pd.Series([""]*len(ref_values))).tolist()
    linea_sys    = df_src.get("__LINEA_SYS__",    pd.Series([""]*len(ref_values))).tolist()
    sublinea_sys = df_src.get("__SUBLINEA_SYS__", pd.Series([""]*len(ref_values))).tolist()
    costo_prom   = df_src.get("__COSTO__",        pd.Series([np.nan]*len(ref_values))).tolist()

    def paste_if_exists(col_name, values, number_format=None):
        cidx = hdrn_copia.get(_norm(col_name))
        if not cidx:
            log(f"  - Columna no encontrada: {col_name}")
            return
        
        if col_name == "REFERENCIA":
            has_slash = any("/" in str(v) for v in values if v not in (None, "", np.nan))
            
            if has_slash:
                
                rng = ws_inv_copia.Range(
                    ws_inv_copia.Cells(start_data_row, cidx),
                    ws_inv_copia.Cells(start_data_row + len(values) - 1, cidx)
                )
                
                rng.NumberFormat = "@"
                ws_fill_column_values(ws_inv_copia, cidx, start_data_row, values)
                
                try:
                    converted_values = []
                    for v in values:
                        if v in (None, "", np.nan):
                            converted_values.append([""])
                        elif "/" in str(v) or not str(v).replace(".", "").replace("-", "").isdigit():
                            converted_values.append([v])
                        else:
                            try:
                                converted_values.append([float(v)])
                            except:
                                converted_values.append([v])
                    
                    rng.Value = converted_values
                except Exception as e:
                    log(f"    Aviso en conversión: {e}")
                
                rng.NumberFormat = "0"
                
                try:
                    rng.HorizontalAlignment = -4131  # xlLeft
                except Exception as e:
                    log(f"    Aviso en alineación: {e}")
                
                try:
                    for i in range(1, 8):
                        try:
                            rng.Errors.Item(i).Ignore = True
                        except:
                            pass
                    ws_inv_copia.Parent.Application.ErrorCheckingOptions.NumberAsText = False
                except Exception:
                    pass
                
                log(f"Pegada columna: {col_name} (formato número, alineación izquierda)")
                return
        
        ws_fill_column_values(ws_inv_copia, cidx, start_data_row, values)
        if number_format:
            ws_inv_copia.Columns(cidx).NumberFormat = number_format


    paste_if_exists("REFERENCIA", ref_values, number_format="0")
    paste_if_exists("NOMBRE ODOO", nombre_odoo)
    paste_if_exists("Marca sistema", marca_sys)
    paste_if_exists("Linea sistema", linea_sys)
    paste_if_exists("Sub- linea sistema", sublinea_sys)
    paste_if_exists("COSTO PROMEDIO", costo_prom)

    log("Recalculando rango de datos después de pegar...")
    new_last_row = start_data_row + len(ref_values) - 1

    # Verificar si hay pivots que limiten el rango
    pivot_top = ws_first_pivot_row(ws_inv_copia)
    if pivot_top and pivot_top > header_row_used:
        # Si los nuevos datos sobrepasan el pivot, advertir
        if new_last_row >= pivot_top:
            log(f"⚠ ADVERTENCIA: Los datos ({new_last_row} filas) sobrepasan el inicio de la tabla pivote (fila {pivot_top})")
            log(f"  Se procesarán solo las filas hasta {pivot_top - 1}")
            last_row = pivot_top - 1
        else:
            last_row = new_last_row
    else:
        last_row = new_last_row

    log(f"Rango de datos actualizado: filas {start_data_row} a {last_row} ({last_row - start_data_row + 1} registros)")

    # Actualizar el rango usado en la hoja para asegurar que Excel lo reconozca
    try:
        ws_inv_copia.UsedRange.Calculate()
    except Exception as e:
        log(f"Aviso: no se pudo recalcular UsedRange: {e}")

    # APLICAR BORDES A TODO EL RANGO
    log("Aplicando bordes a todo el rango de datos...")
    try:
        used_range = ws_inv_copia.UsedRange
        first_col = used_range.Column
        last_col = first_col + used_range.Columns.Count - 1
        ws_apply_borders_to_range(ws_inv_copia, header_row_used, last_row, first_col, last_col)
        
    except Exception as e:
        log(f"⚠ Error al aplicar bordes: {e}")
        import traceback
        log(traceback.format_exc())
    log("Limpiando formato no deseado...")
    try:
        used_range = ws_inv_copia.UsedRange
        first_col = used_range.Column
        last_col = first_col + used_range.Columns.Count - 1
        ws_remove_formatting_from_range(ws_inv_copia, start_data_row, last_row, first_col, last_col)
        
    except Exception as e:
        log(f"⚠ Error al limpiar formato: {e}")

    # ACTUALIZAR FÓRMULAS DE SUBTOTAL EN FILA 1
    log("Actualizando fórmulas de subtotal en fila 1...")
    try:
        ws_update_subtotal_formula(ws_inv_copia, 1, last_row)
    except Exception as e:
        log(f"⚠ Error al actualizar fórmulas de subtotal: {e}")

    # 10) Arrastrar fórmulas en INVENTARIO COPIA
    log("Arrastrando fórmulas en INVENTARIO COPIA...")
    for colname in ["Dif marca", "Dif linea", "Dif sub-linea"]:
        cidx = hdrn_copia.get(_norm(colname))
        if cidx:
            ws_copy_down_formula(ws_inv_copia, cidx, start_data_row, last_row)

    col_total_inv = hdrn_copia.get(_norm("TOTAL INV"))
    if col_total_inv:
        ws_copy_down_formula(ws_inv_copia, col_total_inv, start_data_row, last_row)

    col_exist = ws_ensure_existencia_header(ws_inv_copia, header_row_used)
    ws_copy_down_formula(ws_inv_copia, col_exist, start_data_row, last_row)


    # 11)Actualizar NOMBRE LISTA desde Matriz USD
    log("Actualizando NOMBRE LISTA desde Matriz USD...")
    if len(matriz_map) > 0:
        try:
            col_nombre_lista = hdrn_copia.get(_norm("NOMBRE LISTA"))
            if col_nombre_lista:
                refs_copia = read_range_as_array(ws_inv_copia, start_data_row, last_row, ref_col_idx)
                refs_copia = [to_num_str(r) for r in refs_copia]
                
                descripciones = []
                matched_count = 0
                for ref in refs_copia:
                    if ref in matriz_map:
                        desc = matriz_map[ref]
                        # Si hay descripción, usarla; si no, poner "0"
                        descripciones.append(desc if desc else "0")
                        if desc:
                            matched_count += 1
                    else:
                        # Si no hay coincidencia, poner "0"
                        descripciones.append("0")
                
                write_range_as_array(ws_inv_copia, start_data_row, col_nombre_lista, descripciones)
                log(f"{matched_count} descripciones actualizadas desde Matriz USD")
            else:
                log("  ⚠ Columna 'NOMBRE LISTA' no encontrada en INVENTARIO COPIA")
        except Exception as e:
            log(f"  ⚠ Error al actualizar NOMBRE LISTA: {e}")
            import traceback
            log(traceback.format_exc())
    else:
        log("  ⚠ No hay datos de Matriz USD disponibles - saltando actualización de NOMBRE LISTA")

    # 11.5) Llenar NOMBRE MYR con prioridad NOMBRE LISTA -> NOMBRE ODOO
    log("Actualizando NOMBRE MYR (prioridad: NOMBRE LISTA → NOMBRE ODOO)...")
    try:
        col_nombre_myr = hdrn_copia.get(_norm("NOMBRE MYR"))
        col_nombre_lista = hdrn_copia.get(_norm("NOMBRE LISTA"))
        col_nombre_odoo = hdrn_copia.get(_norm("NOMBRE ODOO"))
        
        if col_nombre_myr:
            if col_nombre_lista and col_nombre_odoo:
                cols_to_read = [col_nombre_lista, col_nombre_odoo]
                data = read_multiple_columns_optimized(ws_inv_copia, start_data_row, last_row, cols_to_read)
                
                nombres_lista = data.get(col_nombre_lista, [])
                nombres_odoo = data.get(col_nombre_odoo, [])
                
                nombres_myr = []
                from_lista = 0
                from_odoo = 0
                
                for i in range(len(nombres_lista)):
                    lista_val = str(nombres_lista[i]).strip() if nombres_lista[i] not in (None, "", "None", 0) else ""
                    odoo_val = str(nombres_odoo[i]).strip() if nombres_odoo[i] not in (None, "", "None") else ""
                    
                    if lista_val:
                        nombres_myr.append(lista_val)
                        from_lista += 1
                    elif odoo_val:
                        nombres_myr.append(odoo_val)
                        from_odoo += 1
                    else:
                        nombres_myr.append("")
                
                write_range_as_array(ws_inv_copia, start_data_row, col_nombre_myr, nombres_myr)
                log(f"NOMBRE MYR actualizado: {from_lista} desde NOMBRE LISTA, {from_odoo} desde NOMBRE ODOO")
                
            elif col_nombre_lista:
                nombres_lista = read_range_as_array(ws_inv_copia, start_data_row, last_row, col_nombre_lista)
                write_range_as_array(ws_inv_copia, start_data_row, col_nombre_myr, nombres_lista)
                log(f"NOMBRE MYR copiado desde NOMBRE LISTA")
                
            elif col_nombre_odoo:
                nombres_odoo = read_range_as_array(ws_inv_copia, start_data_row, last_row, col_nombre_odoo)
                write_range_as_array(ws_inv_copia, start_data_row, col_nombre_myr, nombres_odoo)
                log(f"NOMBRE MYR copiado desde NOMBRE ODOO")
            else:
                log("  ⚠ No se encontraron columnas NOMBRE LISTA ni NOMBRE ODOO")
        else:
            log("  ⚠ Columna 'NOMBRE MYR' no encontrada en INVENTARIO COPIA")
            
    except Exception as e:
        log(f"  ⚠ Error al actualizar NOMBRE MYR: {e}")
        import traceback
        log(traceback.format_exc())

    # 12) Llevar EXISTENCIA_CALC en INVENTARIO COPIA 
    log("Escribiendo EXISTENCIA consolidada en INVENTARIO COPIA .")
    try:
        refs_copia = read_range_as_array(ws_inv_copia, start_data_row, last_row, ref_col_idx)
        refs_copia = [to_num_str(r) for r in refs_copia]
        
        existencias = []
        valores_encontrados = 0
        for key in refs_copia:
            if key:
                val = exist_map.get(key)
                if pd.notna(val):
                    existencias.append(float(val))
                    valores_encontrados += 1
                else:
                    # Si no hay valor, poner 0
                    existencias.append(0)
            else:
                # Si no hay referencia, poner 0
                existencias.append(0)
        
        write_range_as_array(ws_inv_copia, start_data_row, col_exist, existencias)
        log(f"{valores_encontrados} existencias actualizadas, {len(existencias) - valores_encontrados} con valor 0")
    except Exception as e:
        log(f"⚠ Error al escribir existencias: {e}")

    # 13) Traer columnas desde INVENTARIO ORIGINAL
    log("Trayendo columnas desde INVENTARIO ORIGINAL por REFERENCIA.")
    try:
        hr_orig, hdr_orig, hdrn_orig = ws_headers_smart(ws_inv_orig, HEADER_ROW_INV, ["REFERENCIA"])
        ref_idx_orig = hdrn_orig.get(_norm("REFERENCIA")) or find_reference_col_idx(hdrn_orig, ws_inv_orig, hr_orig)
        
        if ref_idx_orig:
            last_orig = ws_last_row(ws_inv_orig, ref_idx_orig, hr_orig)
            
            pivot_top_orig = ws_first_pivot_row(ws_inv_orig)
            if pivot_top_orig and pivot_top_orig > hr_orig:
                last_orig = min(last_orig, pivot_top_orig - 1)
            
            max_rows = min(last_orig, hr_orig + 50000)
            
            log(f"Leyendo {max_rows - hr_orig} filas desde INVENTARIO ORIGINAL...")
            
            cols_to_read = {ref_idx_orig: "__REF__"}
            for colname in COLS_DESDE_ORIGINAL:
                cidx = hdrn_orig.get(_norm(colname))
                if cidx:
                    cols_to_read[cidx] = colname
            
            if len(cols_to_read) <= 1:
                log("⚠ No hay columnas adicionales para traer")
            else:
                col_indices = sorted(cols_to_read.keys())
                all_data = read_multiple_columns_optimized(ws_inv_orig, hr_orig + 1, max_rows, col_indices)
                
                refs_orig = all_data[ref_idx_orig]
                refs_orig_normalized = [to_num_str(r) for r in refs_orig]
                
                maps = {}
                for col_idx in col_indices:
                    if col_idx == ref_idx_orig:
                        continue
                    colname = cols_to_read[col_idx]
                    maps[colname] = dict(zip(refs_orig_normalized, all_data[col_idx]))
                
                log("Leyendo referencias de INVENTARIO COPIA...")
                refs_copia = read_range_as_array(ws_inv_copia, start_data_row, last_row, ref_col_idx)
                refs_copia_normalized = [to_num_str(r) for r in refs_copia]
                
                log("Construyendo valores a escribir...")
                for colname in COLS_DESDE_ORIGINAL:
                    tgt_idx = hdrn_copia.get(_norm(colname))
                    if not tgt_idx or colname not in maps:
                        continue
                    
                    values_to_write = []
                    matched = 0
                    for ref in refs_copia_normalized:
                        if ref and ref in maps[colname]:
                            val = maps[colname][ref]
                            values_to_write.append(val if val not in (None, "", "None") else "")
                            if val not in (None, "", "None"):
                                matched += 1
                        else:
                            values_to_write.append("")
                    
                    write_range_as_array(ws_inv_copia, start_data_row, tgt_idx, values_to_write)
                
                log(f"Columnas traídas exitosamente desde INVENTARIO ORIGINAL")
                
                try:
                    inv_bodega_idx = hdrn_copia.get(_norm("INV BODEGA GERENCIA"))
                    if inv_bodega_idx:
                        rng = ws_inv_copia.Range(
                            ws_inv_copia.Cells(start_data_row, inv_bodega_idx),
                            ws_inv_copia.Cells(last_row, inv_bodega_idx)
                        )
                        rng.HorizontalAlignment = -4108  # xlCenter
 
                    else:
                        log("  ⚠ Columna INV BODEGA GERENCIA no encontrada")
                except Exception as e:
                    log(f"  ⚠ Error al centrar INV BODEGA GERENCIA: {e}")
                
    except Exception as e:
        log(f"⚠ Error al traer columnas desde original: {e}")
        import traceback
        log(traceback.format_exc())

    # AGREGAR SUBTOTALES FINALES 
    log("Agregando subtotales finales...")
    try:
        ws_add_final_subtotals(ws_inv_copia, last_row, header_row_used, hdrn_copia)
    except Exception as e:
        log(f"⚠ Error al agregar subtotales finales: {e}")

    # Inmovilizar las dos primeras filas
    try:
        # Seleccionar la celda A3 (fila 3, columna 1)
        ws_inv_copia.Cells(3, 1).Select()
        excel.ActiveWindow.FreezePanes = True
    except Exception as e:
        log(f"⚠ Error al inmovilizar paneles: {e}")

    # Aplicar reglas de marcas propias 
    log("Aplicando reglas de negocio para marcas propias...")
    aplicar_reglas_marcas_propias(
        ws_inv_copia, 
        start_data_row, 
        last_row, 
        ref_col_idx, 
        hdrn_copia, 
        marcas_propias, 
        distribucion
    )

    # 14) Llenar REFERENCIA FERTRAC en INV LISTA PRECIOS
    log("Llenando REFERENCIA FERTRAC en INV LISTA PRECIOS desde INVENTARIO COPIA...")
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
                referencias_copia = read_range_as_array(ws_inv_copia, start_data_row, last_row, ref_col_idx)
                referencias_copia = [r for r in referencias_copia if r is not None and str(r).strip()]
                
                last_row_lp = hr_lp + len(referencias_copia)
                write_range_as_array(ws_lp, hr_lp + 1, ref_fertrac_idx, referencias_copia)
                
                try:
                    rng = ws_lp.Range(ws_lp.Cells(hr_lp + 1, ref_fertrac_idx), 
                                     ws_lp.Cells(last_row_lp, ref_fertrac_idx))
                    rng.NumberFormat = "0"
                except Exception as e:
                    log(f"Aviso: no se pudo aplicar formato numérico: {e}")
                
                log(f"{len(referencias_copia)} referencias copiadas a REFERENCIA FERTRAC")
            else:
                log("No se encontró columna REFERENCIA FERTRAC")
        else:
            log("No se encontró la hoja INV LISTA PRECIOS")
            
    except Exception as e:
        log(f"Error al llenar REFERENCIA FERTRAC: {e}")

    # 15) GUARDADO COMO ARCHIVO NUEVO
    log("Preparando guardado del archivo...")

    try:
        ws_count = int(wb.Worksheets.Count)
        has_visible = False
        for i in range(1, ws_count + 1):
            try:
                if int(wb.Worksheets(i).Visible) == -1:
                    has_visible = True
                    break
            except Exception:
                pass
        if not has_visible and ws_count >= 1:
            wb.Worksheets(1).Visible = -1
            wb.Worksheets(1).Activate()
    except Exception:
        pass

    with contextlib.suppress(Exception):
        wb.IsAddin = False
    with contextlib.suppress(Exception):
        wb.Windows(1).Visible = True

    try:
        base_name = OUTPUT_BASENAME
    except NameError:
        base_name = f"{Path(FN_INV_PLANTILLA).stem} (ACTUALIZADO)"
    out_name = f"{base_name} {datetime.now():%Y%m%d_%H%M}.xlsx"
    out_path = BASE_PATH / out_name

    log("Restaurando cálculo automático...")
    try:
        excel.Calculation = -4105  
    except Exception as e:
        log(f"Aviso al restaurar cálculo: {e}")

    log(f"Guardando archivo: {out_name}")
    apply_pw = saveinfo.get("reapply_password")
    if apply_pw:
        wb.SaveAs(str(out_path), FileFormat=51, Password=apply_pw)
    else:
        wb.SaveAs(str(out_path), FileFormat=51)

    log(f"✅ Archivo NUEVO creado: {out_path}")

    excel_close(excel, wb, save=False)

    tmp = saveinfo.get("tmp_path")
    if tmp and os.path.exists(tmp):
        with contextlib.suppress(Exception):
            os.remove(tmp)

    log("== Proceso completado exitosamente ==")


if __name__ == "__main__":
    main()