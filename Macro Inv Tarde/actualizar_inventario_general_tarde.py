# actualizar_existencias_costos_nuevo_archivo.py
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
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# ==== CONFIGURACIÓN ====
BASE_PATH = Path(__file__).resolve().parent
# BASE_PATH = Path(r"C:\Users\jperez\Desktop\Tecnologia\Inventario General")

PASS_INV = "Compras2025"
PASSWORDS_TRY = ["Compras2025", "Compras2026"]

# Nombre del archivo de salida
OUTPUT_BASENAME = "2025 INVENTARIO GENERAL ACTUALIZADO EXIST-COSTO"
APPLY_PASSWORD_TO_OUTPUT = True  # Si deseas proteger el archivo de salida

def buscar_hoja_por_patron(workbook, patron, ignorar_dolares=True):
    """
    Busca una hoja en el workbook que coincida con el patrón dado.

    """
    for sheet_name in workbook.sheetnames:
        nombre_limpio = sheet_name
        if ignorar_dolares:
            # Remover todos los $ del inicio
            nombre_limpio = re.sub(r'^\$+', '', sheet_name).strip()
        
        if patron in nombre_limpio:
            return sheet_name
    return None

def buscar_archivo_por_patron(directorio, patron, ignorar_dolares=True):
    """
    Busca un archivo en el directorio que coincida con el patrón dado.
    """
    dir_path = Path(directorio)
    for archivo in dir_path.glob("*.xlsx"):
        nombre_limpio = archivo.stem  # nombre sin extensión
        if ignorar_dolares:
            # Remover todos los $ del inicio
            nombre_limpio = re.sub(r'^\$+', '', nombre_limpio).strip()
        
        if patron in nombre_limpio:
            return archivo
    return None

# Prefijos de archivos
PFX_VAL_GENERAL      = "VALORIZADO GENERAL"
PFX_VAL_FALT_IMPO    = "VALORIZADO FALTANTES IMPO"
PFX_VAL_FALT         = "VALORIZADO FALTANTES"
PFX_VAL_TOBERIN      = "VALORIZADO TOBERIN"

PATRON_INV_FILE = "2025 INVENTARIO GENERAL"
SHEET_INV = "INVENTARIO"
HEADER_ROW_INV = 2
HEADER_ROW_LP = 1  # Fila de encabezado para INV LISTA PRECIOS
HEADER_ROW_VAL = 9

# ==== DEPENDENCIAS COM ====
try:
    import win32com.client as win32
    HAS_COM = True
except Exception:
    HAS_COM = False

# ==== UTILIDADES ====
def log(msg): 
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

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
    """Convierte a referencia numérica segura."""
    if pd.isna(x): 
        return ""
    
    if isinstance(x, str):
        s = x.strip()
        if not s:
            return ""
        
        if any(c.isalpha() or c in '()/' for c in s):
            return s
        
        s_clean = s.replace(".", "").replace(",", "")
        try:
            f = float(s_clean)
            if abs(f - int(f)) < 1e-9:
                return str(int(f))
            return str(f)
        except:
            return s
    
    try:
        s = str(x).strip().replace(",", "")
        f = float(s)
        if abs(f - int(f)) < 1e-9:
            return str(int(f))
        return str(f)
    except:
        return str(x).strip()

# ==== LECTURA DE ARCHIVOS ====
def _strip_dol_tmp(name: str) -> str:
    base = Path(name).stem.replace("~$", "")
    base = re.sub(r"^\$+", "", base)
    return base

def find_by_prefix(base_dir: Path, prefix: str, exts=(".xlsx",".xlsm",".xls",".csv")) -> Path:
    pref = _norm(prefix)
    cands = []
    for f in base_dir.iterdir():
        if not (f.is_file() and f.suffix.lower() in exts):
            continue
        nn = _norm(_strip_dol_tmp(f.name))
        if nn.startswith(pref) or pref in nn:
            cands.append(f)
            continue
        tokens = pref.split()
        if all(t in nn for t in tokens):
            cands.append(f)
    if not cands:
        raise FileNotFoundError(f"No encontré archivos para '{prefix}' en {base_dir}")
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

def save_bytesio_to_temp(bio: io.BytesIO, stem: str) -> Path:
    tmp = Path(tempfile.gettempdir()) / f"~dec_{stem}_{datetime.now().strftime('%H%M%S')}.xlsx"
    with open(tmp, "wb") as out:
        out.write(bio.getvalue())
    return tmp

def com_convert_to_xlsx(path: Path, passwords: list[str] | None = None) -> Path:
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
                wb = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True, 
                                         IgnoreReadOnlyRecommended=True, Password=pw)
            else:
                wb = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True, 
                                         IgnoreReadOnlyRecommended=True)
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
    src = open_as_excel_source(path, PASSWORDS_TRY)
    hdr_idx0 = header_row_visible - 1
    chosen = find_sheet_name_flexible_pd(src, targets=(sheet, "INVENTARIO", "INVENTARIO GENERAL", "INV", "Sheet1", "Sheet 1", "Hoja1")) \
             if isinstance(sheet, str) else sheet
    df = pd.read_excel(src, sheet_name=chosen, engine="openpyxl", header=hdr_idx0)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")].copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

# ==== CARGAR VALORIZADOS ====
def cargar_valorizado(base_dir: Path, prefix: str) -> pd.DataFrame:
    """Lee VALORIZADO* con Referencia, Cantidad y Costo Promedio."""
    p = find_by_prefix(base_dir, prefix)
    log(f"  Abriendo: {p.name}")
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
    
    # Buscar columnas
    refc = idx.get("referencia interna") or idx.get("referencia") or idx.get("ref") \
           or next((real for kn, real in idx.items() if "referenc" in kn), None)
    cant = idx.get("cantidad") or next((real for kn, real in idx.items() if kn.startswith("cant")), None)
    
    # Buscar columna de costo promedio
    costo = (idx.get("costo promedio") or idx.get("costo prom") or idx.get("costo") 
             or next((real for kn, real in idx.items() if "costo" in kn and "prom" in kn), None)
             or next((real for kn, real in idx.items() if "costo" in kn), None))

    if not refc: 
        raise KeyError(f"{p.name}: no encuentro 'Referencia interna'. Encabezados: {list(df.columns)}")
    if not cant: 
        raise KeyError(f"{p.name}: no encuentro 'Cantidad'. Encabezados: {list(df.columns)}")

    out = pd.DataFrame()
    out["__REF_INT__"] = df[refc].apply(to_num_str)
    out["__CANT__"] = pd.to_numeric(df[cant], errors="coerce").fillna(0.0)
    
    if costo:
        out["__COSTO__"] = pd.to_numeric(df[costo], errors="coerce").fillna(0.0)
        costos_validos = (out["__COSTO__"] > 0).sum()
        log(f"    ✓ Columna COSTO PROMEDIO encontrada ({costos_validos} valores > 0)")
    else:
        out["__COSTO__"] = 0.0
        log(f"    ⚠ No se encontró columna COSTO PROMEDIO - usando 0")
    
    return out

def cargar_valorizado_desde_ruta(archivo_path: Path) -> pd.DataFrame:
    """Lee VALORIZADO desde ruta específica."""
    log(f"Abriendo: {archivo_path.name}")
    src = open_as_excel_source(archivo_path, PASSWORDS_TRY)

    if archivo_path.suffix.lower() == ".csv":
        df_all = pd.read_csv(src, header=None, dtype=str)
    else:
        df_all = pd.read_excel(src, sheet_name=0, engine="openpyxl", header=None)

    hdr_row0 = HEADER_ROW_VAL - 1
    df = df_all.iloc[hdr_row0:].reset_index(drop=True)
    df.columns = [str(c).strip() for c in df.iloc[0]]
    df = df.iloc[1:].reset_index(drop=True)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]

    idx = {_norm(c): c for c in df.columns}
    refc = idx.get("referencia interna")
    cant = idx.get("cantidad")

    if not refc: raise KeyError(f"{archivo_path.name}: no encuentro 'Referencia interna'")
    if not cant: raise KeyError(f"{archivo_path.name}: no encuentro 'Cantidad'")

    out = pd.DataFrame()
    out["__REF_INT__"] = df[refc].apply(to_num_str)
    out["__CANT__"]    = pd.to_numeric(df[cant], errors="coerce").fillna(0.0)
    
    return out

# ==== FUNCIONES COM EXCEL ====
def excel_open(path: Path, password: str | None = None):
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
        wb = excel.Workbooks.Open(str(src_path), UpdateLinks=0, ReadOnly=False, 
                                  IgnoreReadOnlyRecommended=True)
        try:
            excel.Calculation = -4135
        except Exception as e:
            log(f"  Aviso: no se pudo establecer cálculo manual: {e}")
        return excel, wb, info
    except Exception as e:
        excel.Quit()
        raise RuntimeError(f"No pude abrir el libro {path.name}") from e

def excel_close(excel, wb, save=True):
    try:
        if save:
            excel.Calculation = -4105
        wb.Close(SaveChanges=save)
    finally:
        excel.Quit()

def ws_headers(ws, header_row_visible: int) -> tuple[dict, dict]:
    used_cols = ws.UsedRange.Columns.Count
    hdr = {}
    for c in range(1, used_cols+1):
        v = ws.Cells(header_row_visible, c).Value
        if v is None: 
            continue
        s = str(v).strip()
        if s and s != "None":
            hdr[s] = c
    hdrn = {_norm(k): v for k, v in hdr.items()}
    return hdr, hdrn

def ws_last_row(ws, key_col_idx: int, header_row_visible: int):
    last = ws.Cells(ws.Rows.Count, key_col_idx).End(-4162).Row
    return max(last, header_row_visible)

def read_range_as_array(ws, start_row: int, end_row: int, col_idx: int):
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
    if not values:
        return
    end_row = start_row + len(values) - 1
    rng = ws.Range(ws.Cells(start_row, col_idx), ws.Cells(end_row, col_idx))
    rng.Value = [[v] for v in values]

def ws_clear_column(ws, col_idx: int, start_row: int, end_row: int):
    if end_row < start_row:
        return
    rng = ws.Range(ws.Cells(start_row, col_idx), ws.Cells(end_row, col_idx))
    rng.ClearContents()

# ==== PROCESO PRINCIPAL ====
def main():
    if not HAS_COM:
        raise RuntimeError("Este script requiere Excel COM (win32com).")

    log("="*70)
    log("ACTUALIZACIÓN DE EXISTENCIAS Y COSTOS PROMEDIO")
    log("Genera archivo nuevo sin modificar el original")
    log("="*70)

    # 1) Cargar archivos valorizados
    log("\n📂 1. Cargando archivos valorizados...")
    df_val_gen   = cargar_valorizado(BASE_PATH, PFX_VAL_GENERAL)
    df_val_impo  = cargar_valorizado(BASE_PATH, PFX_VAL_FALT_IMPO)

    # FIX: Buscar VALORIZADO FALTANTES sin IMPO
    log("Buscando VALORIZADO FALTANTES (excluyendo IMPO)...")
    archivo_faltantes = None
    for f in BASE_PATH.iterdir():
        if f.is_file() and f.suffix.lower() in ('.xlsx', '.xlsm'):
            nombre_sin_simbolos = _strip_dol_tmp(f.name)
            nombre_normalizado = _norm(nombre_sin_simbolos)
            
            # Debe ser exactamente "VALORIZADO FALTANTES" (sin IMPO)
            if nombre_normalizado == _norm("VALORIZADO FALTANTES"):
                archivo_faltantes = f
                log(f"  → Encontrado: {f.name}")
                break

    if archivo_faltantes:
        df_val_falt = cargar_valorizado_desde_ruta(archivo_faltantes)
    else:
        log(f"  ⚠️ No encontrado, usando vacío")
        df_val_falt = pd.DataFrame(columns=["__REF_INT__", "__CANT__"])

    df_val_tob   = cargar_valorizado(BASE_PATH, PFX_VAL_TOBERIN)

    # 2) Crear mapas de cantidades
    log("\n🔄 2. Procesando datos de valorizados...")
    val_map_impo = df_val_impo.set_index("__REF_INT__")["__CANT__"]
    val_map_falt = df_val_falt.set_index("__REF_INT__")["__CANT__"]
    val_map_tob = df_val_tob.set_index("__REF_INT__")["__CANT__"]

    # 3) Calcular existencia consolidada en VALORIZADO GENERAL
    df_val_gen = df_val_gen.copy()
    df_val_gen["__IMPO_CANT__"] = df_val_gen["__REF_INT__"].map(val_map_impo).fillna(0.0)
    df_val_gen["__FALT_CANT__"] = df_val_gen["__REF_INT__"].map(val_map_falt).fillna(0.0)
    df_val_gen["__TOB_CANT__"] = df_val_gen["__REF_INT__"].map(val_map_tob).fillna(0.0)

    # FÓRMULA: VALORIZADO GENERAL - FALTANTES IMPO - FALTANTES - TOBERÍN
    df_val_gen["__EXIST_CALC__"] = (
        df_val_gen["__CANT__"] 
        - df_val_gen["__IMPO_CANT__"] 
        - df_val_gen["__FALT_CANT__"] 
        - df_val_gen["__TOB_CANT__"]
    )
    
    # Crear mapas de existencia y costo
    exist_map = df_val_gen.set_index("__REF_INT__")["__EXIST_CALC__"].to_dict()
    costo_map = df_val_gen.set_index("__REF_INT__")["__COSTO__"].to_dict()
    
    # Estadísticas
    exist_positivas = sum(1 for v in exist_map.values() if v > 0)
    costos_positivos = sum(1 for v in costo_map.values() if v > 0)
    
    log(f"  ✓ {len(exist_map)} referencias con existencia calculada ({exist_positivas} positivas)")
    log(f"  ✓ {len(costo_map)} referencias con costo promedio ({costos_positivos} > 0)")

    # 4) Abrir archivo de inventario
    log(f"\n📄 3. Abriendo archivo: {PATRON_INV_FILE}")
    p_inv = find_by_prefix(BASE_PATH, PATRON_INV_FILE)
    if not p_inv.exists():
        raise FileNotFoundError(f"No encontré el archivo: {p_inv}")
    
    excel, wb, saveinfo = excel_open(p_inv, password=PASSWORDS_TRY)

    try:
        # 5) Buscar hoja INVENTARIO
        log("\n🔍 4. Procesando hoja INVENTARIO...")
        ws_inv = None
        for i in range(1, wb.Worksheets.Count + 1):
            sheet_name = wb.Worksheets(i).Name
            if _norm(sheet_name) == _norm(SHEET_INV):
                ws_inv = wb.Worksheets(i)
                log(f"  ✓ Hoja encontrada: '{sheet_name}'")
                break
        
        if not ws_inv:
            raise ValueError(f"No se encontró la hoja '{SHEET_INV}'")

        # 6) Obtener encabezados
        hdr, hdrn = ws_headers(ws_inv, HEADER_ROW_INV)
        
        # Buscar columnas necesarias
        ref_col = hdrn.get(_norm("REFERENCIA")) or hdrn.get(_norm("REFERENCIA FERTRAC"))
        if not ref_col:
            raise KeyError(f"No se encontró columna REFERENCIA. Columnas: {list(hdr.keys())}")
        
        # Buscar columna EXISTENCIA (con fecha)
        exist_col = None
        for name, col in hdr.items():
            if _norm(name).startswith("existencia "):
                exist_col = col
                break
        
        costo_col = hdrn.get(_norm("COSTO PROMEDIO"))
        
        if not exist_col:
            log("  ⚠ No se encontró columna EXISTENCIA - buscando alternativas...")
            exist_col = hdrn.get(_norm("EXISTENCIA"))
        
        if not costo_col:
            log("  ⚠ No se encontró columna COSTO PROMEDIO")
        
        log(f"  ✓ REFERENCIA: columna {ref_col}")
        log(f"  ✓ EXISTENCIA: columna {exist_col if exist_col else 'NO ENCONTRADA'}")
        log(f"  ✓ COSTO PROMEDIO: columna {costo_col if costo_col else 'NO ENCONTRADA'}")

        # 7) Determinar rango de datos
        start_data_row = HEADER_ROW_INV + 1
        last_row = ws_last_row(ws_inv, ref_col, HEADER_ROW_INV)
        log(f"  ✓ Rango de datos: filas {start_data_row} a {last_row} ({last_row - start_data_row + 1} registros)")

        # 8) Actualizar encabezado de EXISTENCIA con fecha actual
        if exist_col:
            target_header = exist_col_title_for_today()
            ws_inv.Cells(HEADER_ROW_INV, exist_col).Value = target_header
            log(f"  ✓ Encabezado actualizado: '{target_header}'")

        # 9) LIMPIAR columnas
        log("\n🧹 5. Limpiando columnas...")
        if exist_col:
            ws_clear_column(ws_inv, exist_col, start_data_row, last_row)
            log(f"  ✓ EXISTENCIA limpiada")
        
        if costo_col:
            ws_clear_column(ws_inv, costo_col, start_data_row, last_row)
            log(f"  ✓ COSTO PROMEDIO limpiado")

        # 10) Leer referencias del inventario
        log("\n📖 6. Leyendo referencias del inventario...")
        refs_inv = read_range_as_array(ws_inv, start_data_row, last_row, ref_col)
        refs_inv_norm = [to_num_str(r) for r in refs_inv]
        log(f"  ✓ {len(refs_inv_norm)} referencias leídas")

        # 11) ACTUALIZAR EXISTENCIA
        existencias = []
        if exist_col:
            log("\n✍️ 7. Actualizando EXISTENCIA...")
            matched = 0
            positivas = 0
            negativas = 0
            ceros = 0
            
            for ref in refs_inv_norm:
                if ref and ref in exist_map:
                    val = exist_map[ref]
                    if pd.notna(val):
                        existencias.append(float(val))
                        matched += 1
                        if val > 0:
                            positivas += 1
                        elif val < 0:
                            negativas += 1
                        else:
                            ceros += 1
                    else:
                        existencias.append(0)
                        ceros += 1
                else:
                    existencias.append(0)
                    ceros += 1
            
            write_range_as_array(ws_inv, start_data_row, exist_col, existencias)
            log(f"  ✓ EXISTENCIA actualizada:")
            log(f"    - Total procesado: {len(existencias)}")
            log(f"    - Coincidencias: {matched}")
            log(f"    - Positivas: {positivas}")
            log(f"    - Negativas: {negativas}")
            log(f"    - Ceros: {ceros}")

        # 12) ACTUALIZAR COSTO PROMEDIO
        if costo_col:
            log("\n💰 8. Actualizando COSTO PROMEDIO...")
            costos = []
            matched_costo = 0
            costos_cero_por_existencia = 0
            costos_con_valor = 0
            
            for i, ref in enumerate(refs_inv_norm):
                # Obtener existencia para aplicar regla
                exist_val = existencias[i] if exist_col and i < len(existencias) else 0
                
                # REGLA: Si existencia es 0, costo es 0
                if exist_val == 0:
                    costos.append(0)
                    costos_cero_por_existencia += 1
                elif ref and ref in costo_map:
                    val = costo_map[ref]
                    if pd.notna(val) and val != 0:
                        # Redondear a 2 decimales (pero NO cambiar formato de Excel)
                        costo_redondeado = round(float(val), 2)
                        costos.append(costo_redondeado)
                        matched_costo += 1
                        costos_con_valor += 1
                    else:
                        costos.append(0)
                else:
                    costos.append(0)
            
            write_range_as_array(ws_inv, start_data_row, costo_col, costos)
            log(f"  ✓ COSTO PROMEDIO actualizado:")
            log(f"    - Total procesado: {len(costos)}")
            log(f"    - Coincidencias encontradas: {matched_costo}")
            log(f"    - Valores con costo > 0: {costos_con_valor}")
            log(f"    - Ceros por existencia=0: {costos_cero_por_existencia}")
            log(f"    - Sin valor: {len(costos) - matched_costo - costos_cero_por_existencia}")
            log(f"    - Valores redondeados a 2 decimales (formato Excel preservado)")

        # 13) ORDENAR INV LISTA PRECIOS SEGÚN EL ORDEN DE INVENTARIO
        log("\n🔄 9. Ordenando INV LISTA PRECIOS según orden de INVENTARIO...")
        try:
            # Buscar la hoja INV LISTA PRECIOS
            ws_lp = None
            for i in range(1, wb.Worksheets.Count + 1):
                sheet_name = wb.Worksheets(i).Name
                sheet_norm = _norm(sheet_name)
                if "inv" in sheet_norm and "lista" in sheet_norm and "precio" in sheet_norm:
                    ws_lp = wb.Worksheets(i)
                    log(f"  ✓ Hoja INV LISTA PRECIOS encontrada: '{sheet_name}'")
                    break
            
            if ws_lp:
                # Obtener encabezados de INV LISTA PRECIOS (fila 1)
                hdr_lp, hdrn_lp = ws_headers(ws_lp, HEADER_ROW_LP)
                
                # Buscar columna REFERENCIA FERTRAC
                ref_fertrac_col = hdrn_lp.get(_norm("REFERENCIA FERTRAC"))
                
                if ref_fertrac_col:
                    # Determinar última fila con datos
                    last_row_lp = ws_last_row(ws_lp, ref_fertrac_col, HEADER_ROW_LP)
                    
                    # Determinar última columna
                    used_range = ws_lp.UsedRange
                    last_col = used_range.Columns.Count
                    
                    log(f"  ✓ Reordenando desde fila {HEADER_ROW_LP + 1} hasta fila {last_row_lp}")
                    log(f"  ✓ Usando orden de INVENTARIO ({len(refs_inv_norm)} referencias)")
                    
                    # Leer todas las referencias de INV LISTA PRECIOS
                    start_data_row_lp = HEADER_ROW_LP + 1
                    refs_lp = read_range_as_array(ws_lp, start_data_row_lp, last_row_lp, ref_fertrac_col)
                    refs_lp_norm = [to_num_str(r) for r in refs_lp]
                    
                    # Leer todas las filas de datos de INV LISTA PRECIOS de una sola vez
                    log(f"  📖 Leyendo {last_row_lp - HEADER_ROW_LP} filas de INV LISTA PRECIOS...")
                    
                    # Leer todo el rango de datos de una vez (mucho más rápido)
                    data_range = ws_lp.Range(
                        ws_lp.Cells(start_data_row_lp, 1),
                        ws_lp.Cells(last_row_lp, last_col)
                    )
                    data_values = data_range.Value
                    
                    # Convertir a lista de filas
                    if data_values is None:
                        filas_lp = []
                    elif not isinstance(data_values, (list, tuple)):
                        # Solo una celda
                        filas_lp = [[data_values]]
                    elif last_row_lp - start_data_row_lp == 0:
                        # Una sola fila
                        filas_lp = [list(data_values) if isinstance(data_values, (list, tuple)) else [data_values]]
                    else:
                        # Múltiples filas
                        filas_lp = [list(row) if isinstance(row, (list, tuple)) else [row] for row in data_values]
                    
                    log(f"  ✅ {len(filas_lp)} filas leídas exitosamente")
                    
                    # Crear diccionario: referencia_normalizada -> fila_completa
                    ref_to_row = {}
                    for i, ref_norm in enumerate(refs_lp_norm):
                        if ref_norm:  # Ignorar referencias vacías
                            ref_to_row[ref_norm] = filas_lp[i]
                    
                    log(f"  🗂️ {len(ref_to_row)} referencias únicas encontradas en INV LISTA PRECIOS")
                    
                    # Reordenar según el orden de INVENTARIO
                    filas_ordenadas = []
                    refs_encontradas = set()
                    refs_no_encontradas = []
                    
                    for ref_inv in refs_inv_norm:
                        if ref_inv and ref_inv in ref_to_row:
                            filas_ordenadas.append(ref_to_row[ref_inv])
                            refs_encontradas.add(ref_inv)
                        else:
                            if ref_inv:
                                refs_no_encontradas.append(ref_inv)
                    
                    # Agregar al final las referencias que están en LP pero no en INVENTARIO
                    refs_extra = []
                    for ref_lp in refs_lp_norm:
                        if ref_lp and ref_lp not in refs_encontradas and ref_lp in ref_to_row:
                            filas_ordenadas.append(ref_to_row[ref_lp])
                            refs_extra.append(ref_lp)
                    
                    log(f"  ✓ Coincidencias: {len(refs_encontradas)}")
                    log(f"  ⚠ Referencias en INVENTARIO sin precio: {len(refs_no_encontradas)}")
                    log(f"  ℹ️ Referencias extra en LISTA PRECIOS: {len(refs_extra)}")
                    
                    # Escribir las filas ordenadas de vuelta a Excel (optimizado)
                    log(f"  ✍️ Escribiendo {len(filas_ordenadas)} filas reordenadas...")
                    
                    if filas_ordenadas:
                        # Escribir todo el rango de una vez (mucho más rápido)
                        write_range = ws_lp.Range(
                            ws_lp.Cells(start_data_row_lp, 1),
                            ws_lp.Cells(start_data_row_lp + len(filas_ordenadas) - 1, last_col)
                        )
                        
                        # Convertir filas_ordenadas a formato que Excel espera
                        if len(filas_ordenadas) == 1:
                            # Una sola fila - necesita ser tupla
                            write_range.Value = tuple(filas_ordenadas[0])
                        else:
                            # Múltiples filas - necesita ser tupla de tuplas
                            write_range.Value = tuple(tuple(fila) for fila in filas_ordenadas)
                    
                    log(f"  ✅ INV LISTA PRECIOS reordenada según INVENTARIO ({len(filas_ordenadas)} filas)")
                else:
                    log(f"  ⚠ No se encontró columna REFERENCIA FERTRAC")
                    log(f"     Columnas disponibles: {list(hdr_lp.keys())}")
            else:
                log(f"  ℹ️ No se encontró hoja INV LISTA PRECIOS (esto es normal si no existe)")
                
        except Exception as e:
            log(f"  ⚠ Error al ordenar INV LISTA PRECIOS: {e}")
            import traceback
            log(traceback.format_exc())


        # 14) Guardar como archivo NUEVO
        log("\n💾 10. Guardando archivo nuevo...")

        
        # Restaurar cálculo automático
        try:
            excel.Calculation = -4105
        except Exception as e:
            log(f"  Aviso al restaurar cálculo: {e}")
        
        # Generar nombre del archivo de salida
        out_name = f"{OUTPUT_BASENAME} {datetime.now():%Y%m%d_%H%M}.xlsx"
        out_path = BASE_PATH / out_name
        
        log(f"  📁 Guardando como: {out_name}")
        
        # Aplicar contraseña si está configurado
        apply_pw = saveinfo.get("reapply_password") if APPLY_PASSWORD_TO_OUTPUT else None
        
        if apply_pw:
            wb.SaveAs(str(out_path), FileFormat=51, Password=apply_pw)
            log(f"  🔒 Archivo protegido con contraseña")
        else:
            wb.SaveAs(str(out_path), FileFormat=51)
        
        log(f"  ✅ Archivo guardado exitosamente: {out_path}")

    finally:
        # 14) Cerrar Excel
        log("\n🔒 10. Cerrando Excel...")
        excel_close(excel, wb, save=False)
        
        # Limpiar archivo temporal si existe
        tmp = saveinfo.get("tmp_path")
        if tmp and os.path.exists(tmp):
            with contextlib.suppress(Exception):
                os.remove(tmp)
                log(f"  ✓ Archivo temporal eliminado")

    log("\n" + "="*70)
    log("✅ PROCESO COMPLETADO EXITOSAMENTE")
    log(f"📂 Archivo generado: {out_name}")
    log(f"📍 Ubicación: {BASE_PATH}")
    log("="*70)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"\n❌ ERROR CRÍTICO: {e}")
        import traceback
        log(traceback.format_exc())
        input("\nPresiona ENTER para salir...")