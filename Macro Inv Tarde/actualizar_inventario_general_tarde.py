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

PASS_INV = "Compras2027"
PASSWORDS_TRY = ["Compras2026", "Compras2027"]

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

# ==== DEPENDENCIAS Com ====
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

def find_column_flexible(columns, search_terms):
    """
    Busca una columna de manera flexible usando múltiples términos de búsqueda.
    Retorna el nombre de la columna encontrada o None.
    """
    columns_norm = {_norm(str(c)): c for c in columns}
    
    for term in search_terms:
        term_norm = _norm(term)
        # Búsqueda exacta
        if term_norm in columns_norm:
            return columns_norm[term_norm]
        
        # Búsqueda parcial
        for col_norm, col_orig in columns_norm.items():
            if term_norm in col_norm:
                return col_orig
    
    return None

# ==== LECTURA DE ARCHIVOS ====
def _strip_dol_tmp(name: str) -> str:
    base = Path(name).stem.replace("~$", "")
    base = re.sub(r"^\$+", "", base)
    return base

def find_by_prefix(base_dir: Path, prefix: str, exts=(".xlsx",".xlsm",".xls",".csv")) -> Path:
    """
    Busca un archivo por prefijo, priorizando coincidencias exactas.
    """
    pref = _norm(prefix)
    exact_matches = []
    partial_matches = []
    
    for f in base_dir.iterdir():
        if not (f.is_file() and f.suffix.lower() in exts):
            continue
        
        nn = _norm(_strip_dol_tmp(f.name))
        
        # Coincidencia EXACTA (mismo número de palabras y mismo orden)
        if nn == pref:
            exact_matches.append(f)
            continue
        
        # Coincidencia por inicio exacto
        if nn.startswith(pref + " ") or nn.startswith(pref):
            # Verificar que no tenga palabras extra antes del prefijo
            partial_matches.append((f, 1))  # prioridad 1 (alta)
            continue
        
        # Coincidencia parcial (contiene el prefijo)
        if pref in nn:
            partial_matches.append((f, 2))  # prioridad 2 (media)
            continue
        
        # Búsqueda por tokens
        tokens = pref.split()
        if all(t in nn for t in tokens):
            partial_matches.append((f, 3))  # prioridad 3 (baja)
    
    # Priorizar coincidencias exactas
    if exact_matches:
        exact_matches.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        return exact_matches[0]
    
    # Si no hay exactas, usar parciales ordenadas por prioridad
    if partial_matches:
        # Ordenar por prioridad primero, luego por fecha
        partial_matches.sort(key=lambda x: (x[1], -x[0].stat().st_mtime))
        return partial_matches[0][0]
    
    raise FileNotFoundError(f"No encontré archivos para '{prefix}' en {base_dir}")

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
            log(f"  ✓ Abierto con {'contraseña' if pw else 'sin contraseña'}")
            break
        except Exception as e:
            last_err = e
            if wb:
                try: wb.Close(SaveChanges=False)
                except Exception: pass
                wb = None
    
    if not wb:
        try: excel.Quit()
        except Exception: pass
        raise ValueError(f"No se pudo abrir {path.name}: {last_err}")

    tmp = Path(tempfile.gettempdir()) / f"~converted_{path.stem}_{datetime.now().strftime('%H%M%S')}.xlsx"
    try:
        wb.SaveAs(str(tmp), FileFormat=51)
        log(f"  ✓ Convertido a: {tmp.name}")
    except Exception as ex:
        try: wb.Close(SaveChanges=False)
        except Exception: pass
        try: excel.Quit()
        except Exception: pass
        raise RuntimeError(f"Error al guardar convertido: {ex}")
    finally:
        try: wb.Close(SaveChanges=False)
        except Exception: pass
        try: excel.Quit()
        except Exception: pass

    return tmp

def open_excel_file(path: Path, passwords: list[str]) -> tuple[Path, str | None]:
    """
    Intenta abrir un archivo Excel y devuelve (ruta_archivo, contraseña_usada).
    """
    last_err = None
    
    # Probar sin contraseña
    try:
        with open(path, "rb") as f:
            office = msoffcrypto.OfficeFile(f)
            if not getattr(office, "is_encrypted", False):
                log(f"  ✓ Archivo sin contraseña")
                return path, None
    except Exception:
        pass
    
    # Probar con contraseñas
    for pw in passwords:
        try:
            bio = decrypt_to_stream(path, pw)
            tmp = save_bytesio_to_temp(bio, path.stem)
            log(f"  ✓ Desencriptado con contraseña")
            return tmp, pw
        except Exception as e:
            last_err = e
    
    # Si falla, intentar conversión COM
    if HAS_COM:
        try:
            log("  → Intentando conversión COM...")
            tmp = com_convert_to_xlsx(path, passwords)
            return tmp, None
        except Exception as e:
            log(f"  ✗ Conversión COM falló: {e}")
    
    raise ValueError(f"No se pudo abrir {path.name}: {last_err}")

def excel_open_workbook(path: Path, passwords: list[str], excel=None):
    """
    Abre un workbook en Excel COM.
    """
    if excel is None:
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
    
    wb = None
    last_err = None
    
    # Intentar sin contraseña
    try:
        wb = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False, 
                                 IgnoreReadOnlyRecommended=True)
        log(f"  ✓ Abierto sin contraseña")
        return excel, wb, None
    except Exception as e:
        last_err = e
        if wb:
            try: wb.Close(SaveChanges=False)
            except Exception: pass
            wb = None
    
    # Intentar con contraseñas
    for pw in passwords:
        try:
            wb = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=False, 
                                     IgnoreReadOnlyRecommended=True, Password=pw)
            log(f"  ✓ Abierto con contraseña")
            return excel, wb, pw
        except Exception as e:
            last_err = e
            if wb:
                try: wb.Close(SaveChanges=False)
                except Exception: pass
                wb = None
    
    raise ValueError(f"No se pudo abrir {path.name}: {last_err}")

def excel_close(excel, wb, save=False):
    """
    Cierra el workbook y Excel.
    """
    if wb:
        try:
            wb.Close(SaveChanges=save)
        except Exception as e:
            log(f"  Aviso al cerrar workbook: {e}")
    
    if excel:
        try:
            excel.Quit()
        except Exception as e:
            log(f"  Aviso al cerrar Excel: {e}")

def ws_headers(ws, header_row: int) -> tuple[dict, dict]:
    """
    Lee los encabezados de una hoja y devuelve dos diccionarios:
    - hdr: {nombre_original: columna_index}
    - hdrn: {nombre_normalizado: columna_index}
    """
    used_range = ws.UsedRange
    max_col = used_range.Columns.Count
    
    hdr = {}
    hdrn = {}
    
    for col in range(1, max_col + 1):
        try:
            val = ws.Cells(header_row, col).Value
            if val:
                name = str(val).strip()
                hdr[name] = col
                hdrn[_norm(name)] = col
        except Exception:
            continue
    
    return hdr, hdrn

def ws_last_row(ws, col: int, start_row: int = 1) -> int:
    """
    Encuentra la última fila con datos en una columna específica.
    """
    last_row = ws.Cells(ws.Rows.Count, col).End(-4162).Row  # xlUp = -4162
    return max(last_row, start_row)

def read_range_as_array(ws, start_row: int, end_row: int, col: int) -> list:
    """
    Lee un rango de celdas como un array.
    """
    if start_row > end_row:
        return []
    
    rng = ws.Range(ws.Cells(start_row, col), ws.Cells(end_row, col))
    values = rng.Value
    
    if values is None:
        return []
    elif not isinstance(values, (list, tuple)):
        return [values]
    elif end_row - start_row == 0:
        return [values[0] if isinstance(values, (list, tuple)) else values]
    else:
        return [v[0] if isinstance(v, (list, tuple)) else v for v in values]

def write_range_as_array(ws, start_row: int, col: int, values: list):
    """
    Escribe un array de valores en una columna.
    """
    if not values:
        return
    
    end_row = start_row + len(values) - 1
    rng = ws.Range(ws.Cells(start_row, col), ws.Cells(end_row, col))
    
    # Convertir a formato que Excel espera
    if len(values) == 1:
        rng.Value = values[0]
    else:
        rng.Value = tuple((v,) for v in values)

def load_valorizado_to_df(path: Path, header_row: int) -> pd.DataFrame:
    """
    Carga un archivo VALORIZADO en un DataFrame.
    """
    try:
        # Intentar leer directamente
        df = pd.read_excel(path, sheet_name=0, header=header_row - 1)
        return df
    except Exception as e:
        log(f"  ✗ Error al cargar {path.name}: {e}")
        return pd.DataFrame()

def main():
    log("\n" + "="*70)
    log("🚀 ACTUALIZACIÓN DE EXISTENCIAS Y COSTOS - ARCHIVO NUEVO")
    log("="*70)
    
    # Verificar dependencias
    if not HAS_COM:
        log("❌ ERROR: win32com no está instalado")
        log("   Instalar con: pip install pywin32")
        return
    
    # 1) Buscar archivo de inventario
    log("\n📁 1. Buscando archivo de inventario...")
    try:
        inv_file = find_by_prefix(BASE_PATH, PATRON_INV_FILE)
        log(f"  ✓ Encontrado: {inv_file.name}")
    except FileNotFoundError as e:
        log(f"  ✗ {e}")
        return
    
    # 2) Abrir archivo de inventario
    log("\n🔓 2. Abriendo archivo de inventario...")
    try:
        inv_path, inv_password = open_excel_file(inv_file, PASSWORDS_TRY)
    except Exception as e:
        log(f"  ✗ Error al abrir inventario: {e}")
        return
    
    saveinfo = {
        "tmp_path": inv_path if inv_path != inv_file else None,
        "reapply_password": inv_password
    }
    
    # 3) Abrir en Excel COM
    log("\n📂 3. Abriendo en Excel COM...")
    try:
        excel, wb, _ = excel_open_workbook(inv_path, PASSWORDS_TRY)
    except Exception as e:
        log(f"  ✗ Error al abrir en Excel COM: {e}")
        return
    
    try:
        # Deshabilitar cálculo automático
        try:
            excel.Calculation = -4135  # xlCalculationManual
            log("  ✓ Cálculo automático deshabilitado")
        except Exception as e:
            log(f"  Aviso: No se pudo deshabilitar cálculo: {e}")
        
        # 4) Buscar hoja INVENTARIO
        log("\n🔍 4. Buscando hoja INVENTARIO...")
        ws_inv = None
        for i in range(1, wb.Worksheets.Count + 1):
            sheet_name = wb.Worksheets(i).Name
            if _norm(sheet_name) == _norm(SHEET_INV):
                ws_inv = wb.Worksheets(i)
                log(f"  ✓ Hoja encontrada: '{sheet_name}'")
                break
        
        if not ws_inv:
            log(f"  ✗ No se encontró hoja '{SHEET_INV}'")
            return
        
        # 5) Leer encabezados
        log("\n📋 5. Leyendo encabezados...")
        hdr, hdrn = ws_headers(ws_inv, HEADER_ROW_INV)
        log(f"  ✓ {len(hdr)} columnas encontradas")
        
        # 6) Verificar columnas necesarias
        ref_col = hdrn.get(_norm("REFERENCIA"))
        if not ref_col:
            log("  ✗ No se encontró columna REFERENCIA")
            return
        
        # Buscar columna de existencia
        exist_col = None
        for name, col in hdr.items():
            name_upper = str(name).upper()
            if name_upper.startswith("EXISTENCIA"):
                exist_col = col
                log(f"  ✓ Columna EXISTENCIA: '{name}' (col {col})")
                break
        
        if not exist_col:
            log("  ✗ No se encontró columna EXISTENCIA")
            return
        
        # Buscar columna de costo
        costo_col = hdrn.get(_norm("COSTO PROMEDIO"))
        if not costo_col:
            log("  ✗ No se encontró columna COSTO PROMEDIO")
            return
        
        log(f"  ✓ REFERENCIA: col {ref_col}")
        log(f"  ✓ EXISTENCIA: col {exist_col}")
        log(f"  ✓ COSTO PROMEDIO: col {costo_col}")
        
        # 7) Leer referencias del inventario
        log("\n📖 6. Leyendo referencias de INVENTARIO...")
        last_row_inv = ws_last_row(ws_inv, ref_col, HEADER_ROW_INV)
        start_data_row = HEADER_ROW_INV + 1
        
        refs_inv = read_range_as_array(ws_inv, start_data_row, last_row_inv, ref_col)
        refs_inv_norm = [to_num_str(r) for r in refs_inv]
        
        log(f"  ✓ {len(refs_inv)} referencias leídas")
        
        # 8) Cargar archivos VALORIZADO
        log("\n📥 7. Cargando archivos VALORIZADO...")
        
        # IMPORTANTE: Procesar primero los más específicos para evitar confusiones
        val_files = [
            (PFX_VAL_GENERAL, "VALORIZADO GENERAL"),
            (PFX_VAL_FALT_IMPO, "VALORIZADO FALTANTES IMPO"),  # Primero el más específico
            (PFX_VAL_FALT, "VALORIZADO FALTANTES"),            # Luego el más general
            (PFX_VAL_TOBERIN, "VALORIZADO TOBERIN")
        ]
        
        dfs_val = []
        for prefix, nombre in val_files:
            try:
                val_file = find_by_prefix(BASE_PATH, prefix)
                log(f"  ✓ {nombre}:")
                log(f"    → Archivo: {val_file.name}")
                df = load_valorizado_to_df(val_file, HEADER_ROW_VAL)
                if not df.empty:
                    dfs_val.append(df)
                    log(f"    → Cargado: {len(df)} filas")
            except FileNotFoundError:
                log(f"  ⚠ {nombre}: No encontrado")
        
        if not dfs_val:
            log("  ✗ No se encontraron archivos VALORIZADO")
            return
        
        # 9) Consolidar datos - VERSIÓN CORREGIDA (CON RESTA)
        log("\n🔀 8. Consolidando datos de VALORIZADO...")

        # Recargar archivos de forma separada (no concatenados)
        try:
            val_file_general = find_by_prefix(BASE_PATH, PFX_VAL_GENERAL)
            df_general = load_valorizado_to_df(val_file_general, HEADER_ROW_VAL)
            log(f"  ✓ GENERAL recargado: {len(df_general)} filas")
        except:
            df_general = pd.DataFrame()
            log(f"  ⚠ GENERAL no disponible")

        try:
            val_file_falt_impo = find_by_prefix(BASE_PATH, PFX_VAL_FALT_IMPO)
            df_falt_impo = load_valorizado_to_df(val_file_falt_impo, HEADER_ROW_VAL)
            log(f"  ✓ FALTANTES IMPO recargado: {len(df_falt_impo)} filas")
        except:
            df_falt_impo = pd.DataFrame()

        try:
            val_file_faltantes = find_by_prefix(BASE_PATH, PFX_VAL_FALT)
            df_faltantes = load_valorizado_to_df(val_file_faltantes, HEADER_ROW_VAL)
            log(f"  ✓ FALTANTES recargado: {len(df_faltantes)} filas")
        except:
            df_faltantes = pd.DataFrame()

        try:
            val_file_toberin = find_by_prefix(BASE_PATH, PFX_VAL_TOBERIN)
            df_toberin = load_valorizado_to_df(val_file_toberin, HEADER_ROW_VAL)
            log(f"  ✓ TOBERIN recargado: {len(df_toberin)} filas")
        except:
            df_toberin = pd.DataFrame()

        # Normalizar columnas en cada DataFrame
        for df_temp in [df_general, df_falt_impo, df_faltantes, df_toberin]:
            if not df_temp.empty:
                df_temp.columns = [_norm(str(c)) for c in df_temp.columns]

        if df_general.empty:
            log("  ✗ VALORIZADO GENERAL está vacío")
            return

        col_referencia = find_column_flexible(df_general.columns, 
            ["referencia", "ref", "codigo", "codigo producto"])
        col_existencia = find_column_flexible(df_general.columns, 
            ["existencia en bodega", "existencia bodega", "existencia", "cantidad"])
        col_costo = find_column_flexible(df_general.columns, 
            ["costo promedio", "costo prom", "costo", "precio costo"])

        if not col_referencia or not col_existencia or not col_costo:
            log("  ✗ Columnas requeridas no encontradas")
            return

        log(f"  ✓ Columnas: REFERENCIA='{col_referencia}', EXISTENCIA='{col_existencia}'")

        # Normalizar referencias
        if not df_general.empty:
            df_general["ref_norm"] = df_general[col_referencia].apply(to_num_str)
        if not df_falt_impo.empty:
            df_falt_impo["ref_norm"] = df_falt_impo[col_referencia].apply(to_num_str)
        if not df_faltantes.empty:
            df_faltantes["ref_norm"] = df_faltantes[col_referencia].apply(to_num_str)
        if not df_toberin.empty:
            df_toberin["ref_norm"] = df_toberin[col_referencia].apply(to_num_str)

        # Consolidar cantidades a restar
        log("\n📊 Consolidando cantidades a restar...")
        log("  Fórmula: GENERAL - (FALTANTES IMPO + FALTANTES + TOBERIN)")

        resta_map = {}

        if not df_falt_impo.empty:
            for _, row in df_falt_impo.iterrows():
                ref = row.get("ref_norm")
                cantidad = float(row.get(col_existencia, 0)) if pd.notna(row.get(col_existencia)) else 0
                if ref is not None and cantidad > 0:
                    resta_map[ref] = resta_map.get(ref, 0) + cantidad

        if not df_faltantes.empty:
            for _, row in df_faltantes.iterrows():
                ref = row.get("ref_norm")
                cantidad = float(row.get(col_existencia, 0)) if pd.notna(row.get(col_existencia)) else 0
                if ref is not None and cantidad > 0:
                    resta_map[ref] = resta_map.get(ref, 0) + cantidad

        if not df_toberin.empty:
            for _, row in df_toberin.iterrows():
                ref = row.get("ref_norm")
                cantidad = float(row.get(col_existencia, 0)) if pd.notna(row.get(col_existencia)) else 0
                if ref is not None and cantidad > 0:
                    resta_map[ref] = resta_map.get(ref, 0) + cantidad

        log(f"  ✓ {len(resta_map)} referencias con cantidades a restar")

        # Calcular existencia final
        log("\n🧮 Calculando existencia final: GENERAL - RESTAS...")

        exist_map = {}
        costo_map = {}
        restas_aplicadas = 0

        if not df_general.empty:
            for _, row in df_general.iterrows():
                ref = row.get("ref_norm")
                exist_general = float(row.get(col_existencia, 0)) if pd.notna(row.get(col_existencia)) else 0
                costo_val = float(row.get(col_costo, 0)) if pd.notna(row.get(col_costo)) else 0
                
                if ref is not None:
                    cantidad_resta = resta_map.get(ref, 0)
                    existencia_final = exist_general - cantidad_resta
                    
                    exist_map[ref] = existencia_final
                    if costo_val > 0:
                        costo_map[ref] = costo_val
                    
                    if cantidad_resta > 0:
                        restas_aplicadas += 1
                        if restas_aplicadas <= 10:
                            log(f"    Ref {ref}: {exist_general} - {cantidad_resta} = {existencia_final}")

        log(f"  ✓ {len(exist_map)} referencias procesadas")
        log(f"  ✓ {restas_aplicadas} con restas aplicadas")

        negativos = [ref for ref, val in exist_map.items() if val < 0]
        if negativos:
            log(f"  ⚠️ {len(negativos)} referencias con existencia NEGATIVA")

        
        # 10) Actualizar EXISTENCIAS
        log("\n✍️ 8. Actualizando columna EXISTENCIA en INVENTARIO...")
        existencias = []
        matched_exist = 0
        
        for ref_norm in refs_inv_norm:
            if ref_norm and ref_norm in exist_map:
                val = exist_map[ref_norm]
                if pd.notna(val):
                    existencias.append(val)
                    matched_exist += 1
                else:
                    existencias.append(0)
            else:
                existencias.append(0)
        
        write_range_as_array(ws_inv, start_data_row, exist_col, existencias)
        log(f"  ✓ EXISTENCIA actualizada:")
        log(f"    - Total procesado: {len(existencias)}")
        log(f"    - Coincidencias encontradas: {matched_exist}")
        log(f"    - Sin valor: {len(existencias) - matched_exist}")
        

        # 11) Actualizar COSTOS
        log("\n✍️ 8. Actualizando columna COSTO PROMEDIO en INVENTARIO...")

        # Leer existencias actualizadas para determinar cuáles deben ser 0
        existencias_actuales = read_range_as_array(ws_inv, start_data_row, last_row_inv, exist_col)

        costos = []
        matched_costo = 0
        costos_con_valor = 0
        costos_cero_por_existencia = 0

        for i, ref_norm in enumerate(refs_inv_norm):
            # Si la existencia es 0, el costo debe ser 0
            try:
                exist_actual = float(existencias_actuales[i]) if pd.notna(existencias_actuales[i]) else 0
            except:
                exist_actual = 0
            
            if exist_actual == 0:
                costos.append(0)
                costos_cero_por_existencia += 1
            elif ref_norm and ref_norm in costo_map:
                val = costo_map[ref_norm]
                if pd.notna(val) and val != 0:
                    # Mantener valor exacto del VALORIZADO GENERAL (sin redondear)
                    costos.append(float(val))
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
        log(f"    - Valores exactos de VALORIZADO GENERAL (sin redondeo)")

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
                
                # ===== ACTUALIZAR COLUMNA EXISTENCIA + FECHA ACTUAL =====
                log("\n  📊 Actualizando columna EXISTENCIA en INV LISTA PRECIOS...")
                try:
                    # Buscar columna EXISTENCIA en INV LISTA PRECIOS (ej: EXISTENCIA OCT 22)
                    exist_lp_col = None
                    for name, col in hdr_lp.items():
                        name_upper = str(name).upper()
                        if name_upper.startswith("EXISTENCIA"):
                            exist_lp_col = col
                            log(f"    ✓ Columna encontrada en LP: '{name}' (col {col})")
                            break
                    
                    if not exist_lp_col:
                        log("    ⚠ No se encontró columna EXISTENCIA en INV LISTA PRECIOS")
                    else:
                        # CRÍTICO: Después de reordenar, las filas siguen el orden de INVENTARIO
                        # Debemos actualizar EXISTENCIA en el MISMO orden
                        
                        # Crear diccionario: referencia -> existencia (de INVENTARIO)
                        exist_map_lp = {}
                        for i, ref_norm in enumerate(refs_inv_norm):
                            if ref_norm and i < len(existencias):
                                exist_map_lp[ref_norm] = existencias[i]
                        
                        log(f"    ✓ {len(exist_map_lp)} valores de existencia disponibles de INVENTARIO")
                        
                        # Actualizar valores en el MISMO orden que el reordenamiento
                        valores_existencia_lp = []
                        coincidencias = 0
                        
                        # Primero: referencias que están en INVENTARIO (en orden de INVENTARIO)
                        for ref_inv in refs_inv_norm:
                            if ref_inv and ref_inv in refs_encontradas:
                                # Esta referencia está en LP, agregar su existencia
                                val = exist_map_lp.get(ref_inv, 0)
                                if pd.notna(val) and val != "":
                                    valores_existencia_lp.append(val)
                                    if val != 0:
                                        coincidencias += 1
                                else:
                                    valores_existencia_lp.append(0)
                        
                        # Segundo: referencias extra que están en LP pero no en INVENTARIO
                        for ref_extra in refs_extra:
                            valores_existencia_lp.append(0)
                        
                        # Verificar que tenemos el número correcto de valores
                        if len(valores_existencia_lp) != len(filas_ordenadas):
                            log(f"    ⚠ ADVERTENCIA: {len(valores_existencia_lp)} valores vs {len(filas_ordenadas)} filas")
                        
                        # Escribir valores
                        write_range_as_array(ws_lp, start_data_row_lp, exist_lp_col, valores_existencia_lp)
                        
                        log(f"    ✅ EXISTENCIA actualizada en INV LISTA PRECIOS:")
                        log(f"       - Total procesado: {len(valores_existencia_lp)}")
                        log(f"       - Con existencia > 0: {coincidencias}")
                        log(f"       - Con existencia = 0: {len(valores_existencia_lp) - coincidencias}")
                
                except Exception as e:
                    log(f"    ⚠ Error al actualizar EXISTENCIA en LP: {e}")
                    import traceback
                    log(traceback.format_exc())
                
            else:
                log(f"  ⚠ No se encontró columna REFERENCIA FERTRAC")
                log(f"     Columnas disponibles: {list(hdr_lp.keys())}")
                
        except Exception as e:
            log(f"  ⚠ Error al ordenar INV LISTA PRECIOS: {e}")
            import traceback
            log(traceback.format_exc())

        # ===== ACTUALIZACIÓN DE TABLAS DINÁMICAS =====
        log("\n📊 10. Actualizando tablas dinámicas...")
        try:
            # Lista de hojas donde buscar tablas dinámicas
            hojas_con_pivots = [
                "Hoja2",
                "RESUMEN LINEA",
            ]
            
            tablas_actualizadas = 0
            
            for nombre_hoja in hojas_con_pivots:
                try:
                    # Buscar la hoja usando comparación exacta normalizada
                    ws_pivot = None
                    nombre_normalizado = _norm(nombre_hoja)
                    
                    for i in range(1, wb.Worksheets.Count + 1):
                        sheet_name = wb.Worksheets(i).Name
                        if _norm(sheet_name) == nombre_normalizado:
                            ws_pivot = wb.Worksheets(i)
                            log(f"\n  📑 Procesando hoja: '{sheet_name}'")
                            break
                    
                    if not ws_pivot:
                        log(f"  ⚠ Hoja '{nombre_hoja}' no encontrada")
                        continue
                    
                    # Obtener el número de tablas dinámicas en la hoja
                    try:
                        pivot_count = int(getattr(ws_pivot.PivotTables(), "Count", 0))
                    except:
                        pivot_count = 0
                    
                    if pivot_count == 0:
                        log(f"  ⚠ No se encontraron tablas dinámicas en '{sheet_name}'")
                        continue
                    
                    # Actualizar cada tabla dinámica de la hoja
                    log(f"  Actualizando {pivot_count} tabla(s) dinámica(s)...")
                    
                    for j in range(1, pivot_count + 1):
                        try:
                            pivot_table = ws_pivot.PivotTables(j)
                            
                            # Obtener el nombre de la tabla dinámica si existe
                            try:
                                pivot_name = pivot_table.Name
                                log(f"    → Actualizando tabla: {pivot_name}")
                            except:
                                pivot_name = f"Tabla_{j}"
                                log(f"    → Actualizando tabla {j}")
                            
                            # MÉTODO 1: Habilitar actualización automática
                            try:
                                pivot_table.ManualUpdate = False
                            except:
                                pass
                            
                            # MÉTODO 2: Refrescar el PivotCache primero
                            try:
                                pivot_table.PivotCache().Refresh()
                                log(f"      ✓ Cache actualizado")
                            except Exception as e:
                                log(f"      ⚠ No se pudo actualizar cache: {e}")
                            
                            # MÉTODO 3: Refrescar la tabla dinámica
                            try:
                                pivot_table.RefreshTable()
                                log(f"      ✓ Tabla actualizada")
                            except Exception as e:
                                log(f"      ✗ Error al refrescar tabla: {e}")
                                continue
                            
                            # MÉTODO 4: Forzar actualización de datos
                            try:
                                pivot_table.Update()
                                log(f"      ✓ Update() ejecutado")
                            except:
                                pass
                            
                            tablas_actualizadas += 1
                            log(f"    ✅ Tabla '{pivot_name}' actualizada completamente")
                            
                        except Exception as e:
                            log(f"    ✗ Error al actualizar tabla {j}: {e}")
                            import traceback
                            log(f"    {traceback.format_exc()}")
                            
                except Exception as e:
                    log(f"  ✗ Error al procesar hoja '{nombre_hoja}': {e}")
            
            if tablas_actualizadas > 0:
                log(f"\n✅ {tablas_actualizadas} tabla(s) dinámica(s) actualizada(s) exitosamente")
            else:
                log("\n⚠ No se actualizaron tablas dinámicas")
            
        except Exception as e:
            log(f"❌ ERROR al actualizar tablas dinámicas: {e}")
            import traceback
            log(traceback.format_exc())

        # CRÍTICO: GUARDAR y RECALCULAR después de actualizar tablas dinámicas
        if tablas_actualizadas > 0:
            log("\n🔄 Procesando cambios de tablas dinámicas...")
            try:
                # Paso 1: Recalcular TODO el workbook
                log("  1️⃣ Ejecutando cálculo completo del workbook...")
                try:
                    excel.CalculateFull()
                    log("     ✓ Cálculo completo ejecutado")
                except Exception as e:
                    log(f"     ⚠ Error en CalculateFull: {e}")
                
                # Paso 2: Actualizar todos los objetos OLE
                log("  2️⃣ Actualizando objetos del workbook...")
                try:
                    wb.RefreshAll()
                    log("     ✓ RefreshAll ejecutado")
                except Exception as e:
                    log(f"     ⚠ Error en RefreshAll: {e}")
                
                # Paso 3: GUARDAR los cambios en el archivo actual
                log("  3️⃣ Guardando cambios en archivo temporal...")
                wb.Save()
                log("     ✓ Cambios guardados exitosamente")
                
                # Paso 4: Pequeña pausa para asegurar que Excel procese todo
                import time
                time.sleep(0.5)
                log("  ✅ Tablas dinámicas persistidas correctamente")
                
            except Exception as e:
                log(f"  ⚠ Error al guardar cambios: {e}")
                import traceback
                log(traceback.format_exc())

        # 14) Guardar como archivo NUEVO
        log("\n💾 11. Guardando archivo nuevo como copia...")

        
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
        log("\n🔒 12. Cerrando Excel...")
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
