"""
Script para descargar informes Valorizados por Almacén desde ERP Fertrac
- Navega a Inventario > Informes > Valorizado
- Descarga 4 informes (Fertrac Principal, Toberin, Faltantes, Faltantes_Impo)
- Los renombra automáticamente con nombres específicos
- Los guarda en la carpeta de destino

CAMBIOS v7:
- Verificación real de login (no asume éxito por click)
- Manejo del banner de cookies antes del login
- Reemplazados time.sleep() fijos por WebDriverWait inteligentes
- Reintento automático de login si detecta que sigue en página de login
- Mejor detección del menú Informes con espera explícita
- Headless activado por defecto para entorno servidor
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime
import time
import os
import sys
import glob
import shutil
import smtplib
import tempfile
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configurar encoding UTF-8 para la salida
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except:
        pass

# ============== CONFIGURACION PRINCIPAL ==============

USUARIO = os.getenv("FERTRAC_USER", "consultas")
CLAVE   = os.getenv("FERTRAC_PASS", "Fertrac20231*")

URL_LOGIN      = "https://erp.fertrac.com/web/login"
URL_INVENTARIO = "https://erp.fertrac.com/web?#action=246&model=stock.picking.type&view_type=kanban&menu_id=174"

RUTA_DESCARGA = r"D:\Fertrac\Usuarios\infocompras\ARCHIVOS DIARIOS 2026\Pruebas Inv General\Valorizados"

# CAMBIO: headless=True para entorno servidor Windows sin pantalla
MODO_HEADLESS   = True
TIMEOUT_DESCARGA = 60

ALMACENES_CONFIG = [
    {
        "tipo_consulta": None,
        "ubicacion": None,
        "nombre_archivo": "VALORIZADO GENERAL.xlsx"
    },
    {
        "tipo_consulta": "Ubicación",
        "ubicacion": "3/Aforo Impo",
        "nombre_archivo": "VALORIZADO TOBERIN.xlsx"
    },
    {
        "tipo_consulta": "Ubicación",
        "ubicacion": "4/Faltantes",
        "nombre_archivo": "VALORIZADO FALTANTES.xlsx"
    },
    {
        "tipo_consulta": "Ubicación",
        "ubicacion": "7/Faltantes_Impo",
        "nombre_archivo": "VALORIZADO FALTANTES IMPO.xlsx"
    }
]

# ============== CONFIGURACION DE EMAIL ==============
EMAIL_CONFIG = {
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587,
    "sender_email": "data_science@fertrac.com",
    "sender_password": "jprm cfec elhh fvfn",
    "recipient_emails": [
        "analista_automatizacion@fertrac.com",
        "data_science@fertrac.com",
    ],
    "enabled": True
}

# ============== FUNCIONES AUXILIARES ==============

def crear_carpeta_destino():
    if not os.path.exists(RUTA_DESCARGA):
        os.makedirs(RUTA_DESCARGA, exist_ok=True)
        print(f"[+] Carpeta creada: {RUTA_DESCARGA}")
    else:
        print(f"[+] Carpeta ya existe: {RUTA_DESCARGA}")
    return RUTA_DESCARGA


# ============== CONFIGURACION DEL DRIVER ==============

def configurar_driver(carpeta_descarga):
    print("[*] Configurando Chrome Driver...")

    chrome_options = Options()

    prefs = {
        "download.default_directory": carpeta_descarga,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0,
    }
    chrome_options.add_experimental_option("prefs", prefs)

    # CAMBIO: headless siempre activo en servidor + tamaño explícito
    if MODO_HEADLESS:
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--window-size=1920,1080")
        print("[*] Modo sin ventana activado")

    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--disable-extensions")

    # CAMBIO: directorio temporal único por ejecución para evitar conflictos
    temp_dir = tempfile.mkdtemp()
    chrome_options.add_argument(f"--user-data-dir={temp_dir}")

    driver = webdriver.Chrome(options=chrome_options)
    driver.set_page_load_timeout(120)
    driver.set_script_timeout(120)

    # CAMBIO: sin implicitly_wait global — se usan esperas explícitas por función
    # driver.implicitly_wait(10)  ← REMOVIDO: genera comportamientos impredecibles

    if not MODO_HEADLESS:
        driver.maximize_window()

    print("[OK] Driver configurado correctamente")
    return driver


# ============== FUNCIONES DE AUTOMATIZACION ==============

def _aceptar_cookies_si_existe(driver):
    """Cierra el banner de cookies si está presente. No falla si no existe."""
    try:
        boton = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH,
                "//a[contains(@class,'btn') and contains(translate(., 'ok', 'OK'), 'OK')] | "
                "//button[contains(translate(., 'ok', 'OK'), 'OK') and string-length(normalize-space(.)) <= 5]"
            ))
        )
        boton.click()
        print("[*] Banner de cookies aceptado")
    except:
        pass  # No había banner, continuar normal


def _esta_en_login(driver):
    """Retorna True si la página actual es la de login."""
    return "web/login" in driver.current_url or "login" in driver.current_url


def hacer_login(driver, max_intentos=3):
    """
    Realiza el login y VERIFICA que realmente entró al sistema.
    CAMBIO PRINCIPAL: ya no asume éxito por el click — comprueba la URL post-login.
    """
    print("[*] Iniciando sesion...")

    for intento in range(1, max_intentos + 1):
        if intento > 1:
            print(f"[*] Reintentando login (intento {intento}/{max_intentos})...")

        driver.get(URL_LOGIN)

        wait = WebDriverWait(driver, 20)

        try:
            # CAMBIO: esperar que el campo exista antes de escribir
            campo_usuario = wait.until(
                EC.presence_of_element_located((By.NAME, "login"))
            )

            # CAMBIO: aceptar cookies antes de interactuar con el form
            _aceptar_cookies_si_existe(driver)

            campo_usuario.clear()
            campo_usuario.send_keys(USUARIO)
            print("[OK] Usuario ingresado")

            campo_clave = driver.find_element(By.NAME, "password")
            campo_clave.clear()
            campo_clave.send_keys(CLAVE)
            print("[OK] Contrasena ingresada")

            boton_login = driver.find_element(By.XPATH, "//button[@type='submit']")
            boton_login.click()

            # CAMBIO: esperar hasta que la URL cambie (máx 20s)
            # Si la URL sigue siendo la de login, el intento falló
            try:
                WebDriverWait(driver, 20).until(
                    lambda d: "web/login" not in d.current_url
                )
                print("[OK] Sesion iniciada correctamente")

                # Esperar que el menú principal esté cargado
                print("[*] Esperando carga completa del sistema...")
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH,
                        "//nav | //div[contains(@class,'o_main_navbar')] | //div[contains(@class,'o_menu')]"
                    ))
                )
                return True

            except Exception:
                print(f"[!] Login no completado en intento {intento}, URL actual: {driver.current_url}")
                if intento < max_intentos:
                    time.sleep(3)
                continue

        except Exception as e:
            print(f"[ERROR] Error en login intento {intento}: {str(e)}")
            if intento < max_intentos:
                time.sleep(3)
            continue

    print("[ERROR] Login fallido después de todos los intentos")
    return False


def navegar_a_inventario(driver):
    """Navega a la sección de Inventario y espera que cargue."""
    print("[*] Navegando a Inventario...")
    driver.get(URL_INVENTARIO)

    # CAMBIO: esperar elemento de inventario en lugar de sleep fijo
    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH,
                "//nav | //div[contains(@class,'o_kanban')] | //div[contains(@class,'o_content')]"
            ))
        )
        print("[OK] En la seccion de Inventario")

        # Esperar que el menú de navegación superior esté disponible
        print("[*] Esperando carga completa de Inventario...")
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH,
                "//*[contains(text(),'Informes') or contains(text(),'informes')]"
            ))
        )
        print("[OK] Menu visible, Inventario cargado")

    except Exception:
        # Si no encontró el menú en 20s, igual continúa (puede ser que el selector no matchee)
        print("[!] Timeout esperando menu, continuando de todas formas...")


def abrir_menu_informes(driver):
    """
    Hace click en el menú 'Informes'.
    CAMBIO: usa WebDriverWait en lugar de sleep fijo, y espera que el elemento
    sea clickeable (no solo presente).
    """
    print("[*] Abriendo menu 'Informes'...")

    selectores_informes = [
        "//a[normalize-space(text())='Informes']",
        "//span[normalize-space(text())='Informes']",
        "//li[normalize-space(.)='Informes']/a",
        "//div[normalize-space(text())='Informes']",
        "//button[normalize-space(.)='Informes']",
        "//*[normalize-space(text())='Informes']",
    ]

    try:
        for selector in selectores_informes:
            try:
                print(f"[*] Buscando 'Informes' con: {selector[:70]}...")

                # CAMBIO: espera hasta 15s que el elemento sea clickeable
                elem = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )

                if elem.is_displayed():
                    driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                    time.sleep(0.3)
                    try:
                        elem.click()
                    except:
                        driver.execute_script("arguments[0].click();", elem)

                    print("[OK] Menu 'Informes' abierto")

                    # Esperar que el submenú aparezca
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH,
                            "//*[contains(text(),'Valorizado')]"
                        ))
                    )
                    return True

            except:
                continue

        raise Exception("No se encontro el menu 'Informes'")

    except Exception as e:
        print(f"[ERROR] Error abriendo menu Informes: {str(e)}")
        try:
            screenshot_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "error_informes.png")
            driver.save_screenshot(screenshot_path)
            print(f"[*] Screenshot guardado en: {screenshot_path}")
        except:
            pass
        return False


def seleccionar_valorizado(driver, max_intentos=3):
    """Selecciona la opción 'Valorizado' del menú Informes con reintentos."""
    print("[*] Seleccionando 'Valorizado'...")

    selectores_valorizado = [
        "//a[normalize-space(text())='Valorizado']",
        "//span[normalize-space(text())='Valorizado']",
        "//*[normalize-space(text())='Valorizado']",
    ]

    for intento in range(1, max_intentos + 1):
        print(f"[*] Intento {intento}/{max_intentos}...")

        for selector in selectores_valorizado:
            try:
                print(f"[*] Buscando 'Valorizado' con: {selector[:60]}...")

                elem = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )

                if elem.is_displayed():
                    try:
                        elem.click()
                    except:
                        driver.execute_script("arguments[0].click();", elem)

                    print("[OK] 'Valorizado' seleccionado - Esperando modal...")

                    # Esperar que el modal aparezca
                    WebDriverWait(driver, 15).until(
                        EC.presence_of_element_located((By.XPATH,
                            "//div[contains(@class,'modal') and contains(@class,'show')] | "
                            "//div[@role='dialog']"
                        ))
                    )
                    print("[OK] Modal abierto")
                    return True

            except:
                continue

        if intento < max_intentos:
            print(f"[!] No se encontró 'Valorizado', esperando 5s antes de reintentar...")
            time.sleep(5)

    print("[ERROR] No se encontró la opción 'Valorizado' después de todos los intentos")
    try:
        screenshot_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "error_valorizado.png")
        driver.save_screenshot(screenshot_path)
        print(f"[*] Screenshot guardado en: {screenshot_path}")
    except:
        pass
    return False


def seleccionar_tipo_consulta(driver, tipo_consulta):
    """Selecciona el tipo de consulta especificado."""
    print(f"[*] Seleccionando '{tipo_consulta}' en 'Tipo de consulta'...")

    try:
        # Esperar que el modal esté completamente cargado
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'modal')]//select"))
        )

        selectores_dropdown = [
            "//div[contains(@class, 'modal')]//select",
            "//select",
        ]

        for selector in selectores_dropdown:
            elementos = driver.find_elements(By.XPATH, selector)
            for elem in elementos:
                if elem.is_displayed():
                    try:
                        select = Select(elem)
                        opciones_texto = [opt.text for opt in select.options]
                        if any(tipo_consulta.lower() in opt.lower() for opt in opciones_texto):
                            print("[OK] Dropdown 'Tipo de consulta' encontrado")
                            for opcion in select.options:
                                if tipo_consulta.lower() in opcion.text.lower():
                                    print(f"[*] Seleccionando opcion: '{opcion.text}'")
                                    select.select_by_visible_text(opcion.text)
                                    time.sleep(1)
                                    print(f"[OK] '{tipo_consulta}' seleccionado")
                                    return True
                    except:
                        continue

        raise Exception(f"No se encontro el dropdown o la opcion '{tipo_consulta}'")

    except Exception as e:
        print(f"[ERROR] Error seleccionando tipo consulta: {str(e)}")
        try:
            screenshot_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"error_tipo_consulta_{tipo_consulta}.png")
            driver.save_screenshot(screenshot_path)
        except:
            pass
        return False


def seleccionar_ubicacion_dropdown(driver, nombre_ubicacion):
    """Selecciona una ubicación específica del dropdown 'Ubicación'."""
    print(f"[*] Seleccionando ubicacion '{nombre_ubicacion}'...")

    tiempo_inicio = time.time()
    timeout_total = 30

    try:
        time.sleep(2)

        selects_en_modal = driver.find_elements(By.XPATH, "//div[contains(@class, 'modal')]//select")
        selects_visibles = [s for s in selects_en_modal if s.is_displayed()]
        print(f"[*] {len(selects_visibles)} SELECT(s) visible(s)")

        if len(selects_visibles) == 1:
            print("[*] El campo de Ubicación NO es un SELECT, buscando INPUT...")
            try:
                inputs = driver.find_elements(By.XPATH,
                    "//div[contains(@class, 'modal')]//input[not(@type='hidden') and not(@type='checkbox') and not(@type='radio')]"
                )
                inputs_visibles = [inp for inp in inputs if inp.is_displayed()]
                print(f"[*] {len(inputs_visibles)} INPUT(s) visible(s)")

                campo_ubicacion = inputs_visibles[-1] if len(inputs_visibles) >= 2 else (inputs_visibles[0] if inputs_visibles else None)

                if campo_ubicacion:
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_ubicacion)
                    driver.execute_script("arguments[0].value = '';", campo_ubicacion)
                    time.sleep(0.3)
                    try:
                        campo_ubicacion.click()
                    except:
                        driver.execute_script("arguments[0].click();", campo_ubicacion)
                    time.sleep(0.5)

                    campo_ubicacion.send_keys(nombre_ubicacion)

                    # Esperar opciones del autocomplete
                    selectores_opciones = [
                        "//ul[contains(@class, 'ui-autocomplete')]//li",
                        "//ul[@role='listbox']//li",
                        "//div[contains(@class, 'ui-menu')]//li",
                        "//ul[contains(@class, 'dropdown-menu')]//li",
                        "//ul[contains(@class, 'o_m2o')]//li",
                    ]

                    opciones_encontradas = []
                    for selector in selectores_opciones:
                        if time.time() - tiempo_inicio > timeout_total:
                            break
                        try:
                            WebDriverWait(driver, 3).until(
                                EC.presence_of_element_located((By.XPATH, selector))
                            )
                            opciones = driver.find_elements(By.XPATH, selector)
                            opciones_visibles = [opt for opt in opciones if opt.is_displayed()]
                            if opciones_visibles:
                                opciones_encontradas = opciones_visibles
                                break
                        except:
                            continue

                    if opciones_encontradas:
                        for opcion in opciones_encontradas:
                            try:
                                if opcion.text.strip() == nombre_ubicacion:
                                    opcion.click()
                                    time.sleep(0.5)
                                    print(f"[OK] Ubicación '{nombre_ubicacion}' seleccionada")
                                    return True
                            except:
                                continue
                        # Primera opción como fallback
                        try:
                            opciones_encontradas[0].click()
                            time.sleep(0.5)
                            print("[OK] Primera opción seleccionada (fallback)")
                            return True
                        except:
                            pass

                    # Sin opciones: presionar ENTER
                    campo_ubicacion.send_keys(Keys.RETURN)
                    time.sleep(0.5)
                    print("[OK] ENTER presionado")
                    return True

            except Exception as e:
                print(f"[!] Error buscando INPUT: {e}")

        elif len(selects_visibles) >= 2:
            select_ubicacion = selects_visibles[1]
            select = Select(select_ubicacion)
            for opcion in select.options:
                if opcion.text.strip() == nombre_ubicacion:
                    select.select_by_visible_text(opcion.text.strip())
                    time.sleep(0.5)
                    print(f"[OK] Ubicación '{nombre_ubicacion}' seleccionada")
                    return True
            print(f"[!] '{nombre_ubicacion}' no encontrada en SELECT")
            return False

        print("[ERROR] No se pudo seleccionar la ubicación")
        return False

    except Exception as e:
        print(f"[ERROR] Error después de {time.time()-tiempo_inicio:.1f}s: {str(e)}")
        try:
            screenshot_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"error_ubicacion_{nombre_ubicacion}.png")
            driver.save_screenshot(screenshot_path)
        except:
            pass
        return False


def seleccionar_almacen_dropdown(driver, nombre_almacen):
    """Selecciona un almacén específico del combobox 'Almacén'."""
    print(f"[*] Seleccionando almacen '{nombre_almacen}'...")

    tiempo_inicio = time.time()
    timeout_total = 30

    try:
        time.sleep(2)

        selects_en_modal = driver.find_elements(By.XPATH, "//div[contains(@class, 'modal')]//select")
        selects_visibles = [s for s in selects_en_modal if s.is_displayed()]
        print(f"[*] {len(selects_visibles)} SELECT(s) visible(s)")

        if len(selects_visibles) == 1:
            inputs = driver.find_elements(By.XPATH,
                "//div[contains(@class, 'modal')]//input[not(@type='hidden') and not(@type='checkbox') and not(@type='radio')]"
            )
            inputs_visibles = [inp for inp in inputs if inp.is_displayed()]
            campo_almacen = inputs_visibles[-1] if len(inputs_visibles) >= 2 else (inputs_visibles[0] if inputs_visibles else None)

            if campo_almacen:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_almacen)
                driver.execute_script("arguments[0].value = '';", campo_almacen)
                time.sleep(0.3)
                try:
                    campo_almacen.click()
                except:
                    driver.execute_script("arguments[0].click();", campo_almacen)
                time.sleep(0.5)
                campo_almacen.send_keys(nombre_almacen)

                selectores_opciones = [
                    "//ul[contains(@class, 'ui-autocomplete')]//li",
                    "//ul[@role='listbox']//li",
                    "//div[contains(@class, 'ui-menu')]//li",
                    "//ul[contains(@class, 'dropdown-menu')]//li",
                    "//ul[contains(@class, 'o_m2o')]//li",
                ]

                opciones_encontradas = []
                for selector in selectores_opciones:
                    if time.time() - tiempo_inicio > timeout_total:
                        break
                    try:
                        WebDriverWait(driver, 3).until(
                            EC.presence_of_element_located((By.XPATH, selector))
                        )
                        opciones = driver.find_elements(By.XPATH, selector)
                        opciones_visibles = [opt for opt in opciones if opt.is_displayed()]
                        if opciones_visibles:
                            opciones_encontradas = opciones_visibles
                            break
                    except:
                        continue

                if opciones_encontradas:
                    for opcion in opciones_encontradas:
                        try:
                            if opcion.text.strip() == nombre_almacen:
                                opcion.click()
                                time.sleep(0.5)
                                print(f"[OK] Almacén '{nombre_almacen}' seleccionado")
                                return True
                        except:
                            continue
                    try:
                        opciones_encontradas[0].click()
                        time.sleep(0.5)
                        print("[OK] Primera opción seleccionada (fallback)")
                        return True
                    except:
                        pass

                campo_almacen.send_keys(Keys.RETURN)
                time.sleep(0.5)
                return True

        elif len(selects_visibles) >= 2:
            select = Select(selects_visibles[1])
            for opcion in select.options:
                if opcion.text.strip() == nombre_almacen:
                    select.select_by_visible_text(opcion.text.strip())
                    time.sleep(0.5)
                    print(f"[OK] Almacén '{nombre_almacen}' seleccionado")
                    return True
            return False

        return False

    except Exception as e:
        print(f"[ERROR] Error seleccionando almacen: {str(e)}")
        return False


def generar_xlsx(driver):
    """Hace click en el botón 'Generar XLSX'."""
    print("[*] Haciendo click en 'Generar XLSX'...")

    selectores_boton = [
        "//button[contains(text(), 'Generar XLSX')]",
        "//button[contains(., 'Generar XLSX')]",
        "//button[contains(@class, 'btn-primary') and contains(., 'Generar')]",
    ]

    try:
        for selector in selectores_boton:
            try:
                boton = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                boton.click()
                time.sleep(1)
                print("[OK] 'Generar XLSX' clickeado - Descarga iniciada")
                return True
            except:
                continue
        raise Exception("No se encontro el boton 'Generar XLSX'")
    except Exception as e:
        print(f"[ERROR] Error generando XLSX: {str(e)}")
        return False


def esperar_descarga_archivo(carpeta, timeout=300, archivos_existentes_previos=None):
    """Espera a que se complete la descarga del archivo."""
    print(f"[*] Esperando descarga del archivo (maximo {timeout//60} minutos)...")

    tiempo_inicio = time.time()
    ultimo_reporte = tiempo_inicio

    if archivos_existentes_previos is not None:
        archivos_existentes = archivos_existentes_previos
        print(f"[*] Usando snapshot previo: {len(archivos_existentes)} archivos")
    else:
        archivos_existentes = {}
        try:
            for archivo in glob.glob(os.path.join(carpeta, "*.xlsx")):
                if not os.path.basename(archivo).startswith("~$"):
                    archivos_existentes[archivo] = os.path.getmtime(archivo)
            for archivo in glob.glob(os.path.join(carpeta, "*.xls")):
                if not os.path.basename(archivo).startswith("~$"):
                    archivos_existentes[archivo] = os.path.getmtime(archivo)
        except:
            pass

    while time.time() - tiempo_inicio < timeout:
        tiempo_actual = time.time()
        if tiempo_actual - ultimo_reporte >= 5:
            print(f"[*] Esperando... {int(tiempo_actual - tiempo_inicio)}s transcurridos")
            ultimo_reporte = tiempo_actual

        archivos_temp = glob.glob(os.path.join(carpeta, "*.crdownload"))
        archivos_temp += glob.glob(os.path.join(carpeta, "*.tmp"))
        if archivos_temp:
            time.sleep(1)
            continue

        try:
            archivos_actuales = {}
            for archivo in glob.glob(os.path.join(carpeta, "*.xlsx")):
                if not os.path.basename(archivo).startswith("~$"):
                    archivos_actuales[archivo] = os.path.getmtime(archivo)
            for archivo in glob.glob(os.path.join(carpeta, "*.xls")):
                if not os.path.basename(archivo).startswith("~$"):
                    archivos_actuales[archivo] = os.path.getmtime(archivo)

            for archivo, mtime_actual in archivos_actuales.items():
                if archivo not in archivos_existentes:
                    edad = time.time() - mtime_actual
                    if edad < 90:
                        print(f"[OK] Archivo descargado: {os.path.basename(archivo)}")
                        return archivo
                elif mtime_actual > archivos_existentes[archivo]:
                    edad = time.time() - mtime_actual
                    if edad < 90:
                        print(f"[OK] Archivo descargado (modificado): {os.path.basename(archivo)}")
                        return archivo
        except Exception as e:
            print(f"[!] Error verificando archivos: {e}")

        time.sleep(1)

    print("[!] Timeout esperando la descarga")
    return None


def renombrar_y_mover_archivo(archivo_original, nuevo_nombre, carpeta_destino):
    """Renombra el archivo descargado con el nombre final."""
    print(f"[*] Renombrando archivo a: {nuevo_nombre}")
    try:
        nueva_ruta = os.path.join(carpeta_destino, nuevo_nombre)
        if os.path.exists(nueva_ruta):
            os.remove(nueva_ruta)
        if os.path.dirname(archivo_original) != carpeta_destino:
            shutil.move(archivo_original, nueva_ruta)
        else:
            os.rename(archivo_original, nueva_ruta)
        print(f"[OK] Archivo guardado como: {nuevo_nombre}")
        return nueva_ruta
    except Exception as e:
        print(f"[ERROR] Error renombrando archivo: {str(e)}")
        return archivo_original


def cerrar_modal(driver):
    """Cierra el modal activo."""
    print("[*] Cerrando modal...")
    selectores_cerrar = [
        "//button[contains(@class, 'close')]",
        "//button[@class='btn-close']",
        "//span[contains(@class, 'close')]",
        "//button[contains(@aria-label, 'Close')]",
    ]
    for selector in selectores_cerrar:
        try:
            elementos = driver.find_elements(By.XPATH, selector)
            for elem in elementos:
                if elem.is_displayed():
                    elem.click()
                    time.sleep(1)
                    print("[OK] Modal cerrado")
                    return True
        except:
            continue

    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    time.sleep(1)
    print("[OK] Modal cerrado con ESC")
    return True


# ============== EMAIL ==============

def enviar_email_notificacion(exito=True, archivos_descargados=None, archivos_fallidos=None, tiempo_total=None, error=None):
    if not EMAIL_CONFIG.get("enabled"):
        return False

    try:
        msg = MIMEMultipart('alternative')
        fecha_hora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

        if exito:
            msg['Subject'] = f"✅ Valorizados FERTRAC - Descarga completada {fecha_hora}"
        else:
            msg['Subject'] = f"❌ Valorizados FERTRAC - Error en descarga {fecha_hora}"

        msg['From'] = EMAIL_CONFIG['sender_email']
        msg['To']   = ", ".join(EMAIL_CONFIG['recipient_emails'])

        if exito:
            archivos_html = "".join(
                f"<li>✅ {os.path.basename(a)}</li>"
                for a in (archivos_descargados or [])
            )
            fallidos_html = ""
            if archivos_fallidos:
                fallidos_html = "<h3>⚠️ No descargados:</h3><ul>" + "".join(
                    f"<li>❌ {a}</li>" for a in archivos_fallidos
                ) + "</ul>"

            html = f"""
            <html><body>
            <h2>✅ Descarga de Valorizados completada</h2>
            <p><b>Fecha:</b> {fecha_hora}</p>
            <p><b>Tiempo total:</b> {tiempo_total}</p>
            <h3>Archivos descargados:</h3><ul>{archivos_html}</ul>
            {fallidos_html}
            <hr><p><small>Mensaje automático — descargar_valorizados.py</small></p>
            </body></html>
            """
        else:
            html = f"""
            <html><body>
            <h2>❌ Error en descarga de Valorizados</h2>
            <p><b>Fecha:</b> {fecha_hora}</p>
            <p><b>Error:</b> {error}</p>
            <hr><p><small>Mensaje automático — descargar_valorizados.py</small></p>
            </body></html>
            """

        msg.attach(MIMEText(html, 'html'))

        with smtplib.SMTP(EMAIL_CONFIG['smtp_server'], EMAIL_CONFIG['smtp_port']) as server:
            server.starttls()
            server.login(EMAIL_CONFIG['sender_email'], EMAIL_CONFIG['sender_password'])
            server.send_message(msg)

        print("[OK] Email enviado exitosamente")
        return True

    except Exception as e:
        print(f"[ERROR] Error enviando email: {str(e)}")
        return False


# ============== MAIN ==============

def main():
    driver = None
    tiempo_inicio_total = time.time()
    archivos_descargados = []

    try:
        print("=" * 70)
        print("AUTOMATIZACION FERTRAC - VALORIZADOS POR ALMACEN")
        print("=" * 70)
        print(f"Fecha y hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        print(f"Usuario: {USUARIO}")
        print("-" * 70)

        carpeta_destino = crear_carpeta_destino()
        print(f"Carpeta de destino: {carpeta_destino}")
        print("-" * 70)

        driver = configurar_driver(carpeta_destino)

        if not hacer_login(driver):
            raise Exception("Fallo en el login")

        navegar_a_inventario(driver)

        if not abrir_menu_informes(driver):
            raise Exception("Fallo al abrir menu Informes")

        if not seleccionar_valorizado(driver):
            raise Exception("Fallo al seleccionar Valorizado")

        print("\n" + "=" * 70)
        print("DESCARGANDO INFORMES")
        print("=" * 70)

        for i, config in enumerate(ALMACENES_CONFIG, 1):
            print(f"\n[*] Procesando {i}/{len(ALMACENES_CONFIG)}: {config['nombre_archivo']}")
            print("-" * 70)

            if config['tipo_consulta'] is not None:
                if not seleccionar_tipo_consulta(driver, config['tipo_consulta']):
                    print(f"[!] No se pudo seleccionar tipo consulta '{config['tipo_consulta']}', continuando...")
                    continue
            else:
                print("[*] Usando tipo de consulta predeterminado (Compañía)")
                time.sleep(2)

            if config['tipo_consulta'] == "Ubicación" and config['ubicacion']:
                if not seleccionar_ubicacion_dropdown(driver, config['ubicacion']):
                    print(f"[!] No se pudo seleccionar ubicacion '{config['ubicacion']}', continuando...")
                    continue

            # Snapshot ANTES de la descarga
            archivos_antes_descarga = {}
            try:
                for archivo in glob.glob(os.path.join(carpeta_destino, "*.xlsx")):
                    if not os.path.basename(archivo).startswith("~$"):
                        archivos_antes_descarga[archivo] = os.path.getmtime(archivo)
                for archivo in glob.glob(os.path.join(carpeta_destino, "*.xls")):
                    if not os.path.basename(archivo).startswith("~$"):
                        archivos_antes_descarga[archivo] = os.path.getmtime(archivo)
                print(f"    - {len(archivos_antes_descarga)} archivos en carpeta")
            except:
                pass

            if not generar_xlsx(driver):
                print(f"[!] No se pudo generar XLSX para '{config['nombre_archivo']}', continuando...")
                continue

            archivo_descargado = esperar_descarga_archivo(
                carpeta_destino,
                timeout=TIMEOUT_DESCARGA,
                archivos_existentes_previos=archivos_antes_descarga
            )

            if archivo_descargado:
                time.sleep(2)
                archivo_final = renombrar_y_mover_archivo(
                    archivo_descargado,
                    config['nombre_archivo'],
                    carpeta_destino
                )
                archivos_descargados.append(archivo_final)
                print(f"[OK] {config['nombre_archivo']} descargado y renombrado")
            else:
                print(f"[!] Timeout esperando descarga de '{config['nombre_archivo']}'")
                try:
                    sp = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"error_descarga_{config['nombre_archivo'].replace('.xlsx', '')}.png")
                    driver.save_screenshot(sp)
                    print(f"[*] Screenshot guardado: {sp}")
                except:
                    pass

            if i < len(ALMACENES_CONFIG):
                print("[*] Preparando para siguiente descarga...")
                cerrar_modal(driver)
                time.sleep(1)
                if not abrir_menu_informes(driver):
                    print("[!] No se pudo reabrir menu Informes")
                    break
                if not seleccionar_valorizado(driver):
                    print("[!] No se pudo reabrir Valorizado")
                    break

        tiempo_total_segundos = int(time.time() - tiempo_inicio_total)
        tiempo_total_texto = f"{tiempo_total_segundos // 60} minutos {tiempo_total_segundos % 60} segundos"

        print("\n" + "=" * 70)
        print("PROCESO COMPLETADO")
        print("=" * 70)
        print(f"Archivos descargados: {len(archivos_descargados)}/{len(ALMACENES_CONFIG)}")
        for archivo in archivos_descargados:
            print(f"  - {os.path.basename(archivo)}")
        print(f"Ubicacion: {carpeta_destino}")
        print(f"Tiempo total: {tiempo_total_texto}")
        print("=" * 70)

        nombres_descargados = [os.path.basename(a) for a in archivos_descargados]
        archivos_no_descargados = [
            c['nombre_archivo'] for c in ALMACENES_CONFIG
            if c['nombre_archivo'] not in nombres_descargados
        ]

        enviar_email_notificacion(
            exito=True,
            archivos_descargados=archivos_descargados,
            archivos_fallidos=archivos_no_descargados if archivos_no_descargados else None,
            tiempo_total=tiempo_total_texto
        )

        print("\n[*] Cerrando navegador en 3 segundos...")
        time.sleep(3)

    except Exception as e:
        print(f"\n[ERROR] ERROR: {str(e)}")

        if driver:
            try:
                sp = os.path.join(os.path.dirname(os.path.abspath(__file__)), "error_general.png")
                driver.save_screenshot(sp)
                print(f"[*] Screenshot guardado en: {sp}")
            except:
                pass

        enviar_email_notificacion(exito=False, error=str(e))

        print("\n[*] Cerrando navegador en 3 segundos...")
        time.sleep(3)

    finally:
        if driver:
            driver.quit()
            print("\n[*] Navegador cerrado")

        print(f"[*] Finalizado: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")


# ============== EJECUCION ==============

if __name__ == "__main__":
    main()