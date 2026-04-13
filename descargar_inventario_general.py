"""
Script para descargar informe de Productos desde ERP Fertrac
- Navega a Inventario > Productos
- Cambia vista a lista
- Detecta total de registros
- Modifica rango para mostrar todos (1-TOTAL)
- Espera que termine de cargar (puede tardar 30+ minutos)

CAMBIOS v2:
- Reintento automático de arranque de Chrome (hasta 3 intentos con espera)
- MODO_HEADLESS = True por defecto para entorno servidor
- Directorio temporal único por ejecución (evita conflictos entre corridas)
- Verificación real de login (no asume éxito por click)
- Manejo del banner de cookies antes del login
- Reintento automático de login si la URL no cambia
- implicitly_wait removido (mezclaba mal con WebDriverWait)
- time.sleep() fijos reemplazados por WebDriverWait en puntos críticos
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from datetime import datetime
import time
import os
import sys
import re
import glob
import shutil
import tempfile
import smtplib
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
URL_INVENTARIO = "https://erp.fertrac.com/web#action=246&model=stock.picking.type&view_type=kanban&menu_id=174"
URL_PRODUCTOS  = "https://erp.fertrac.com/web#action=278&model=product.template&view_type=kanban&menu_id=174"

RUTA_BASE = r"D:\Fertrac\Usuarios\infocompras\ARCHIVOS DIARIOS 2026\INFORMES\INVENTARIO GENERAL ACTUALIZADO"

# CAMBIO: headless=True para entorno servidor Windows sin pantalla
MODO_HEADLESS        = True
TIMEOUT_CARGA_MAXIMA = 3600  # 60 minutos

# ============== CONFIGURACION DE EMAIL ==============
EMAIL_CONFIG = {
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587,
    "sender_email": "data_science@fertrac.com",
    "sender_password": "jprm cfec elhh fvfn",
    "recipient_emails": [
        "analista_automatizacion@fertrac.com",
        "data_science@fertrac.com",
        "asistentecompras@fertrac.com",
        "analistacompras5@fertrac.com",
    ],
    "enabled": True
}

# ============== FUNCIONES AUXILIARES ==============

def obtener_nombre_mes_carpeta():
    meses = {
        1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
        5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
        9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
    }
    mes_actual = datetime.now().month
    return f"{mes_actual:02d}. {meses[mes_actual]}"

def crear_carpeta_mes():
    nombre_carpeta_mes = obtener_nombre_mes_carpeta()
    ruta_completa = os.path.join(RUTA_BASE, nombre_carpeta_mes)
    if not os.path.exists(ruta_completa):
        os.makedirs(ruta_completa, exist_ok=True)
        print(f"[+] Carpeta creada: {nombre_carpeta_mes}")
    else:
        print(f"[+] Carpeta ya existe: {nombre_carpeta_mes}")
    return ruta_completa


# ============== CONFIGURACION DEL DRIVER ==============

def _crear_opciones_chrome(carpeta_descarga):
    """Construye las opciones de Chrome. Separado para reutilizar en reintentos."""
    chrome_options = Options()

    prefs = {
        "download.default_directory": carpeta_descarga,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_settings.popups": 0,
    }
    chrome_options.add_experimental_option("prefs", prefs)

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

    return chrome_options


def configurar_driver(carpeta_descarga, max_intentos=3):
    """
    Configura el driver de Chrome con reintentos.
    CAMBIO PRINCIPAL: si Chrome no arranca (DevToolsActivePort), reintenta
    hasta max_intentos veces esperando entre cada uno.
    """
    print("[*] Configurando Chrome Driver...")

    for intento in range(1, max_intentos + 1):
        try:
            if intento > 1:
                print(f"[*] Reintentando arranque de Chrome (intento {intento}/{max_intentos})...")
                # Espera entre intentos: da tiempo al servidor a liberar recursos
                time.sleep(10)

            chrome_options = _crear_opciones_chrome(carpeta_descarga)
            driver = webdriver.Chrome(options=chrome_options)

            # Timeouts más largos porque esta descarga puede demorar mucho
            driver.set_page_load_timeout(300)
            driver.set_script_timeout(300)
            # CAMBIO: sin implicitly_wait — se usan esperas explícitas por función

            if not MODO_HEADLESS:
                driver.maximize_window()

            print("[OK] Driver configurado correctamente")
            return driver

        except Exception as e:
            print(f"[!] Error arrancando Chrome en intento {intento}: {str(e)}")
            if intento == max_intentos:
                raise Exception(
                    f"Chrome no pudo arrancar después de {max_intentos} intentos. "
                    f"Último error: {str(e)}"
                )
            continue

    return None


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
        pass


def hacer_login(driver, max_intentos=3):
    """
    Realiza el login y VERIFICA que realmente entró al sistema.
    CAMBIO PRINCIPAL: verifica cambio de URL y presencia del menú post-login.
    """
    print("[*] Iniciando sesion...")
    print(f"[*] Abriendo URL: {URL_LOGIN}")

    for intento in range(1, max_intentos + 1):
        if intento > 1:
            print(f"[*] Reintentando login (intento {intento}/{max_intentos})...")

        driver.get(URL_LOGIN)
        wait = WebDriverWait(driver, 20)

        try:
            campo_usuario = wait.until(
                EC.presence_of_element_located((By.NAME, "login"))
            )

            # CAMBIO: aceptar cookies antes de interactuar
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
            try:
                WebDriverWait(driver, 20).until(
                    lambda d: "web/login" not in d.current_url
                )
                print("[OK] Sesion iniciada correctamente")

                # Esperar que el menú principal esté cargado
                WebDriverWait(driver, 30).until(
                    EC.presence_of_element_located((By.XPATH,
                        "//nav | //div[contains(@class,'o_main_navbar')] | //div[contains(@class,'o_menu')]"
                    ))
                )
                print(f"[*] URL post-login: {driver.current_url}")
                return True

            except Exception:
                print(f"[!] Login no completado en intento {intento}, URL: {driver.current_url}")
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
    print(f"[*] URL objetivo: {URL_INVENTARIO}")
    driver.get(URL_INVENTARIO)

    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH,
                "//nav | //div[contains(@class,'o_content')] | //div[contains(@class,'o_kanban')]"
            ))
        )
        print(f"[*] URL actual: {driver.current_url}")
        print("[OK] En la seccion de Inventario")
    except Exception:
        print(f"[!] Timeout esperando inventario, continuando. URL: {driver.current_url}")


def seleccionar_productos(driver):
    """Hace click en 'Datos principales' y luego en 'Productos'."""
    print("[*] Seleccionando 'Datos principales' > 'Productos'...")
    wait = WebDriverWait(driver, 15)

    try:
        # PASO 1: Click en "Datos principales"
        selectores_datos_principales = [
            "//a[contains(text(), 'Datos principales')]",
            "//span[contains(text(), 'Datos principales')]",
            "//div[contains(text(), 'Datos principales')]",
            "//button[contains(., 'Datos principales')]",
        ]

        menu_encontrado = False
        for selector in selectores_datos_principales:
            try:
                menu = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                menu.click()
                menu_encontrado = True
                print("[OK] Menu 'Datos principales' abierto")
                break
            except:
                continue

        if not menu_encontrado:
            raise Exception("No se encontro el menu 'Datos principales'")

        # PASO 2: Click en "Productos"
        selectores_productos = [
            "//a[normalize-space(text())='Productos']",
            "//a[contains(text(), 'Productos') and not(contains(text(), 'Reglas'))]",
            "//span[contains(text(), 'Productos') and not(contains(text(), 'Reglas'))]",
        ]

        for selector in selectores_productos:
            try:
                productos = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                productos.click()
                print("[OK] 'Productos' seleccionado")

                # Verificar URL
                WebDriverWait(driver, 15).until(
                    lambda d: "product.template" in d.current_url
                    or "action=278" in d.current_url
                    or d.current_url != URL_INVENTARIO
                )

                url_actual = driver.current_url
                if "product.template" not in url_actual and "action=278" not in url_actual:
                    print("[!] URL inesperada, navegando directamente a Productos...")
                    driver.get(URL_PRODUCTOS)
                    WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'o_content')]"))
                    )

                print(f"[OK] En Productos. URL: {driver.current_url}")
                return True
            except:
                continue

        raise Exception("No se encontro la opcion 'Productos'")

    except Exception as e:
        print(f"[ERROR] Error seleccionando Productos: {str(e)}")
        try:
            sp = os.path.join(os.path.dirname(os.path.abspath(__file__)), "error_seleccionar_productos.png")
            driver.save_screenshot(sp)
            print(f"[*] Screenshot: {sp}")
        except:
            pass
        return False


def cambiar_a_vista_lista(driver):
    """Cambia la vista a 'Lista'."""
    print("[*] Verificando/cambiando a vista Lista...")

    try:
        time.sleep(2)

        selectores_vista_lista = [
            "//button[contains(@class, 'o_cp_switch_list')]",
            "//button[@data-view-type='list']",
            "//button[contains(@title, 'List')]",
            "//button[contains(@title, 'Lista')]",
            "//i[contains(@class, 'fa-list-ul')]/parent::button",
            "//i[contains(@class, 'oi-view-list')]/parent::button",
            "//button[contains(@class, 'o_list')]",
        ]

        for selector in selectores_vista_lista:
            try:
                botones = driver.find_elements(By.XPATH, selector)
                for boton in botones:
                    if boton.is_displayed():
                        clases = boton.get_attribute("class") or ""
                        aria_pressed = boton.get_attribute("aria-pressed") or ""
                        if "active" in clases or "btn-primary" in clases or aria_pressed == "true":
                            print("[OK] Ya estamos en vista Lista")
                            return True
                        driver.execute_script("arguments[0].scrollIntoView(true);", boton)
                        time.sleep(0.3)
                        try:
                            boton.click()
                        except:
                            driver.execute_script("arguments[0].click();", boton)
                        time.sleep(2)
                        print("[OK] Vista Lista activada")
                        return True
            except:
                continue

        # Fallback: cambiar por URL
        url_actual = driver.current_url
        if "view_type=kanban" in url_actual:
            driver.get(url_actual.replace("view_type=kanban", "view_type=list"))
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, "//table | //div[contains(@class,'o_list')]"))
            )
            print("[OK] Vista lista forzada por URL")
            return True

        print("[!] Asumiendo que ya estamos en vista adecuada")
        return True

    except Exception as e:
        print(f"[ERROR] Error cambiando a vista lista: {str(e)}")
        return False


def detectar_total_registros(driver):
    """Detecta el número total de registros del paginador."""
    print("[*] Detectando total de registros...")

    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH,
                "//span[contains(@class, 'o_pager')] | //div[contains(@class, 'o_cp_pager')]"
            ))
        )

        selectores_paginador = [
            "//span[contains(@class, 'o_pager')]",
            "//div[contains(@class, 'o_cp_pager')]",
            "//*[contains(text(), '/')]",
        ]

        for selector in selectores_paginador:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elem in elementos:
                    texto = elem.text.strip()
                    match = re.search(r'(\d+)\s*/\s*(\d+)', texto)
                    if match:
                        total = int(match.group(2))
                        print(f"[OK] Total de registros detectado: {total}")

                        if total < 10000:
                            print(f"[!] Solo {total} registros, navegando directamente...")
                            driver.get(URL_PRODUCTOS)
                            time.sleep(5)
                            for retry_elem in driver.find_elements(By.XPATH, "//span[contains(@class, 'o_pager')]"):
                                try:
                                    retry_texto = retry_elem.text.strip()
                                    retry_match = re.search(r'(\d+)\s*/\s*(\d+)', retry_texto)
                                    if retry_match:
                                        retry_total = int(retry_match.group(2))
                                        if retry_total >= 10000:
                                            print(f"[OK] Ahora sí: {retry_total} registros")
                                            return retry_total, retry_elem
                                except:
                                    continue
                            raise Exception(f"Vista incorrecta: solo {total} registros. URL: {driver.current_url}")

                        return total, elem
            except:
                continue

        raise Exception("No se pudo detectar el total de registros")

    except Exception as e:
        print(f"[ERROR] Error detectando total: {str(e)}")
        return None, None


def modificar_rango_registros(driver, total):
    """Modifica el campo de rango para mostrar '1-TOTAL'."""
    print(f"[*] Modificando rango a 1-{total}...")

    try:
        time.sleep(1)

        selectores = [
            "//span[contains(@class, 'o_pager_value')]",
            "//span[contains(@class, 'o_pager_limit')]",
            "//*[contains(text(), '1-')]",
        ]

        campo_rango = None
        for selector in selectores:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elem in elementos:
                    if elem.is_displayed():
                        texto = elem.text.strip()
                        if re.match(r'^\d+-\d+$', texto):
                            print(f"[OK] Campo de rango encontrado: '{texto}'")
                            campo_rango = elem
                            break
                if campo_rango:
                    break
            except:
                continue

        if not campo_rango:
            raise Exception("No se encontró el campo de rango")

        driver.execute_script("arguments[0].scrollIntoView(true);", campo_rango)
        time.sleep(0.3)

        actions = ActionChains(driver)
        actions.move_to_element(campo_rango).click().click().click().perform()
        time.sleep(0.3)

        nuevo_rango = f"1-{total}"
        print(f"[*] Escribiendo: {nuevo_rango}")
        actions.send_keys(nuevo_rango).perform()
        time.sleep(0.5)
        actions.send_keys(Keys.RETURN).perform()
        time.sleep(2)

        print(f"[OK] Rango modificado a {nuevo_rango}")
        return True

    except Exception as e:
        print(f"[ERROR] Error modificando rango: {str(e)}")
        try:
            sp = os.path.join(os.path.dirname(os.path.abspath(__file__)), "error_modificar_rango.png")
            driver.save_screenshot(sp)
        except:
            pass
        return False


def esperar_carga_completa(driver, total, timeout=3600):
    """Espera a que terminen de cargar todos los registros."""
    print(f"[*] Esperando carga de {total} registros (máximo {timeout//60} minutos)...")
    print("[*] Este proceso puede tardar 30+ minutos, es normal.")

    tiempo_inicio  = time.time()
    ultimo_reporte = tiempo_inicio

    try:
        while time.time() - tiempo_inicio < timeout:
            tiempo_actual = time.time()

            if tiempo_actual - ultimo_reporte >= 60:
                minutos = int((tiempo_actual - tiempo_inicio) / 60)
                print(f"[*] Esperando carga... {minutos} minuto(s) transcurrido(s)")
                ultimo_reporte = tiempo_actual

            hay_actividad_carga = False

            # Mensajes de carga
            try:
                mensajes = driver.find_elements(By.XPATH,
                    "//*[contains(text(), 'Cargando') or contains(text(), 'Loading') or "
                    "contains(text(), 'Procesando') or contains(text(), 'Espere')]"
                )
                for msg in mensajes:
                    try:
                        if msg.is_displayed() and msg.text.strip():
                            hay_actividad_carga = True
                            break
                    except:
                        continue
            except:
                pass

            if hay_actividad_carga:
                time.sleep(10)
                continue

            # Spinners
            try:
                spinners = driver.find_elements(By.XPATH,
                    "//span[contains(@class, 'fa-spinner')] | //div[contains(@class, 'o_loading')]"
                )
                if any(s.is_displayed() for s in spinners):
                    hay_actividad_carga = True
            except:
                pass

            if hay_actividad_carga:
                time.sleep(5)
                continue

            # Overlay bloqueante
            try:
                overlay = driver.execute_script("""
                    var overlays = document.querySelectorAll(
                        '[class*="blockUI"], [class*="o_loading"], [class*="modal-backdrop"], .o_blockUI, .blockUI'
                    );
                    for (var i = 0; i < overlays.length; i++) {
                        var style = window.getComputedStyle(overlays[i]);
                        if (style.display !== 'none' && style.visibility !== 'hidden') {
                            var rect = overlays[i].getBoundingClientRect();
                            if (rect.width > 500 && rect.height > 500) return true;
                        }
                    }
                    return false;
                """)
                if overlay:
                    time.sleep(10)
                    continue
            except:
                pass

            # Paginador
            try:
                elementos_paginador = driver.find_elements(By.XPATH,
                    "//*[contains(@class, 'o_pager')] | //span[contains(@class, 'o_pager_value')]"
                )
                for elem in elementos_paginador:
                    try:
                        texto = elem.text.strip() or elem.get_attribute('textContent').strip()
                        match = re.search(r'(\d+)\s*-\s*(\d+)\s*/\s*(\d+)', texto)
                        if match:
                            final         = int(match.group(2))
                            total_mostrado = int(match.group(3))
                            if final >= total_mostrado * 0.98:
                                tiempo_total = int((time.time() - tiempo_inicio) / 60)
                                print(f"[OK] Carga completada: {texto} ({tiempo_total} minutos)")
                                return True
                            else:
                                porcentaje = (final / total_mostrado) * 100
                                if tiempo_actual - ultimo_reporte >= 30:
                                    print(f"[*] Progreso: {final}/{total_mostrado} ({porcentaje:.1f}%)")
                    except:
                        continue
            except:
                pass

            time.sleep(5)

        print(f"[!] Timeout ({timeout//60} min)")
        return False

    except Exception as e:
        print(f"[ERROR] Error esperando carga: {str(e)}")
        return False


def seleccionar_todos_registros(driver):
    """Marca el checkbox del header para seleccionar todos los registros."""
    print("[*] Seleccionando todos los registros...")

    try:
        checkbox_header = None
        selectores_checkbox = [
            "//thead//th[1]//input[@type='checkbox']",
            "//th[@class='o_list_record_selector']//input[@type='checkbox']",
            "//table//thead//th//input[@type='checkbox']",
        ]

        for selector in selectores_checkbox:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elem in elementos:
                    if elem.is_displayed():
                        checkbox_header = elem
                        print("[OK] Checkbox del header encontrado")
                        break
                if checkbox_header:
                    break
            except:
                continue

        if not checkbox_header:
            checkbox_header = driver.execute_script("""
                var c = document.querySelector('thead th input[type="checkbox"]');
                if (c) return c;
                return document.querySelector('th.o_list_record_selector input[type="checkbox"]');
            """)
            if checkbox_header:
                print("[OK] Checkbox encontrado con JavaScript")

        if not checkbox_header:
            raise Exception("No se pudo encontrar el checkbox del header")

        driver.execute_script("arguments[0].scrollIntoView(true);", checkbox_header)
        time.sleep(0.3)

        try:
            checkbox_header.click()
            print("[OK] Click en checkbox realizado")
        except Exception as e:
            print(f"[!] Click normal falló: {str(e)}, intentando con JS...")
            try:
                driver.execute_script("arguments[0].click();", checkbox_header)
                print("[OK] Click con JavaScript realizado")
            except Exception as js_error:
                if "timeout" in str(js_error).lower():
                    print("[!] Timeout en JS (normal con muchos registros), continuando...")
                else:
                    raise

        # Esperar procesamiento
        print("[*] Esperando procesamiento de selección (hasta 5 minutos)...")
        for minuto in range(5):
            time.sleep(60)
            print(f"[*] {minuto + 1}/5 minutos esperados...")

        # Verificar botón Acción
        selectores_accion = [
            "//button[contains(., 'Acción')]",
            "//button[contains(., 'Accion')]",
            "//button[contains(text(), 'Acción')]",
        ]

        for selector in selectores_accion:
            try:
                botones = driver.find_elements(By.XPATH, selector)
                for boton in botones:
                    if boton.is_displayed():
                        print(f"[OK] Botón 'Acción' visible: '{boton.text.strip()}'")
                        return True
            except:
                continue

        raise Exception("El botón 'Acción' NO apareció — registros no seleccionados")

    except Exception as e:
        print(f"[ERROR] Error seleccionando registros: {str(e)}")
        try:
            sp = os.path.join(os.path.dirname(os.path.abspath(__file__)), "error_seleccion.png")
            driver.save_screenshot(sp)
            print(f"[*] Screenshot: {sp}")
        except:
            pass
        return False


def abrir_menu_accion(driver):
    """Abre el menú 'Acción'."""
    print("[*] Abriendo menu 'Accion'...")

    selectores_accion = [
        "//button[contains(., 'Acción')]",
        "//button[contains(text(), 'Acción')]",
        "//a[contains(text(), 'Acción')]",
        "//button[contains(., 'Accion')]",
        "//button[contains(text(), 'Accion')]",
        "//button[contains(@class, 'dropdown') and contains(., 'Acc')]",
        "//*[contains(text(), 'Action')]",
    ]

    try:
        for selector in selectores_accion:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for boton in elementos:
                    if boton.is_displayed():
                        texto = boton.text.strip()
                        if 'acci' in texto.lower() or 'action' in texto.lower():
                            driver.execute_script("arguments[0].scrollIntoView(true);", boton)
                            time.sleep(0.3)
                            try:
                                boton.click()
                            except:
                                driver.execute_script("arguments[0].click();", boton)
                            time.sleep(1)
                            print("[OK] Menu 'Accion' abierto")
                            return True
            except:
                continue

        raise Exception("No se encontró el botón 'Acción'")

    except Exception as e:
        print(f"[ERROR] Error abriendo menu Accion: {str(e)}")
        try:
            sp = os.path.join(os.path.dirname(os.path.abspath(__file__)), "error_accion.png")
            driver.save_screenshot(sp)
        except:
            pass
        return False


def seleccionar_exportar(driver):
    """Selecciona 'Exportar' del menú Acción."""
    print("[*] Seleccionando 'Exportar'...")

    selectores_exportar = [
        "//a[contains(text(), 'Exportar')]",
        "//span[contains(text(), 'Exportar')]",
        "//*[contains(text(), 'Exportar') and not(contains(text(), 'fichero'))]",
    ]

    try:
        for selector in selectores_exportar:
            try:
                opcion = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                opcion.click()
                time.sleep(2)
                print("[OK] 'Exportar' seleccionado")
                return True
            except:
                continue

        raise Exception("No se encontró la opción 'Exportar'")

    except Exception as e:
        print(f"[ERROR] Error seleccionando Exportar: {str(e)}")
        return False


def seleccionar_inventario_general(driver):
    """Selecciona 'INVENTARIO GENERAL' del dropdown de exportaciones guardadas."""
    print("[*] Seleccionando 'INVENTARIO GENERAL'...")

    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH,
                "//div[contains(@class,'modal') and contains(@class,'show')] | //div[@role='dialog']"
            ))
        )

        selectores_dropdown = [
            "//select[contains(@name, 'export')]",
            "//div[contains(@class, 'modal')]//select",
            "//select",
        ]

        dropdown = None
        for selector in selectores_dropdown:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elem in elementos:
                    if elem.is_displayed():
                        dropdown = elem
                        print("[OK] Dropdown encontrado")
                        break
                if dropdown:
                    break
            except:
                continue

        if not dropdown:
            dropdown = driver.execute_script("""
                var selects = document.querySelectorAll('select');
                for (var i = 0; i < selects.length; i++) {
                    if (selects[i].offsetParent !== null) return selects[i];
                }
                return null;
            """)
            if dropdown:
                print("[OK] Dropdown encontrado con JavaScript")

        if not dropdown:
            raise Exception("No se encontró el dropdown")

        driver.execute_script("arguments[0].scrollIntoView(true);", dropdown)
        time.sleep(0.3)

        select  = Select(dropdown)
        opciones = select.options
        print(f"[*] Opciones disponibles ({len(opciones)}):")
        for i, opc in enumerate(opciones):
            texto = opc.text.strip()
            print(f"    {i+1}. '{texto}'")
            if "INVENTARIO" in texto.upper() and "GENERAL" in texto.upper():
                print(f"[*] Seleccionando: '{texto}'")
                try:
                    select.select_by_visible_text(texto)
                except:
                    select.select_by_index(i)
                time.sleep(0.5)
                print("[OK] 'INVENTARIO GENERAL' seleccionado")
                return True

        for opc in opciones:
            valor = opc.get_attribute('value') or ''
            if "inventario" in valor.lower() or "general" in valor.lower():
                select.select_by_value(valor)
                time.sleep(0.5)
                print(f"[OK] Seleccionado por valor: '{opc.text.strip()}'")
                return True

        raise Exception("No se encontró 'INVENTARIO GENERAL' en las opciones")

    except Exception as e:
        print(f"[ERROR] Error seleccionando INVENTARIO GENERAL: {str(e)}")
        try:
            sp = os.path.join(os.path.dirname(os.path.abspath(__file__)), "error_inventario_general.png")
            driver.save_screenshot(sp)
        except:
            pass
        return False


def exportar_fichero(driver):
    """Hace click en 'Exportar a fichero'."""
    print("[*] Haciendo click en 'Exportar a fichero'...")

    selectores_boton = [
        "//button[contains(text(), 'Exportar a fichero')]",
        "//button[contains(., 'Exportar') and contains(., 'fichero')]",
        "//button[contains(@class, 'btn-primary') and contains(., 'Exportar')]",
    ]

    try:
        for selector in selectores_boton:
            try:
                boton = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                boton.click()
                time.sleep(2)
                print("[OK] Exportacion iniciada")
                return True
            except:
                continue

        raise Exception("No se encontró el botón 'Exportar a fichero'")

    except Exception as e:
        print(f"[ERROR] Error exportando fichero: {str(e)}")
        return False


def esperar_descarga_archivo(carpeta, timeout=300):
    """Espera a que se complete la descarga del archivo."""
    print(f"[*] Esperando descarga (máximo {timeout//60} minutos)...")

    tiempo_inicio  = time.time()
    ultimo_reporte = tiempo_inicio

    while time.time() - tiempo_inicio < timeout:
        tiempo_actual = time.time()
        if tiempo_actual - ultimo_reporte >= 30:
            print(f"[*] Esperando descarga... {int(tiempo_actual - tiempo_inicio)}s")
            ultimo_reporte = tiempo_actual

        archivos_temp = (
            glob.glob(os.path.join(carpeta, "*.crdownload")) +
            glob.glob(os.path.join(carpeta, "*.tmp"))
        )

        if not archivos_temp:
            todos_archivos = (
                glob.glob(os.path.join(carpeta, "*.xlsx")) +
                glob.glob(os.path.join(carpeta, "*.xls")) +
                glob.glob(os.path.join(carpeta, "*.csv"))
            )
            todos_archivos = [f for f in todos_archivos if not os.path.basename(f).startswith("~$")]

            if todos_archivos:
                archivo_mas_reciente = max(todos_archivos, key=os.path.getmtime)
                if time.time() - os.path.getmtime(archivo_mas_reciente) < 300:
                    print(f"[OK] Archivo descargado: {os.path.basename(archivo_mas_reciente)}")
                    return archivo_mas_reciente

        time.sleep(2)

    print("[!] Timeout esperando la descarga")
    return None


def renombrar_archivo_con_fecha(archivo_original):
    """Renombra el archivo con formato: INVENTARIO GENERAL ACTUALIZADO DD DE MES DE YYYY."""
    print("[*] Renombrando archivo con fecha...")

    try:
        directorio = os.path.dirname(archivo_original)
        extension  = os.path.splitext(archivo_original)[1]

        fecha_actual = datetime.now()
        meses = {
            1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
            5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
            9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
        }

        nuevo_nombre = (
            f"INVENTARIO GENERAL ACTUALIZADO "
            f"{fecha_actual.day:02d} DE {meses[fecha_actual.month]} DE {fecha_actual.year}"
            f"{extension}"
        )
        nueva_ruta = os.path.join(directorio, nuevo_nombre)

        if os.path.exists(nueva_ruta):
            os.remove(nueva_ruta)

        os.rename(archivo_original, nueva_ruta)
        print(f"[OK] Archivo renombrado a: {nuevo_nombre}")
        return nueva_ruta

    except Exception as e:
        print(f"[ERROR] Error renombrando archivo: {str(e)}")
        return archivo_original


def enviar_email_notificacion(exito=True, archivo_descargado=None, tiempo_total=None, error=None):
    """Envía email de notificación al finalizar."""
    if not EMAIL_CONFIG.get("enabled", False):
        return False

    print("[*] Enviando notificacion por email...")

    try:
        msg = MIMEMultipart('alternative')
        msg['From'] = EMAIL_CONFIG['sender_email']
        msg['To']   = ', '.join(EMAIL_CONFIG['recipient_emails'])
        fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        if exito:
            msg['Subject'] = f"✅ FERTRAC: Descarga de Inventario Completada - {datetime.now().strftime('%d/%m/%Y')}"
            html = f"""
            <html><body style="font-family:Arial,sans-serif;">
            <div style="max-width:600px;margin:0 auto;padding:20px;">
                <div style="background:#28a745;color:white;padding:20px;border-radius:5px;">
                    <h2>✅ Descarga de Inventario Completada</h2>
                </div>
                <div style="background:#f8f9fa;padding:20px;border-radius:5px;margin-top:20px;">
                    <p><b>📅 Fecha:</b> {fecha_actual}</p>
                    <p><b>📁 Archivo:</b> {os.path.basename(archivo_descargado) if archivo_descargado else 'N/A'}</p>
                    <p><b>📂 Ubicación:</b> {os.path.dirname(archivo_descargado) if archivo_descargado else 'N/A'}</p>
                    <p><b>⏱️ Tiempo total:</b> {tiempo_total or 'N/A'}</p>
                    <p><b>🎯 Proceso:</b> Descarga de Inventario General (Productos)</p>
                </div>
                <p style="font-size:12px;color:#999;margin-top:20px;">
                    Mensaje automático — descargar_inventario_general.py
                </p>
            </div>
            </body></html>
            """
        else:
            msg['Subject'] = f"❌ FERTRAC: Error en Descarga de Inventario - {datetime.now().strftime('%d/%m/%Y')}"
            html = f"""
            <html><body style="font-family:Arial,sans-serif;">
            <div style="max-width:600px;margin:0 auto;padding:20px;">
                <div style="background:#dc3545;color:white;padding:20px;border-radius:5px;">
                    <h2>❌ Error en Descarga de Inventario</h2>
                </div>
                <div style="background:#f8f9fa;padding:20px;border-radius:5px;margin-top:20px;">
                    <p><b>📅 Fecha:</b> {fecha_actual}</p>
                    <div style="background:#fff3cd;padding:10px;border-left:4px solid #ffc107;margin:15px 0;">
                        <b>Error:</b><br>{error or 'Error desconocido'}
                    </div>
                    <p><b>⚠️ Acción requerida:</b> Revisar logs y screenshots en la carpeta del script</p>
                </div>
                <p style="font-size:12px;color:#999;margin-top:20px;">
                    Mensaje automático — descargar_inventario_general.py
                </p>
            </div>
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

    try:
        print("=" * 70)
        print("AUTOMATIZACION FERTRAC - PRODUCTOS")
        print("=" * 70)
        print(f"Fecha y hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        print(f"Usuario: {USUARIO}")
        print("-" * 70)

        carpeta_descarga = crear_carpeta_mes()
        print(f"Carpeta de descarga: {carpeta_descarga}")
        print("-" * 70)

        # CAMBIO: configurar_driver ahora tiene reintentos incorporados
        driver = configurar_driver(carpeta_descarga)

        if not hacer_login(driver):
            raise Exception("Fallo en el login")

        navegar_a_inventario(driver)

        if not seleccionar_productos(driver):
            raise Exception("Fallo al seleccionar Productos")

        if not cambiar_a_vista_lista(driver):
            raise Exception("Fallo al cambiar a vista lista")

        total, elemento_paginador = detectar_total_registros(driver)
        if total is None:
            raise Exception("Fallo al detectar total de registros")

        if not modificar_rango_registros(driver, total):
            raise Exception("Fallo al modificar rango de registros")

        if not esperar_carga_completa(driver, total, TIMEOUT_CARGA_MAXIMA):
            raise Exception("Fallo al esperar carga completa")

        print("\n" + "=" * 70)
        print("INICIANDO EXPORTACION")
        print("=" * 70)

        if not seleccionar_todos_registros(driver):
            raise Exception("Fallo al seleccionar todos los registros")

        if not abrir_menu_accion(driver):
            raise Exception("Fallo al abrir menu Accion")

        if not seleccionar_exportar(driver):
            raise Exception("Fallo al seleccionar Exportar")

        if not seleccionar_inventario_general(driver):
            raise Exception("Fallo al seleccionar INVENTARIO GENERAL")

        if not exportar_fichero(driver):
            raise Exception("Fallo al exportar fichero")

        # Esperar procesamiento post-exportación
        print("\n" + "=" * 70)
        print("ESPERANDO PROCESAMIENTO DE EXPORTACION")
        print("=" * 70)
        tiempo_export = time.time()
        while time.time() - tiempo_export < 600:
            try:
                spinners = driver.find_elements(By.XPATH,
                    "//span[contains(@class, 'fa-spinner')] | //div[contains(@class, 'o_loading')]"
                )
                if any(s.is_displayed() for s in spinners):
                    minutos = int((time.time() - tiempo_export) / 60)
                    if minutos > 0:
                        print(f"[*] Procesando exportacion... {minutos} minuto(s)")
                    time.sleep(5)
                else:
                    print("[OK] Procesamiento completado")
                    break
            except:
                time.sleep(2)

        # Esperar descarga
        print("\n" + "=" * 70)
        print("ESPERANDO DESCARGA DEL ARCHIVO")
        print("=" * 70)
        archivo_descargado = esperar_descarga_archivo(carpeta_descarga, timeout=300)

        if archivo_descargado:
            archivo_final = renombrar_archivo_con_fecha(archivo_descargado)

            tiempo_total_segundos = int(time.time() - tiempo_inicio_total)
            tiempo_total_texto    = f"{tiempo_total_segundos // 60} minutos"

            print("\n" + "=" * 70)
            print("PROCESO COMPLETADO EXITOSAMENTE")
            print("=" * 70)
            print(f"Archivo: {os.path.basename(archivo_final)}")
            print(f"Ubicacion: {carpeta_descarga}")
            print(f"Tiempo total: {tiempo_total_texto}")
            print("=" * 70)

            enviar_email_notificacion(
                exito=True,
                archivo_descargado=archivo_final,
                tiempo_total=tiempo_total_texto
            )

            print("\n[*] Cerrando navegador en 3 segundos...")
            time.sleep(3)

        else:
            raise Exception("No se pudo verificar la descarga del archivo")

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