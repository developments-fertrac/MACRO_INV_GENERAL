"""
Script para descargar informe de Productos desde ERP Fertrac
VERSION EN DESARROLLO - PARTE 1
- Navega a Inventario > Productos
- Cambia vista a lista
- Detecta total de registros
- Modifica rango para mostrar todos (1-TOTAL)
- Espera que termine de cargar (puede tardar 30+ minutos)
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import time
import os
import sys
import re
import glob
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

# Credenciales ERP Fertrac
USUARIO = os.getenv("FERTRAC_USER", "consultas")
CLAVE = os.getenv("FERTRAC_PASS", "Fertrac20231*")

# URLs
URL_LOGIN = "https://erp.fertrac.com/web/login"
URL_INVENTARIO = "https://erp.fertrac.com/web#action=246&model=stock.picking.type&view_type=kanban&menu_id=174"
URL_PRODUCTOS = "https://erp.fertrac.com/web#action=278&model=product.template&view_type=kanban&menu_id=174"

# Ruta base para descargas
RUTA_BASE = r"D:\Fertrac\Usuarios\infocompras\ARCHIVOS DIARIOS 2026\INFORMES\INVENTARIO GENERAL ACTUALIZADO"

# Configuracion
MODO_HEADLESS = False
TIMEOUT_CARGA_MAXIMA = 3600  # 60 minutos = 3600 segundos

# ============== CONFIGURACION DE EMAIL ==============
EMAIL_CONFIG = {
    "smtp_server": "smtp.gmail.com",
    "smtp_port": 587,
    "sender_email": "analista_automatizacion@fertrac.com",
    "sender_password": "lbih abom caxy pzbh",  # Contraseña de aplicación de Google
    "recipient_emails": [
        "analista_automatizacion@fertrac.com",
        "data_science@fertrac.com,",
        "asistentecompras@fertrac.com,",
        "analistacompras5@fertrac.com",
    ],
    "enabled": True  # Cambiar a False para desactivar correos
}

# ============== FUNCIONES AUXILIARES ==============

def obtener_nombre_mes_carpeta():
    """
    Retorna el nombre de la carpeta del mes actual
    Formato: "01. ENERO", "02. FEBRERO", etc.
    """
    meses = {
        1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
        5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
        9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
    }
    
    mes_actual = datetime.now().month
    nombre_mes = meses[mes_actual]
    
    return f"{mes_actual:02d}. {nombre_mes}"

def crear_carpeta_mes():
    """
    Crea la carpeta del mes si no existe y retorna la ruta completa
    """
    nombre_carpeta_mes = obtener_nombre_mes_carpeta()
    ruta_completa = os.path.join(RUTA_BASE, nombre_carpeta_mes)
    
    if not os.path.exists(ruta_completa):
        os.makedirs(ruta_completa, exist_ok=True)
        print(f"[+] Carpeta creada: {nombre_carpeta_mes}")
    else:
        print(f"[+] Carpeta ya existe: {nombre_carpeta_mes}")
    
    return ruta_completa

# ============== CONFIGURACION DEL DRIVER ==============

def configurar_driver(carpeta_descarga):
    """Configura el driver de Chrome con la carpeta de descarga"""
    print("[*] Configurando Chrome Driver...")
    
    chrome_options = Options()
    
    # Configurar carpeta de descargas
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
    
    driver = webdriver.Chrome(options=chrome_options)
    
    # Configurar timeouts más largos para operaciones con muchos datos
    driver.set_page_load_timeout(300)  # 5 minutos para cargar páginas
    driver.set_script_timeout(300)  # 5 minutos para ejecutar JavaScript
    driver.implicitly_wait(20)  # 20 segundos para encontrar elementos
    
    if not MODO_HEADLESS:
        driver.maximize_window()
    
    print("[OK] Driver configurado correctamente")
    return driver

# ============== FUNCIONES DE AUTOMATIZACION ==============

def hacer_login(driver):
    """Realiza el login en el sistema"""
    print("[*] Iniciando sesion...")
    print(f"[*] Abriendo URL: {URL_LOGIN}")
    driver.get(URL_LOGIN)
    
    wait = WebDriverWait(driver, 20)
    
    try:
        time.sleep(2)
        
        print("[*] Buscando campo de usuario...")
        campo_usuario = wait.until(EC.presence_of_element_located((By.NAME, "login")))
        campo_usuario.clear()
        campo_usuario.send_keys(USUARIO)
        print("[OK] Usuario ingresado")
        
        print("[*] Buscando campo de contrasena...")
        campo_clave = driver.find_element(By.NAME, "password")
        campo_clave.clear()
        campo_clave.send_keys(CLAVE)
        print("[OK] Contrasena ingresada")
        
        print("[*] Haciendo click en boton de login...")
        boton_login = driver.find_element(By.XPATH, "//button[@type='submit']")
        boton_login.click()
        
        time.sleep(3)
        print("[OK] Sesion iniciada correctamente")
        return True
        
    except Exception as e:
        print(f"[ERROR] Error en login: {str(e)}")
        return False

def navegar_a_inventario(driver):
    """Navega a la seccion de Inventario"""
    print("[*] Navegando a Inventario...")
    print(f"[*] URL objetivo: {URL_INVENTARIO}")
    
    driver.get(URL_INVENTARIO)
    time.sleep(5)
    
    print(f"[*] URL actual: {driver.current_url}")
    print("[OK] En la seccion de Inventario")

def seleccionar_productos(driver):
    """Hace click en 'Datos principales' y luego en 'Productos'"""
    print("[*] Seleccionando 'Datos principales' > 'Productos'...")
    wait = WebDriverWait(driver, 10)
    
    try:
        # PASO 1: Hacer click en "Datos principales"
        print("[*] Buscando menu 'Datos principales'...")
        selectores_datos_principales = [
            "//a[contains(text(), 'Datos principales')]",
            "//span[contains(text(), 'Datos principales')]",
            "//div[contains(text(), 'Datos principales')]",
            "//button[contains(., 'Datos principales')]",
        ]
        
        menu_encontrado = False
        for selector in selectores_datos_principales:
            try:
                print(f"[*] Intentando selector: {selector[:50]}...")
                menu = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                menu.click()
                menu_encontrado = True
                time.sleep(2)
                print("[OK] Menu 'Datos principales' abierto")
                break
            except:
                continue
        
        if not menu_encontrado:
            raise Exception("No se encontro el menu 'Datos principales'")
        
        # PASO 2: Hacer click en "Productos"
        print("[*] Buscando opcion 'Productos'...")
        selectores_productos = [
            "//a[contains(text(), 'Productos') and not(contains(text(), 'Reglas'))]",
            "//span[contains(text(), 'Productos') and not(contains(text(), 'Reglas'))]",
            "//div[contains(text(), 'Productos') and not(contains(text(), 'Reglas'))]",
        ]
        
        for selector in selectores_productos:
            try:
                print(f"[*] Intentando selector productos: {selector[:50]}...")
                productos = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                productos.click()
                time.sleep(3)
                print("[OK] 'Productos' seleccionado")
                
                # VALIDACIÓN CRÍTICA: Verificar que la URL contiene "product.template" o "action=278"
                time.sleep(3)  # Esperar a que la URL cambie
                url_actual = driver.current_url
                
                if "product.template" not in url_actual and "action=278" not in url_actual:
                    print(f"[!] ADVERTENCIA: URL no contiene 'product.template' ni 'action=278'")
                    print(f"[!] URL actual: {url_actual}")
                    print("[*] Navegando DIRECTAMENTE a Productos...")
                    
                    # Navegar directamente con la URL correcta
                    driver.get(URL_PRODUCTOS)
                    time.sleep(5)
                    
                    # Verificar nuevamente
                    url_nueva = driver.current_url
                    if "product.template" not in url_nueva and "action=278" not in url_nueva:
                        raise Exception(f"No se pudo navegar a Productos. URL: {url_nueva}")
                    
                    print("[OK] Navegación directa a Productos exitosa")
                
                return True
            except:
                continue
        
        raise Exception("No se encontro la opcion 'Productos'")
        
    except Exception as e:
        print(f"[ERROR] Error seleccionando Productos: {str(e)}")
        
        # Screenshot para debugging
        try:
            screenshot_path = os.path.join(os.path.dirname(__file__), "error_seleccionar_productos.png")
            driver.save_screenshot(screenshot_path)
            print(f"[*] Screenshot guardado en: {screenshot_path}")
        except:
            pass
        
        return False

def cambiar_a_vista_lista(driver):
    """Cambia la vista a 'Lista' si no esta ya en esa vista"""
    print("[*] Verificando/cambiando a vista Lista...")
    wait = WebDriverWait(driver, 10)
    
    try:
        # Dar tiempo a que cargue
        time.sleep(2)
        
        # Buscar el boton de vista lista con múltiples estrategias
        selectores_vista_lista = [
            "//button[contains(@class, 'o_cp_switch_list')]",
            "//button[@data-view-type='list']",
            "//button[contains(@title, 'List')]",
            "//button[contains(@title, 'Lista')]",
            "//i[contains(@class, 'fa-list-ul')]/parent::button",
            "//i[contains(@class, 'oi-view-list')]/parent::button",
            "//button[contains(@class, 'o_list')]",
            "//a[contains(@data-view-type, 'list')]",
            "//button[contains(., 'List')]",
        ]
        
        boton_encontrado = False
        for selector in selectores_vista_lista:
            try:
                print(f"[*] Intentando selector vista lista: {selector[:60]}...")
                botones = driver.find_elements(By.XPATH, selector)
                
                for boton in botones:
                    if boton.is_displayed():
                        print(f"[*] Botón visible encontrado")
                        
                        # Verificar si ya esta activo
                        clases = boton.get_attribute("class") or ""
                        aria_pressed = boton.get_attribute("aria-pressed") or ""
                        
                        print(f"[*] Clases del botón: {clases}")
                        print(f"[*] aria-pressed: {aria_pressed}")
                        
                        if "active" in clases or "btn-primary" in clases or aria_pressed == "true":
                            print("[OK] Ya estamos en vista Lista")
                            return True
                        
                        # Si no esta activo, hacer click
                        print("[*] Haciendo click en botón de vista lista...")
                        try:
                            driver.execute_script("arguments[0].scrollIntoView(true);", boton)
                            time.sleep(0.5)
                            boton.click()
                        except:
                            driver.execute_script("arguments[0].click();", boton)
                        
                        time.sleep(3)
                        print("[OK] Click realizado en vista Lista")
                        boton_encontrado = True
                        return True
            except Exception as e:
                continue
        
        if not boton_encontrado:
            # Estrategia alternativa: buscar por texto visible
            print("[*] Buscando por texto visible 'View list' o 'Lista'...")
            try:
                elementos = driver.find_elements(By.XPATH, "//*[contains(text(), 'View list') or contains(text(), 'Lista') or contains(text(), 'List')]")
                for elem in elementos:
                    if elem.is_displayed():
                        tag = elem.tag_name
                        if tag in ['button', 'a', 'span']:
                            print(f"[*] Encontrado elemento '{tag}' con texto relacionado")
                            try:
                                elem.click()
                                time.sleep(3)
                                print("[OK] Click realizado")
                                return True
                            except:
                                driver.execute_script("arguments[0].click();", elem)
                                time.sleep(3)
                                print("[OK] Click realizado con JS")
                                return True
            except:
                pass
        
        # Si llegamos aquí, no encontramos el botón pero continuamos
        print("[!] No se encontro boton de vista lista")
        
        # Tomar screenshot para ver qué hay en pantalla
        try:
            screenshot_path = os.path.join(os.path.dirname(__file__), "debug_vista_lista.png")
            driver.save_screenshot(screenshot_path)
            print(f"[*] Screenshot guardado en: {screenshot_path}")
        except:
            pass
        
        # Verificar URL actual
        print(f"[*] URL actual: {driver.current_url}")
        
        # Intentar forzar vista lista por URL
        print("[*] Intentando forzar vista lista por URL...")
        url_actual = driver.current_url
        
        # Si la URL tiene view_type=kanban, cambiarla a list
        if "view_type=kanban" in url_actual:
            nueva_url = url_actual.replace("view_type=kanban", "view_type=list")
            print(f"[*] Navegando a: {nueva_url}")
            driver.get(nueva_url)
            time.sleep(5)
            print("[OK] Vista cambiada por URL")
            return True
        
        # Si no tiene view_type, agregarlo
        if "view_type=" not in url_actual:
            if "?" in url_actual:
                nueva_url = url_actual + "&view_type=list"
            else:
                nueva_url = url_actual + "?view_type=list"
            print(f"[*] Navegando a: {nueva_url}")
            driver.get(nueva_url)
            time.sleep(5)
            print("[OK] Vista cambiada por URL")
            return True
        
        print("[!] Asumiendo que ya estamos en vista adecuada")
        return True
        
    except Exception as e:
        print(f"[ERROR] Error cambiando a vista lista: {str(e)}")
        return False

def detectar_total_registros(driver):
    """Detecta el numero total de registros (el numero despues del '/')"""
    print("[*] Detectando total de registros...")
    wait = WebDriverWait(driver, 10)
    
    try:
        # Buscar el elemento que contiene "1-80 / 14760"
        selectores_paginador = [
            "//span[contains(@class, 'o_pager')]",
            "//div[contains(@class, 'o_cp_pager')]",
            "//*[contains(text(), '/')]",
        ]
        
        for selector in selectores_paginador:
            try:
                print(f"[*] Buscando paginador con: {selector[:50]}...")
                elementos = driver.find_elements(By.XPATH, selector)
                
                for elem in elementos:
                    texto = elem.text.strip()
                    print(f"[*] Texto encontrado: '{texto}'")
                    
                    # Buscar patron "X-Y / TOTAL" o "X / TOTAL"
                    match = re.search(r'(\d+)\s*/\s*(\d+)', texto)
                    if match:
                        total = int(match.group(2))
                        print(f"[OK] Total de registros detectado: {total}")
                        
                        # VALIDACIÓN CRÍTICA: Verificar que es la vista correcta de Productos
                        if total < 10000:
                            print(f"[!] ADVERTENCIA: Solo se detectaron {total} registros")
                            print(f"[!] Se esperaban ~14,760 registros de Productos")
                            print(f"[!] Probablemente NO está en la vista correcta")
                            
                            # Intentar navegar directamente a Productos por URL
                            print("[*] Intentando navegar directamente a Productos...")
                            driver.get(URL_PRODUCTOS)
                            time.sleep(5)
                            
                            # Reintentar detectar total
                            print("[*] Reintentando detección de registros...")
                            retry_total = None
                            retry_elem = None
                            
                            for retry_elemento in driver.find_elements(By.XPATH, "//span[contains(@class, 'o_pager')]"):
                                try:
                                    if retry_elemento.is_displayed():
                                        retry_texto = retry_elemento.text.strip()
                                        print(f"[*] Texto paginador: '{retry_texto}'")
                                        retry_match = re.search(r'(\d+)\s*/\s*(\d+)', retry_texto)
                                        if retry_match:
                                            retry_total = int(retry_match.group(2))
                                            if retry_total >= 10000:
                                                print(f"[OK] Ahora sí: {retry_total} registros detectados")
                                                return retry_total, retry_elemento
                                            else:
                                                print(f"[!] Aún solo hay {retry_total} registros")
                                except:
                                    continue
                            
                            # Si llegamos aquí, el reintento falló
                            error_msg = f"Vista incorrecta: Solo {total} registros en lugar de ~14,760. URL: {driver.current_url}"
                            raise Exception(error_msg)
                        
                        return total, elem
                        
            except:
                continue
        
        raise Exception("No se pudo detectar el total de registros")
        
    except Exception as e:
        print(f"[ERROR] Error detectando total: {str(e)}")
        return None, None

def modificar_rango_registros(driver, total):
    """Modifica el campo de rango para mostrar '1-TOTAL'"""
    print(f"[*] Modificando rango a 1-{total}...")
    wait = WebDriverWait(driver, 10)
    
    try:
        time.sleep(2)
        
        # Buscar el elemento que contiene "1-80" (o similar)
        print("[*] Buscando el elemento con el rango actual (ej: 1-80)...")
        
        # Buscar elementos que contengan texto tipo "1-80" o "1-100"
        selectores = [
            "//span[contains(@class, 'o_pager_value')]",
            "//span[contains(@class, 'o_pager_limit')]", 
            "//*[contains(text(), '1-')]",
        ]
        
        campo_rango = None
        
        for selector in selectores:
            try:
                print(f"[*] Buscando con: {selector[:50]}...")
                elementos = driver.find_elements(By.XPATH, selector)
                
                for elem in elementos:
                    if elem.is_displayed():
                        # Obtener el texto
                        texto = elem.text.strip()
                        
                        print(f"[*] Elemento encontrado: tag={elem.tag_name}, texto='{texto}'")
                        
                        # Verificar si contiene el patrón "1-XX"
                        if re.match(r'^\d+-\d+$', texto):
                            print(f"[OK] Encontrado campo con rango: '{texto}'")
                            campo_rango = elem
                            break
                
                if campo_rango:
                    break
                    
            except Exception as e:
                continue
        
        if not campo_rango:
            raise Exception("No se encontró el campo de rango")
        
        # Hacer scroll al elemento
        driver.execute_script("arguments[0].scrollIntoView(true);", campo_rango)
        time.sleep(0.5)
        
        # ESTRATEGIA: Triple click para seleccionar todo el texto y luego escribir
        print("[*] Haciendo triple click para seleccionar todo el texto...")
        
        # Importar ActionChains
        from selenium.webdriver.common.action_chains import ActionChains
        
        actions = ActionChains(driver)
        
        # Triple click en el elemento para seleccionar todo
        actions.move_to_element(campo_rango).click().click().click().perform()
        time.sleep(0.5)
        print("[OK] Triple click realizado, texto seleccionado")
        
        # Ahora escribir el nuevo rango (sobrescribirá el texto seleccionado)
        nuevo_rango = f"1-{total}"
        print(f"[*] Escribiendo: {nuevo_rango}")
        
        # Usar send_keys directamente en el elemento activo
        actions.send_keys(nuevo_rango).perform()
        time.sleep(1)
        print("[OK] Texto escrito")
        
        # Presionar Enter
        print("[*] Presionando Enter...")
        actions.send_keys(Keys.RETURN).perform()
        time.sleep(3)
        
        print(f"[OK] Rango modificado a {nuevo_rango}")
        print("[*] Esperando a que inicie la carga...")
        time.sleep(2)
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Error modificando rango: {str(e)}")
        
        # Screenshot para debugging
        try:
            screenshot_path = os.path.join(os.path.dirname(__file__), "error_modificar_rango.png")
            driver.save_screenshot(screenshot_path)
            print(f"[*] Screenshot guardado en: {screenshot_path}")
        except:
            pass
        
        return False

def esperar_carga_completa(driver, total, timeout=3600):
    """
    Espera a que termine la carga de todos los registros
    Puede tardar 30+ minutos  
    Importante: Después de que desaparece el texto "Cargando", 
    la página puede seguir procesando datos en segundo plano
    """
    print(f"[*] Esperando que carguen los {total} registros...")
    print(f"[*] Tiempo maximo de espera: {timeout//60} minutos")
    print("[*] Esto puede tardar 30 minutos o mas...")
    print("[*] NOTA: La pagina puede seguir procesando aunque no haya texto visible")
    
    tiempo_inicio = time.time()
    ultimo_reporte = tiempo_inicio
    
    try:
        # Dar tiempo inicial para que empiece a cargar
        time.sleep(5)
        
        print("[*] Monitoreando la carga...")
        
        while time.time() - tiempo_inicio < timeout:
            # Reportar progreso cada minuto
            tiempo_actual = time.time()
            if tiempo_actual - ultimo_reporte >= 60:
                minutos_transcurridos = int((tiempo_actual - tiempo_inicio) / 60)
                print(f"[*] Tiempo transcurrido: {minutos_transcurridos} minuto(s)...")
                ultimo_reporte = tiempo_actual
            
            # ESTRATEGIA 0: Primero verificar si hay actividad de carga
            hay_actividad_carga = False
            
            # Verificar mensajes de carga
            try:
                textos_cargando = driver.find_elements(By.XPATH, 
                    "//*[contains(text(), 'cargando') or contains(text(), 'Cargando') or "
                    "contains(text(), 'loading') or contains(text(), 'Loading') or "
                    "contains(text(), 'paciente') or contains(text(), 'café')]"
                )
                
                for t in textos_cargando:
                    try:
                        if t.is_displayed():
                            texto = t.text.strip()
                            if texto and len(texto) > 5:
                                print(f"[*] Mensaje de carga: '{texto}'")
                                hay_actividad_carga = True
                                break
                    except:
                        continue
            except:
                pass
            
            # Si hay mensajes, esperar y continuar
            if hay_actividad_carga:
                time.sleep(10)
                continue
            
            # Verificar spinners
            try:
                spinners = driver.find_elements(By.XPATH, 
                    "//span[contains(@class, 'fa-spinner')] | "
                    "//div[contains(@class, 'o_loading')]"
                )
                
                spinners_visibles = [s for s in spinners if s.is_displayed()]
                
                if spinners_visibles:
                    print(f"[*] {len(spinners_visibles)} indicador(es) de carga activo(s)...")
                    hay_actividad_carga = True
            except:
                pass
            
            # Si hay spinners, esperar y continuar
            if hay_actividad_carga:
                time.sleep(5)
                continue
            
            # NO hay mensajes NI spinners - verificar overlay gris
            print("[*] No hay mensajes ni spinners, verificando overlay gris...")
            
            try:
                overlay_bloqueando = driver.execute_script("""
                    var overlays = document.querySelectorAll(
                        '[class*="blockUI"], [class*="o_loading"], ' +
                        '[class*="modal-backdrop"], .o_blockUI, .blockUI'
                    );
                    
                    for (var i = 0; i < overlays.length; i++) {
                        var style = window.getComputedStyle(overlays[i]);
                        if (style.display !== 'none' && style.visibility !== 'hidden') {
                            var rect = overlays[i].getBoundingClientRect();
                            if (rect.width > 500 && rect.height > 500) {
                                return true;
                            }
                        }
                    }
                    return false;
                """)
                
                if overlay_bloqueando:
                    print("[*] Overlay gris detectado (pantalla bloqueada sin mensajes)")
                    print("[*] Esperando a que desaparezca el overlay...")
                    time.sleep(10)
                    continue
                else:
                    print("[*] No hay overlay gris - verificando paginador...")
            except:
                pass
            
            # SOLO AHORA verificar el paginador (sin mensajes, sin spinners, sin overlay)
            try:
                elementos_paginador = driver.find_elements(By.XPATH, 
                    "//*[contains(@class, 'o_pager')] | //span[contains(@class, 'o_pager_value')]"
                )
                
                paginador_encontrado = False
                for elem in elementos_paginador:
                    try:
                        # Intentar obtener el texto incluso si no es visible (por overlay)
                        texto = elem.text.strip()
                        if not texto:
                            # Si no tiene texto visible, intentar con get_attribute
                            texto = elem.get_attribute('textContent').strip()
                        
                        if texto:
                            # Buscar patrón "1-14760 / 14760"
                            match = re.search(r'(\d+)\s*-\s*(\d+)\s*/\s*(\d+)', texto)
                            if match:
                                paginador_encontrado = True
                                final = int(match.group(2))
                                total_mostrado = int(match.group(3))
                                
                                # Si el rango final alcanza el 98% o más del total
                                if final >= total_mostrado * 0.98:
                                    print(f"[*] Paginador muestra: {texto} - Rango completo!")
                                    print("[OK] Sin mensajes, sin spinners, sin overlay - Carga completada!")
                                    tiempo_total = int((time.time() - tiempo_inicio) / 60)
                                    print(f"[OK] Tiempo total de carga: {tiempo_total} minuto(s)")
                                    return True
                                else:
                                    # Aún no alcanza el 98%
                                    porcentaje = (final / total_mostrado) * 100
                                    print(f"[*] Progreso: {final}/{total_mostrado} ({porcentaje:.1f}%)")
                    except:
                        continue
                        
            except Exception as e:
                pass
            
        # Timeout
        print(f"[!] Timeout ({timeout//60} min)")
        return False
        
    except Exception as e:
        print(f"[ERROR] Error esperando carga: {str(e)}")
        return False

def seleccionar_todos_registros(driver):
    """Marca el checkbox del header para seleccionar todos los registros"""
    print("[*] Seleccionando todos los registros...")
    
    try:
        print("[*] PASO 1: Buscando checkbox del header...")
        
        # Buscar el checkbox del header con múltiples estrategias
        checkbox_header = None
        
        # Estrategia 1: Selectores estándar
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
                        print(f"[OK] Checkbox encontrado con selector estándar")
                        break
                if checkbox_header:
                    break
            except:
                continue
        
        # Estrategia 2: JavaScript
        if not checkbox_header:
            print("[!] No se encontró con selectores estándar")
            print("[*] Buscando con JavaScript...")
            
            checkbox_js = driver.execute_script("""
                var checkbox = document.querySelector('thead th input[type="checkbox"]');
                if (checkbox) return checkbox;
                
                checkbox = document.querySelector('th.o_list_record_selector input[type="checkbox"]');
                if (checkbox) return checkbox;
                
                return null;
            """)
            
            if checkbox_js:
                checkbox_header = checkbox_js
                print("[OK] Checkbox encontrado con JavaScript")
        
        if not checkbox_header:
            raise Exception("No se pudo encontrar el checkbox del header")
        
        # PASO 2: Hacer click en el checkbox
        print("[*] PASO 2: Haciendo click en el checkbox...")
        print("[*] NOTA: Puede dar timeout - eso es NORMAL con 14,760 registros")
        
        # Intentar click normal primero
        try:
            driver.execute_script("arguments[0].scrollIntoView(true);", checkbox_header)
            time.sleep(0.5)
            checkbox_header.click()
            print("[OK] Click normal realizado")
        except Exception as e:
            print(f"[!] Error en click normal: {str(e)}")
            print("[*] Intentando con JavaScript...")
            
            # El JavaScript puede dar timeout con 14,760 registros - eso es NORMAL
            try:
                driver.execute_script("arguments[0].click();", checkbox_header)
                print("[OK] Click realizado (JavaScript)")
            except Exception as js_error:
                if "timeout" in str(js_error).lower() or "timed out" in str(js_error).lower():
                    print("[!] JavaScript dio timeout (esto es NORMAL con 14,760 registros)")
                    print("[*] El navegador está procesando la selección en segundo plano")
                    print("[*] Continuando con la espera de 5 minutos...")
                else:
                    # Si es otro error, re-lanzarlo
                    raise
        
        print("[OK] PASO 2 COMPLETADO: Click ejecutado en el checkbox")
        
        # PASO 3: Esperar a que Odoo procese la selección
        print("\n" + "=" * 70)
        print("PASO 3: ESPERANDO PROCESAMIENTO DE SELECCIÓN")
        print("=" * 70)
        print("[*] Odoo está procesando la selección de 14,760 registros")
        print("[*] El navegador puede estar bloqueado - esto es normal")
        print("[*] Esperando 5 minutos para que termine de procesar...")
        print("-" * 70)
        
        # Esperar 5 minutos dando feedback cada minuto
        for minuto in range(5):
            time.sleep(60)
            print(f"[*] {minuto + 1} minuto(s) de 5 transcurrido(s)...")
        
        print("\n" + "=" * 70)
        print("[OK] PROCESAMIENTO COMPLETADO")
        print("=" * 70)
        print(f"[*] Esperados 5 minutos")
        
        # PASO 4: VERIFICACIÓN OBLIGATORIA - Debe aparecer el botón "Acción"
        print("[*] PASO 4: Verificando botón 'Acción'...")
        
        try:
            # Buscar el botón "Acción" con múltiples estrategias
            selectores_accion = [
                "//button[contains(., 'Acción')]",
                "//button[contains(., 'Accion')]",
                "//button[contains(text(), 'Acción')]",
                "//button[contains(text(), 'Accion')]",
            ]
            
            boton_encontrado = False
            for selector in selectores_accion:
                try:
                    botones = driver.find_elements(By.XPATH, selector)
                    for boton in botones:
                        if boton.is_displayed():
                            texto = boton.text.strip()
                            print(f"[OK] ¡Botón 'Acción' está visible: '{texto}'")
                            boton_encontrado = True
                            break
                    if boton_encontrado:
                        break
                except:
                    continue
            
            if not boton_encontrado:
                raise Exception("El botón 'Acción' NO apareció - los registros NO están seleccionados")
            
            print("[*] Listo para continuar con la exportación")
            print("-" * 70)
            return True
            
        except Exception as e:
            print(f"[ERROR] {str(e)}")
            
            # Tomar screenshot
            try:
                screenshot_path = os.path.join(os.path.dirname(__file__), "error_sin_boton_accion.png")
                driver.save_screenshot(screenshot_path)
                print(f"[*] Screenshot guardado en: {screenshot_path}")
            except:
                pass
            
            raise
        
    except Exception as e:
        print(f"[ERROR] Error inesperado: {str(e)}")
        
        # Screenshot
        try:
            screenshot_path = os.path.join(os.path.dirname(__file__), "error_seleccion.png")
            driver.save_screenshot(screenshot_path)
            print(f"[*] Screenshot: {screenshot_path}")
        except:
            pass
        
        return False
def abrir_menu_accion(driver):
    """Abre el menú 'Acción'"""
    print("[*] Abriendo menu 'Accion'...")
    wait = WebDriverWait(driver, 10)
    
    try:
        time.sleep(2)
        
        # Buscar el botón "Acción" (con y sin tilde)
        selectores_accion = [
            # Con tilde
            "//button[contains(., 'Acción')]",
            "//button[contains(text(), 'Acción')]",
            "//a[contains(text(), 'Acción')]",
            # Sin tilde
            "//button[contains(., 'Accion')]",
            "//button[contains(text(), 'Accion')]",
            "//a[contains(text(), 'Accion')]",
            # Mayúsculas
            "//button[contains(translate(., 'abcdefghijklmnopqrstuvwxyzáéíóúñ', 'ABCDEFGHIJKLMNOPQRSTUVWXYZAEIOUN'), 'ACCION')]",
            # Por clase o atributo
            "//button[contains(@class, 'dropdown') and contains(., 'Acc')]",
            "//*[@data-toggle='dropdown' and contains(., 'Acc')]",
            # Cualquier elemento con "Action"
            "//*[contains(text(), 'Action')]",
        ]
        
        for selector in selectores_accion:
            try:
                print(f"[*] Buscando boton Accion con: {selector[:70]}...")
                elementos = driver.find_elements(By.XPATH, selector)
                
                for boton in elementos:
                    if boton.is_displayed():
                        texto = boton.text.strip()
                        print(f"[*] Elemento encontrado con texto: '{texto}'")
                        
                        # Verificar que contiene "Acción" o "Accion"
                        if 'acci' in texto.lower() or 'action' in texto.lower():
                            print("[*] Haciendo click...")
                            
                            # Scroll
                            driver.execute_script("arguments[0].scrollIntoView(true);", boton)
                            time.sleep(0.5)
                            
                            # Click
                            try:
                                boton.click()
                            except:
                                driver.execute_script("arguments[0].click();", boton)
                            
                            time.sleep(2)
                            print("[OK] Menu 'Accion' abierto")
                            return True
            except Exception as e:
                continue
        
        # Si no encontró, buscar TODOS los botones y mostrarlos
        print("[*] No se encontró con selectores específicos")
        print("[*] Buscando todos los botones visibles...")
        
        try:
            todos_botones = driver.find_elements(By.XPATH, "//button | //a[@role='button']")
            botones_visibles = []
            
            for btn in todos_botones:
                try:
                    if btn.is_displayed():
                        texto = btn.text.strip()
                        if texto and len(texto) < 30:
                            botones_visibles.append(texto)
                except:
                    continue
            
            if botones_visibles:
                print(f"[*] Botones visibles encontrados ({len(botones_visibles)}):")
                for i, texto in enumerate(botones_visibles[:10]):  # Mostrar solo los primeros 10
                    print(f"    {i+1}. '{texto}'")
        except:
            pass
        
        raise Exception("No se encontró el botón 'Acción'")
        
    except Exception as e:
        print(f"[ERROR] Error abriendo menu Accion: {str(e)}")
        
        # Screenshot
        try:
            screenshot_path = os.path.join(os.path.dirname(__file__), "error_accion.png")
            driver.save_screenshot(screenshot_path)
            print(f"[*] Screenshot guardado en: {screenshot_path}")
        except:
            pass
        
        return False

def seleccionar_exportar(driver):
    """Selecciona 'Exportar' del menú Acción"""
    print("[*] Seleccionando 'Exportar'...")
    wait = WebDriverWait(driver, 10)
    
    try:
        selectores_exportar = [
            "//a[contains(text(), 'Exportar')]",
            "//span[contains(text(), 'Exportar')]",
            "//*[contains(text(), 'Exportar') and not(contains(text(), 'fichero'))]",
        ]
        
        for selector in selectores_exportar:
            try:
                print(f"[*] Buscando opcion Exportar con: {selector[:60]}...")
                opcion = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                opcion.click()
                time.sleep(3)
                print("[OK] 'Exportar' seleccionado")
                return True
            except:
                continue
        
        raise Exception("No se encontró la opción 'Exportar'")
        
    except Exception as e:
        print(f"[ERROR] Error seleccionando Exportar: {str(e)}")
        return False

def seleccionar_inventario_general(driver):
    """Selecciona 'INVENTARIO GENERAL' del dropdown 'Exportaciones guardadas'"""
    print("[*] Seleccionando 'INVENTARIO GENERAL'...")
    wait = WebDriverWait(driver, 10)
    
    try:
        # Esperar a que aparezca el modal
        time.sleep(3)
        
        # PASO 1: Buscar y hacer click en el dropdown para desplegarlo
        print("[*] Buscando dropdown 'Exportaciones guardadas'...")
        selectores_dropdown = [
            "//select[contains(@name, 'export')]",
            "//select",
            "//div[contains(@class, 'modal')]//select",
            "//div[contains(text(), 'Exportaciones guardadas')]//following-sibling::select",
            "//label[contains(text(), 'Exportaciones guardadas')]//following-sibling::select",
        ]
        
        dropdown = None
        for selector in selectores_dropdown:
            try:
                print(f"[*] Intentando: {selector[:70]}...")
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
            # Intentar con JavaScript
            print("[*] Buscando con JavaScript...")
            dropdown_js = driver.execute_script("""
                var selects = document.querySelectorAll('select');
                for (var i = 0; i < selects.length; i++) {
                    if (selects[i].offsetParent !== null) {
                        return selects[i];
                    }
                }
                return null;
            """)
            
            if dropdown_js:
                dropdown = dropdown_js
                print("[OK] Dropdown encontrado con JavaScript")
        
        if not dropdown:
            raise Exception("No se encontró el dropdown")
        
        # PASO 2: Hacer scroll y click en el dropdown
        print("[*] Haciendo scroll al dropdown...")
        driver.execute_script("arguments[0].scrollIntoView(true);", dropdown)
        time.sleep(0.5)
        
        print("[*] Haciendo click en dropdown...")
        try:
            dropdown.click()
        except:
            driver.execute_script("arguments[0].click();", dropdown)
        
        time.sleep(1)
        
        # PASO 3: Seleccionar "INVENTARIO GENERAL"
        print("[*] Buscando opcion 'INVENTARIO GENERAL'...")
        from selenium.webdriver.support.ui import Select
        select = Select(dropdown)
        
        # Listar opciones
        opciones = select.options
        print(f"[*] Opciones disponibles ({len(opciones)}):")
        for i, opc in enumerate(opciones):
            texto = opc.text.strip()
            print(f"    {i+1}. '{texto}'")
            
            # Buscar INVENTARIO GENERAL
            texto_upper = texto.upper()
            if "INVENTARIO" in texto_upper and "GENERAL" in texto_upper:
                print(f"[*] Seleccionando: '{texto}'")
                try:
                    select.select_by_visible_text(texto)
                except:
                    select.select_by_index(i)
                time.sleep(1)
                print("[OK] 'INVENTARIO GENERAL' seleccionado")
                return True
        
        # Si no encontró por texto, buscar por valor
        print("[*] Buscando por valor de opción...")
        for opc in opciones:
            valor = opc.get_attribute('value')
            texto = opc.text.strip()
            if valor and ("inventario" in valor.lower() or "general" in valor.lower()):
                print(f"[*] Seleccionando por valor: '{texto}'")
                select.select_by_value(valor)
                time.sleep(1)
                print("[OK] Seleccionado")
                return True
        
        raise Exception("No se encontró 'INVENTARIO GENERAL' en las opciones")
        
    except Exception as e:
        print(f"[ERROR] Error seleccionando INVENTARIO GENERAL: {str(e)}")
        
        # Screenshot
        try:
            screenshot_path = os.path.join(os.path.dirname(__file__), "error_inventario_general.png")
            driver.save_screenshot(screenshot_path)
            print(f"[*] Screenshot guardado en: {screenshot_path}")
        except:
            pass
        
        return False

def exportar_fichero(driver):
    """Hace click en 'Exportar a fichero'"""
    print("[*] Haciendo click en 'Exportar a fichero'...")
    wait = WebDriverWait(driver, 10)
    
    try:
        selectores_boton = [
            "//button[contains(text(), 'Exportar a fichero')]",
            "//button[contains(., 'Exportar') and contains(., 'fichero')]",
            "//button[contains(@class, 'btn-primary') and contains(., 'Exportar')]",
        ]
        
        for selector in selectores_boton:
            try:
                print(f"[*] Buscando boton con: {selector[:60]}...")
                boton = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                boton.click()
                time.sleep(3)
                print("[OK] Exportacion iniciada")
                return True
            except:
                continue
        
        raise Exception("No se encontró el botón 'Exportar a fichero'")
        
    except Exception as e:
        print(f"[ERROR] Error exportando fichero: {str(e)}")
        return False

def esperar_descarga_archivo(carpeta, timeout=300):
    """
    Espera a que se complete la descarga del archivo
    """
    print(f"[*] Esperando descarga del archivo (maximo {timeout//60} minutos)...")
    
    tiempo_inicio = time.time()
    
    while time.time() - tiempo_inicio < timeout:
        # Buscar archivos temporales de descarga
        archivos_temp = glob.glob(os.path.join(carpeta, "*.crdownload"))
        archivos_temp += glob.glob(os.path.join(carpeta, "*.tmp"))
        
        if not archivos_temp:
            # Buscar archivos descargados recientemente
            archivos_xlsx = glob.glob(os.path.join(carpeta, "*.xlsx"))
            archivos_xls = glob.glob(os.path.join(carpeta, "*.xls"))
            archivos_csv = glob.glob(os.path.join(carpeta, "*.csv"))
            
            todos_archivos = archivos_xlsx + archivos_xls + archivos_csv
            
            if todos_archivos:
                # Filtrar archivos temporales de Excel (~$)
                todos_archivos = [f for f in todos_archivos if not os.path.basename(f).startswith("~$")]
                
                if todos_archivos:
                    # Ordenar por fecha de modificacion (mas reciente primero)
                    archivo_mas_reciente = max(todos_archivos, key=os.path.getmtime)
                    
                    # Verificar que fue creado recientemente (ultimos 5 minutos)
                    tiempo_modificacion = os.path.getmtime(archivo_mas_reciente)
                    if time.time() - tiempo_modificacion < 300:
                        print(f"[OK] Archivo descargado: {os.path.basename(archivo_mas_reciente)}")
                        return archivo_mas_reciente
        
        time.sleep(2)
    
    print("[!] Timeout esperando la descarga")
    return None

def renombrar_archivo_con_fecha(archivo_original):
    """
    Renombra el archivo con formato: INVENTARIO GENERAL ACTUALIZADO DD DE MES DE YYYY
    """
    print("[*] Renombrando archivo con fecha...")
    
    try:
        # Obtener información del archivo
        directorio = os.path.dirname(archivo_original)
        extension = os.path.splitext(archivo_original)[1]  # .xls o .xlsx
        
        # Obtener fecha actual
        fecha_actual = datetime.now()
        dia = fecha_actual.day
        mes = fecha_actual.month
        anio = fecha_actual.year
        
        # Meses en español
        meses = {
            1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
            5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
            9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
        }
        
        nombre_mes = meses[mes]
        
        # Crear nuevo nombre: "INVENTARIO GENERAL ACTUALIZADO 02 DE DICIEMBRE DE 2025"
        nuevo_nombre = f"INVENTARIO GENERAL ACTUALIZADO {dia:02d} DE {nombre_mes} DE {anio}{extension}"
        nueva_ruta = os.path.join(directorio, nuevo_nombre)
        
        # Renombrar el archivo
        if os.path.exists(nueva_ruta):
            # Si ya existe, eliminarlo primero
            print(f"[*] Eliminando archivo antiguo: {nuevo_nombre}")
            os.remove(nueva_ruta)
        
        os.rename(archivo_original, nueva_ruta)
        print(f"[OK] Archivo renombrado a: {nuevo_nombre}")
        
        return nueva_ruta
        
    except Exception as e:
        print(f"[ERROR] Error renombrando archivo: {str(e)}")
        return archivo_original

def enviar_email_notificacion(exito=True, archivo_descargado=None, tiempo_total=None, error=None):
    """
    Envía un email de notificación al finalizar el proceso
    """
    if not EMAIL_CONFIG.get("enabled", False):
        print("[*] Envio de email desactivado")
        return False
    
    print("[*] Enviando notificacion por email...")
    
    try:
        # Configurar mensaje
        msg = MIMEMultipart('alternative')
        msg['From'] = EMAIL_CONFIG['sender_email']
        msg['To'] = ', '.join(EMAIL_CONFIG['recipient_emails'])
        
        fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        if exito:
            msg['Subject'] = f"✅ FERTRAC: Descarga de Inventario Completada - {datetime.now().strftime('%d/%m/%Y')}"
            
            # Cuerpo HTML
            html = f"""
            <html>
            <head>
                <style>
                    body {{ font-family: Arial, sans-serif; }}
                    .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
                    .header {{ background-color: #28a745; color: white; padding: 20px; border-radius: 5px; }}
                    .content {{ background-color: #f8f9fa; padding: 20px; border-radius: 5px; margin-top: 20px; }}
                    .info {{ margin: 10px 0; }}
                    .label {{ font-weight: bold; color: #333; }}
                    .value {{ color: #666; }}
                    .footer {{ margin-top: 20px; font-size: 12px; color: #999; }}
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="header">
                        <h2>✅ Descarga de Inventario Completada</h2>
                    </div>
                    <div class="content">
                        <div class="info">
                            <span class="label">📅 Fecha y hora:</span>
                            <span class="value">{fecha_actual}</span>
                        </div>
                        <div class="info">
                            <span class="label">📁 Archivo descargado:</span>
                            <span class="value">{os.path.basename(archivo_descargado) if archivo_descargado else 'N/A'}</span>
                        </div>
                        <div class="info">
                            <span class="label">📂 Ubicación:</span>
                            <span class="value">{os.path.dirname(archivo_descargado) if archivo_descargado else 'N/A'}</span>
                        </div>
                        <div class="info">
                            <span class="label">⏱️ Tiempo total:</span>
                            <span class="value">{tiempo_total if tiempo_total else 'N/A'}</span>
                        </div>
                        <div class="info">
                            <span class="label">🎯 Proceso:</span>
                            <span class="value">Descarga de Inventario General (Productos)</span>
                        </div>
                    </div>
                    <div class="footer">
                        <p>Este es un mensaje automático generado por el sistema de automatización FERTRAC.</p>
                        <p>Script: descargar_productos_fertrac_PARTE1.py</p>
                    </div>
                </div>
            </body>
            </html>
            """
        else:
            msg['Subject'] = f"❌ FERTRAC: Error en Descarga de Inventario - {datetime.now().strftime('%d/%m/%Y')}"
            
            # Cuerpo HTML para error
            html = f"""
            <html>
            <head>
                <style>
                    body {{ font-family: Arial, sans-serif; }}
                    .container {{ max-width: 600px; margin: 0 auto; padding: 20px; }}
                    .header {{ background-color: #dc3545; color: white; padding: 20px; border-radius: 5px; }}
                    .content {{ background-color: #f8f9fa; padding: 20px; border-radius: 5px; margin-top: 20px; }}
                    .info {{ margin: 10px 0; }}
                    .label {{ font-weight: bold; color: #333; }}
                    .value {{ color: #666; }}
                    .error {{ background-color: #fff3cd; padding: 10px; border-left: 4px solid #ffc107; margin: 15px 0; }}
                    .footer {{ margin-top: 20px; font-size: 12px; color: #999; }}
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="header">
                        <h2>❌ Error en Descarga de Inventario</h2>
                    </div>
                    <div class="content">
                        <div class="info">
                            <span class="label">📅 Fecha y hora:</span>
                            <span class="value">{fecha_actual}</span>
                        </div>
                        <div class="error">
                            <strong>Error:</strong><br>
                            {error if error else 'Error desconocido'}
                        </div>
                        <div class="info">
                            <span class="label">🎯 Proceso:</span>
                            <span class="value">Descarga de Inventario General (Productos)</span>
                        </div>
                        <div class="info">
                            <span class="label">⚠️ Acción requerida:</span>
                            <span class="value">Revisar logs y screenshots en la carpeta del script</span>
                        </div>
                    </div>
                    <div class="footer">
                        <p>Este es un mensaje automático generado por el sistema de automatización FERTRAC.</p>
                        <p>Script: descargar_productos_fertrac_PARTE1.py</p>
                    </div>
                </div>
            </body>
            </html>
            """
        
        # Adjuntar HTML
        part = MIMEText(html, 'html')
        msg.attach(part)
        
        # Enviar email
        with smtplib.SMTP(EMAIL_CONFIG['smtp_server'], EMAIL_CONFIG['smtp_port']) as server:
            server.starttls()
            server.login(EMAIL_CONFIG['sender_email'], EMAIL_CONFIG['sender_password'])
            server.send_message(msg)
        
        print("[OK] Email enviado exitosamente")
        return True
        
    except Exception as e:
        print(f"[ERROR] Error enviando email: {str(e)}")
        return False
    
def main():
    """Funcion principal que ejecuta todo el proceso"""
    driver = None
    tiempo_inicio_total = time.time()
    archivo_final = None
    
    try:
        print("=" * 70)
        print("AUTOMATIZACION FERTRAC - PRODUCTOS")
        print("=" * 70)
        print(f"Fecha y hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        print(f"Usuario: {USUARIO}")
        print("-" * 70)
        
        # Crear carpeta del mes
        carpeta_descarga = crear_carpeta_mes()
        print(f"Carpeta de descarga: {carpeta_descarga}")
        print("-" * 70)
        
        # Inicializar driver con carpeta de descarga
        driver = configurar_driver(carpeta_descarga)
        
        # Login
        if not hacer_login(driver):
            raise Exception("Fallo en el login")
        
        # Navegar a inventario
        navegar_a_inventario(driver)
        
        # Seleccionar Productos
        if not seleccionar_productos(driver):
            raise Exception("Fallo al seleccionar Productos")
        
        # Cambiar a vista lista
        if not cambiar_a_vista_lista(driver):
            raise Exception("Fallo al cambiar a vista lista")
        
        # Detectar total de registros
        total, elemento_paginador = detectar_total_registros(driver)
        if total is None:
            raise Exception("Fallo al detectar total de registros")
        
        # Modificar rango
        if not modificar_rango_registros(driver, total):
            raise Exception("Fallo al modificar rango de registros")
        
        # Esperar carga completa
        if not esperar_carga_completa(driver, total, TIMEOUT_CARGA_MAXIMA):
            raise Exception("Fallo al esperar carga completa")
        
        # NUEVOS PASOS: Exportar
        print("\n" + "=" * 70)
        print("INICIANDO EXPORTACION")
        print("=" * 70)
        
        # Seleccionar todos los registros
        if not seleccionar_todos_registros(driver):
            raise Exception("Fallo al seleccionar todos los registros")
        
        # Abrir menu Accion
        if not abrir_menu_accion(driver):
            raise Exception("Fallo al abrir menu Accion")
        
        # Seleccionar Exportar
        if not seleccionar_exportar(driver):
            raise Exception("Fallo al seleccionar Exportar")
        
        # Seleccionar INVENTARIO GENERAL
        if not seleccionar_inventario_general(driver):
            raise Exception("Fallo al seleccionar INVENTARIO GENERAL")
        
        # Exportar a fichero
        if not exportar_fichero(driver):
            raise Exception("Fallo al exportar fichero")
        
        # IMPORTANTE: Esperar pantalla de carga después de exportar
        print("\n" + "=" * 70)
        print("ESPERANDO PROCESAMIENTO DE EXPORTACION")
        print("=" * 70)
        print("[*] El sistema puede tardar varios minutos procesando los datos...")
        
        # Esperar a que desaparezcan los indicadores de carga
        tiempo_inicio = time.time()
        timeout_exportacion = 600  # 10 minutos máximo
        
        while time.time() - tiempo_inicio < timeout_exportacion:
            # Buscar indicadores de carga
            try:
                spinners = driver.find_elements(By.XPATH, 
                    "//span[contains(@class, 'fa-spinner')] | "
                    "//div[contains(@class, 'o_loading')]"
                )
                
                spinners_visibles = [s for s in spinners if s.is_displayed()]
                
                if spinners_visibles:
                    tiempo_transcurrido = int((time.time() - tiempo_inicio) / 60)
                    if tiempo_transcurrido > 0 and tiempo_transcurrido % 1 == 0:
                        print(f"[*] Procesando exportacion... {tiempo_transcurrido} minuto(s)")
                    time.sleep(5)
                    continue
                else:
                    print("[OK] Procesamiento completado")
                    break
            except:
                time.sleep(2)
                continue
        
        # Esperar a que se descargue el archivo
        print("\n" + "=" * 70)
        print("ESPERANDO DESCARGA DEL ARCHIVO")
        print("=" * 70)
        
        archivo_descargado = esperar_descarga_archivo(carpeta_descarga, timeout=300)
        
        if archivo_descargado:
            # Renombrar el archivo con la fecha
            archivo_final = renombrar_archivo_con_fecha(archivo_descargado)
            
            # Calcular tiempo total
            tiempo_total_segundos = int(time.time() - tiempo_inicio_total)
            tiempo_total_minutos = tiempo_total_segundos // 60
            tiempo_total_texto = f"{tiempo_total_minutos} minutos"
            
            print("\n" + "=" * 70)
            print("PROCESO COMPLETADO EXITOSAMENTE")
            print("=" * 70)
            print(f"Archivo: {os.path.basename(archivo_final)}")
            print(f"Ubicacion: {carpeta_descarga}")
            print(f"Tiempo total: {tiempo_total_texto}")
            print("=" * 70)
            
            # Enviar email de notificación
            enviar_email_notificacion(
                exito=True,
                archivo_descargado=archivo_final,
                tiempo_total=tiempo_total_texto
            )
            
            # Esperar 5 segundos antes de cerrar
            print("\n[*] Cerrando navegador en 5 segundos...")
            time.sleep(5)
        else:
            raise Exception("No se pudo verificar la descarga del archivo")
        
    except Exception as e:
        print(f"\n[ERROR] ERROR: {str(e)}")
        
        if driver:
            try:
                screenshot_path = os.path.join(os.path.dirname(__file__), "error_general.png")
                driver.save_screenshot(screenshot_path)
                print(f"[*] Screenshot guardado en: {screenshot_path}")
            except:
                pass
        
        # Enviar email de error
        enviar_email_notificacion(
            exito=False,
            error=str(e)
        )
        
        # Esperar 5 segundos antes de cerrar en caso de error
        print("\n[*] Cerrando navegador en 5 segundos...")
        time.sleep(5)
        
    finally:
        if driver:
            driver.quit()
            print("\n[*] Navegador cerrado")
        
        print(f"[*] Finalizado: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

# ============== EJECUCION ==============

if __name__ == "__main__":
    main()