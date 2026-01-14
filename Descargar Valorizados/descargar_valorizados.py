"""
Script para descargar informes Valorizados por Almacén desde ERP Fertrac
- Navega a Inventario > Informes > Valorizado
- Descarga 4 informes (Fertrac Principal, Toberin, Faltantes, Faltantes_Impo)
- Los renombra automáticamente con nombres específicos
- Los guarda en la carpeta de destino
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from datetime import datetime
import time
import os
import sys
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
URL_INVENTARIO = "https://erp.fertrac.com/web?#action=246&model=stock.picking.type&view_type=kanban&menu_id=174"

# Ruta base para descargas
RUTA_DESCARGA = r"D:\Fertrac\Usuarios\infocompras\ARCHIVOS DIARIOS 2026\Pruebas Inv General\Valorizados"

# Configuracion
MODO_HEADLESS = False
TIMEOUT_DESCARGA = 60  # 1 minuto máximo por descarga

# Diccionario de almacenes y nombres de archivo
ALMACENES_CONFIG = [
    {
        "tipo_consulta": None,  # No tocar, Compañía está predeterminado
        "ubicacion": None,
        "nombre_archivo": "VALORIZADO GENERAL.xlsx"
    },
    {
        "tipo_consulta": "Ubicación",
        "ubicacion": "3/Toberin",
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
    "sender_email": "analista_automatizacion@fertrac.com",
    "sender_password": "lbih abom caxy pzbh",
    "recipient_emails": [
        "analista_automatizacion@fertrac.com",
        "data_science@fertrac.com",
        # "asistentecompras@fertrac.com",
        # "analistacompras5@fertrac.com",
    ],
    "enabled": True
}

# ============== FUNCIONES AUXILIARES ==============

def crear_carpeta_destino():
    """Crea la carpeta de destino si no existe"""
    if not os.path.exists(RUTA_DESCARGA):
        os.makedirs(RUTA_DESCARGA, exist_ok=True)
        print(f"[+] Carpeta creada: {RUTA_DESCARGA}")
    else:
        print(f"[+] Carpeta ya existe: {RUTA_DESCARGA}")
    
    return RUTA_DESCARGA

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
    
    # Configurar timeouts
    driver.set_page_load_timeout(120)
    driver.set_script_timeout(120)
    driver.implicitly_wait(10)
    
    if not MODO_HEADLESS:
        driver.maximize_window()
    
    print("[OK] Driver configurado correctamente")
    return driver

# ============== FUNCIONES DE AUTOMATIZACION ==============

def hacer_login(driver):
    """Realiza el login en el sistema"""
    print("[*] Iniciando sesion...")
    driver.get(URL_LOGIN)
    
    wait = WebDriverWait(driver, 20)
    
    try:
        time.sleep(2)
        
        campo_usuario = wait.until(EC.presence_of_element_located((By.NAME, "login")))
        campo_usuario.clear()
        campo_usuario.send_keys(USUARIO)
        print("[OK] Usuario ingresado")
        
        campo_clave = driver.find_element(By.NAME, "password")
        campo_clave.clear()
        campo_clave.send_keys(CLAVE)
        print("[OK] Contrasena ingresada")
        
        boton_login = driver.find_element(By.XPATH, "//button[@type='submit']")
        boton_login.click()
        
        time.sleep(3)
        print("[OK] Sesion iniciada correctamente")
        
        # Esperar 5 segundos adicionales para asegurar carga completa
        print("[*] Esperando carga completa del sistema...")
        time.sleep(5)
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Error en login: {str(e)}")
        return False

def navegar_a_inventario(driver):
    """Navega a la sección de Inventario"""
    print("[*] Navegando a Inventario...")
    driver.get(URL_INVENTARIO)
    time.sleep(5)
    print("[OK] En la seccion de Inventario")
    
    # Esperar 5 segundos adicionales para asegurar carga completa
    print("[*] Esperando carga completa de Inventario...")
    time.sleep(5)

def abrir_menu_informes(driver):
    """Hace click en el menú 'Informes'"""
    print("[*] Abriendo menu 'Informes'...")
    wait = WebDriverWait(driver, 10)
    
    try:
        time.sleep(2)
        
        # Buscar el botón/enlace "Informes"
        selectores_informes = [
            "//a[contains(text(), 'Informes')]",
            "//span[contains(text(), 'Informes')]",
            "//div[contains(text(), 'Informes')]",
            "//button[contains(., 'Informes')]",
        ]
        
        for selector in selectores_informes:
            try:
                print(f"[*] Buscando 'Informes' con: {selector[:60]}...")
                elementos = driver.find_elements(By.XPATH, selector)
                
                for elem in elementos:
                    if elem.is_displayed():
                        texto = elem.text.strip()
                        if texto == "Informes":
                            print(f"[*] Haciendo click en 'Informes'...")
                            driver.execute_script("arguments[0].scrollIntoView(true);", elem)
                            time.sleep(0.5)
                            
                            try:
                                elem.click()
                            except:
                                driver.execute_script("arguments[0].click();", elem)
                            
                            time.sleep(2)
                            print("[OK] Menu 'Informes' abierto")
                            
                            # Esperar 5 segundos adicionales para carga del menú
                            print("[*] Esperando carga completa del menu...")
                            time.sleep(5)
                            
                            return True
            except:
                continue
        
        raise Exception("No se encontro el menu 'Informes'")
        
    except Exception as e:
        print(f"[ERROR] Error abriendo menu Informes: {str(e)}")
        
        try:
            screenshot_path = os.path.join(os.path.dirname(__file__), "error_informes.png")
            driver.save_screenshot(screenshot_path)
            print(f"[*] Screenshot guardado en: {screenshot_path}")
        except:
            pass
        
        return False

def seleccionar_valorizado(driver):
    """Selecciona la opción 'Valorizado' del menú Informes"""
    print("[*] Seleccionando 'Valorizado'...")
    wait = WebDriverWait(driver, 10)
    
    try:
        # Buscar la opción "Valorizado"
        selectores_valorizado = [
            "//a[contains(text(), 'Valorizado') and not(contains(text(), 'Valorización'))]",
            "//span[contains(text(), 'Valorizado') and not(contains(text(), 'Valorización'))]",
            "//*[contains(text(), 'Valorizado') and not(contains(text(), 'Valorización'))]",
        ]
        
        for selector in selectores_valorizado:
            try:
                print(f"[*] Buscando 'Valorizado' con: {selector[:60]}...")
                elementos = driver.find_elements(By.XPATH, selector)
                
                for elem in elementos:
                    if elem.is_displayed():
                        texto = elem.text.strip()
                        if texto == "Valorizado":
                            print(f"[*] Haciendo click en 'Valorizado'...")
                            try:
                                elem.click()
                            except:
                                driver.execute_script("arguments[0].click();", elem)
                            
                            time.sleep(3)
                            print("[OK] 'Valorizado' seleccionado - Modal abierto")
                            
                            # Esperar 5 segundos adicionales para carga del modal
                            print("[*] Esperando carga completa del modal...")
                            time.sleep(5)
                            
                            return True
            except:
                continue
        
        raise Exception("No se encontro la opcion 'Valorizado'")
        
    except Exception as e:
        print(f"[ERROR] Error seleccionando Valorizado: {str(e)}")
        
        try:
            screenshot_path = os.path.join(os.path.dirname(__file__), "error_valorizado.png")
            driver.save_screenshot(screenshot_path)
            print(f"[*] Screenshot guardado en: {screenshot_path}")
        except:
            pass
        
        return False

def seleccionar_tipo_consulta(driver, tipo_consulta):
    """Selecciona el tipo de consulta especificado (Compañía, Ubicación, Almacén, etc)"""
    print(f"[*] Seleccionando '{tipo_consulta}' en 'Tipo de consulta'...")
    wait = WebDriverWait(driver, 10)
    
    try:
        time.sleep(2)
        
        # Buscar el dropdown "Tipo de consulta"
        selectores_dropdown = [
            "//select",
            "//div[contains(@class, 'modal')]//select",
        ]
        
        dropdown_encontrado = False
        
        for selector in selectores_dropdown:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                
                for elem in elementos:
                    if elem.is_displayed():
                        # Verificar si este es el dropdown correcto buscando opciones
                        try:
                            select = Select(elem)
                            opciones_texto = [opt.text for opt in select.options]
                            
                            # Si tiene las opciones que esperamos (Compañía, Almacén, Ubicación)
                            if any(tipo_consulta.lower() in opt.lower() for opt in opciones_texto):
                                print("[OK] Dropdown 'Tipo de consulta' encontrado")
                                
                                # Seleccionar la opción especificada
                                for opcion in select.options:
                                    if tipo_consulta.lower() in opcion.text.lower():
                                        print(f"[*] Seleccionando opcion: '{opcion.text}'")
                                        select.select_by_visible_text(opcion.text)
                                        time.sleep(2)
                                        print(f"[OK] '{tipo_consulta}' seleccionado")
                                        dropdown_encontrado = True
                                        break
                                
                                if dropdown_encontrado:
                                    break
                        except:
                            continue
                
                if dropdown_encontrado:
                    break
            except:
                continue
        
        if not dropdown_encontrado:
            raise Exception(f"No se encontro el dropdown 'Tipo de consulta' o la opción '{tipo_consulta}'")
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Error seleccionando tipo consulta: {str(e)}")
        
        try:
            screenshot_path = os.path.join(os.path.dirname(__file__), f"error_tipo_consulta_{tipo_consulta}.png")
            driver.save_screenshot(screenshot_path)
            print(f"[*] Screenshot guardado en: {screenshot_path}")
        except:
            pass
        
        return False

def seleccionar_ubicacion_dropdown(driver, nombre_ubicacion):
    """Selecciona una ubicación específica del dropdown 'Ubicación'"""
    print(f"[*] Seleccionando ubicacion '{nombre_ubicacion}'...")
    
    # Timeout máximo para toda la función: 30 segundos
    import time
    tiempo_inicio = time.time()
    timeout_total = 30
    
    try:
        # Esperar a que el campo de ubicación esté disponible
        time.sleep(3)
        
        print("[*] Buscando el campo 'Ubicación'...")
        
        # PASO 1: Verificar cuántos SELECT hay
        print("[*] Verificando SELECT visibles...")
        selects_en_modal = driver.find_elements(By.XPATH, "//div[contains(@class, 'modal')]//select")
        selects_visibles = [s for s in selects_en_modal if s.is_displayed()]
        
        print(f"[*] {len(selects_visibles)} SELECT(s) visible(s)")
        
        # Si solo hay 1 SELECT, entonces el campo de Ubicación es un INPUT/combobox
        if len(selects_visibles) == 1:
            print("[*] El campo de Ubicación NO es un SELECT, buscando INPUT...")
            
            # Buscar el INPUT del campo "Ubicación"
            try:
                # Buscar todos los inputs visibles en el modal
                inputs = driver.find_elements(By.XPATH, 
                    "//div[contains(@class, 'modal')]//input[not(@type='hidden') and not(@type='checkbox') and not(@type='radio')]"
                )
                
                print(f"[*] {len(inputs)} INPUT(s) encontrado(s)")
                
                # El segundo input después del de fecha debería ser el de ubicación
                inputs_visibles = [inp for inp in inputs if inp.is_displayed()]
                
                print(f"[*] {len(inputs_visibles)} INPUT(s) visible(s)")
                
                campo_ubicacion = None
                
                # Método 1: Buscar por índice (el último input visible probablemente sea el de ubicación)
                if len(inputs_visibles) >= 2:
                    campo_ubicacion = inputs_visibles[-1]  # El último
                    print(f"[OK] Intentando con el último INPUT visible")
                elif len(inputs_visibles) == 1:
                    campo_ubicacion = inputs_visibles[0]
                    print(f"[OK] Solo hay 1 INPUT, usando ese")
                
                if campo_ubicacion:
                    print(f"[*] INPUT encontrado, intentando interactuar...")
                    
                    # Hacer scroll al elemento
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_ubicacion)
                    time.sleep(0.5)
                    
                    # Limpiar el campo
                    driver.execute_script("arguments[0].value = '';", campo_ubicacion)
                    time.sleep(0.3)
                    
                    # Hacer click para activar
                    print(f"[*] Haciendo click en el INPUT...")
                    try:
                        campo_ubicacion.click()
                    except:
                        driver.execute_script("arguments[0].click();", campo_ubicacion)
                    
                    time.sleep(1)
                    
                    # Escribir el nombre de la ubicación
                    print(f"[*] Escribiendo '{nombre_ubicacion}'...")
                    campo_ubicacion.send_keys(nombre_ubicacion)
                    time.sleep(2)
                    
                    # Buscar opciones del dropdown que aparecen
                    print("[*] Buscando opciones en dropdown...")
                    
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
                            time.sleep(0.5)
                            opciones = driver.find_elements(By.XPATH, selector)
                            opciones_visibles = [opt for opt in opciones if opt.is_displayed()]
                            
                            if opciones_visibles:
                                print(f"[*] {len(opciones_visibles)} opción(es) encontrada(s) con: {selector[:40]}")
                                opciones_encontradas = opciones_visibles
                                break
                        except:
                            continue
                    
                    # Si encontramos opciones, hacer click en la correcta
                    if opciones_encontradas:
                        for opcion in opciones_encontradas:
                            try:
                                texto_opcion = opcion.text.strip()
                                print(f"    - '{texto_opcion}'")
                                
                                if texto_opcion == nombre_ubicacion:
                                    print(f"[*] Haciendo click en '{texto_opcion}'")
                                    opcion.click()
                                    time.sleep(1)
                                    print(f"[OK] Ubicación '{nombre_ubicacion}' seleccionada")
                                    return True
                            except:
                                continue
                        
                        # Si no encontró coincidencia exacta, hacer click en la primera opción
                        print("[!] No se encontró coincidencia exacta")
                        print("[*] Intentando con la primera opción...")
                        try:
                            opciones_encontradas[0].click()
                            time.sleep(1)
                            print(f"[OK] Primera opción seleccionada")
                            return True
                        except:
                            pass
                    
                    # Si no aparecieron opciones, presionar ENTER
                    print("[*] No aparecieron opciones, presionando ENTER...")
                    from selenium.webdriver.common.keys import Keys
                    campo_ubicacion.send_keys(Keys.RETURN)
                    time.sleep(1)
                    print(f"[OK] ENTER presionado")
                    return True
                
            except Exception as e:
                print(f"[!] Error buscando INPUT: {e}")
        
        elif len(selects_visibles) >= 2:
            # Si hay 2 o más SELECT, usar el segundo
            print(f"[OK] Hay 2+ SELECT, usando el segundo como campo de ubicación")
            
            from selenium.webdriver.support.ui import Select
            select_ubicacion = selects_visibles[1]
            select = Select(select_ubicacion)
            
            print(f"[*] Opciones en SELECT:")
            for opt in select.options:
                print(f"    - '{opt.text.strip()}'")
            
            for opcion in select.options:
                if opcion.text.strip() == nombre_ubicacion:
                    print(f"[*] Seleccionando '{opcion.text.strip()}'")
                    select.select_by_visible_text(opcion.text.strip())
                    time.sleep(1)
                    print(f"[OK] Ubicación '{nombre_ubicacion}' seleccionada")
                    return True
            
            print(f"[!] '{nombre_ubicacion}' no encontrada")
            return False
        
        # Si llegamos aquí, no funcionó
        print("[ERROR] No se pudo seleccionar la ubicación")
        
        # Screenshot
        try:
            screenshot_path = os.path.join(os.path.dirname(__file__), f"error_ubicacion_{nombre_ubicacion}.png")
            driver.save_screenshot(screenshot_path)
            print(f"[*] Screenshot: {screenshot_path}")
        except:
            pass
        
        return False
        
    except Exception as e:
        tiempo_transcurrido = time.time() - tiempo_inicio
        print(f"[ERROR] Error después de {tiempo_transcurrido:.1f}s: {str(e)}")
        
        try:
            screenshot_path = os.path.join(os.path.dirname(__file__), f"error_timeout_{nombre_ubicacion}.png")
            driver.save_screenshot(screenshot_path)
            print(f"[*] Screenshot: {screenshot_path}")
        except:
            pass
        
        return False

def seleccionar_almacen_dropdown(driver, nombre_almacen):
    """Selecciona un almacén específico del combobox 'Almacén'"""
    print(f"[*] Seleccionando almacen '{nombre_almacen}'...")
    
    # Timeout máximo para toda la función: 30 segundos
    import time
    tiempo_inicio = time.time()
    timeout_total = 30
    
    try:
        # Esperar a que el campo de almacén esté disponible
        time.sleep(3)
        
        print("[*] Buscando el campo 'Almacén'...")
        
        # PASO 1: Verificar cuántos SELECT hay
        print("[*] Verificando SELECT visibles...")
        selects_en_modal = driver.find_elements(By.XPATH, "//div[contains(@class, 'modal')]//select")
        selects_visibles = [s for s in selects_en_modal if s.is_displayed()]
        
        print(f"[*] {len(selects_visibles)} SELECT(s) visible(s)")
        
        # Si solo hay 1 SELECT, entonces el campo de Almacén es un INPUT/combobox
        if len(selects_visibles) == 1:
            print("[*] El campo de Almacén NO es un SELECT, buscando INPUT...")
            
            # Buscar el INPUT del campo "Almacén"
            # Estrategia: Buscar inputs que NO sean hidden y que estén visibles
            
            try:
                # Buscar todos los inputs visibles en el modal
                inputs = driver.find_elements(By.XPATH, 
                    "//div[contains(@class, 'modal')]//input[not(@type='hidden') and not(@type='checkbox') and not(@type='radio')]"
                )
                
                print(f"[*] {len(inputs)} INPUT(s) encontrado(s)")
                
                # El segundo input después del de fecha debería ser el de almacén
                inputs_visibles = [inp for inp in inputs if inp.is_displayed()]
                
                print(f"[*] {len(inputs_visibles)} INPUT(s) visible(s)")
                
                # Buscar el input que tenga el atributo aria-label o esté después del label "Almacén"
                campo_almacen = None
                
                # Método 1: Buscar por índice (el último input visible probablemente sea el de almacén)
                if len(inputs_visibles) >= 2:
                    # Intentar con el segundo input visible (después del de fecha)
                    campo_almacen = inputs_visibles[-1]  # El último
                    print(f"[OK] Intentando con el último INPUT visible")
                elif len(inputs_visibles) == 1:
                    campo_almacen = inputs_visibles[0]
                    print(f"[OK] Solo hay 1 INPUT, usando ese")
                
                if campo_almacen:
                    print(f"[*] INPUT encontrado, intentando interactuar...")
                    
                    # Hacer scroll al elemento
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_almacen)
                    time.sleep(0.5)
                    
                    # Limpiar el campo
                    driver.execute_script("arguments[0].value = '';", campo_almacen)
                    time.sleep(0.3)
                    
                    # Hacer click para activar
                    print(f"[*] Haciendo click en el INPUT...")
                    try:
                        campo_almacen.click()
                    except:
                        driver.execute_script("arguments[0].click();", campo_almacen)
                    
                    time.sleep(1)
                    
                    # Escribir el nombre del almacén
                    print(f"[*] Escribiendo '{nombre_almacen}'...")
                    campo_almacen.send_keys(nombre_almacen)
                    time.sleep(2)
                    
                    # Buscar opciones del dropdown que aparecen
                    print("[*] Buscando opciones en dropdown...")
                    
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
                            time.sleep(0.5)
                            opciones = driver.find_elements(By.XPATH, selector)
                            opciones_visibles = [opt for opt in opciones if opt.is_displayed()]
                            
                            if opciones_visibles:
                                print(f"[*] {len(opciones_visibles)} opción(es) encontrada(s) con: {selector[:40]}")
                                opciones_encontradas = opciones_visibles
                                break
                        except:
                            continue
                    
                    # Si encontramos opciones, hacer click en la correcta
                    if opciones_encontradas:
                        for opcion in opciones_encontradas:
                            try:
                                texto_opcion = opcion.text.strip()
                                print(f"    - '{texto_opcion}'")
                                
                                if texto_opcion == nombre_almacen:
                                    print(f"[*] Haciendo click en '{texto_opcion}'")
                                    opcion.click()
                                    time.sleep(1)
                                    print(f"[OK] Almacén '{nombre_almacen}' seleccionado")
                                    return True
                            except:
                                continue
                        
                        # Si no encontró coincidencia exacta, hacer click en la primera opción
                        print("[!] No se encontró coincidencia exacta")
                        print("[*] Intentando con la primera opción...")
                        try:
                            opciones_encontradas[0].click()
                            time.sleep(1)
                            print(f"[OK] Primera opción seleccionada")
                            return True
                        except:
                            pass
                    
                    # Si no aparecieron opciones, presionar ENTER
                    print("[*] No aparecieron opciones, presionando ENTER...")
                    from selenium.webdriver.common.keys import Keys
                    campo_almacen.send_keys(Keys.RETURN)
                    time.sleep(1)
                    print(f"[OK] ENTER presionado")
                    return True
                
            except Exception as e:
                print(f"[!] Error buscando INPUT: {e}")
        
        elif len(selects_visibles) >= 2:
            # Si hay 2 o más SELECT, usar el segundo
            print(f"[OK] Hay 2+ SELECT, usando el segundo como campo de almacén")
            
            from selenium.webdriver.support.ui import Select
            select_almacen = selects_visibles[1]
            select = Select(select_almacen)
            
            print(f"[*] Opciones en SELECT:")
            for opt in select.options:
                print(f"    - '{opt.text.strip()}'")
            
            for opcion in select.options:
                if opcion.text.strip() == nombre_almacen:
                    print(f"[*] Seleccionando '{opcion.text.strip()}'")
                    select.select_by_visible_text(opcion.text.strip())
                    time.sleep(1)
                    print(f"[OK] Almacén '{nombre_almacen}' seleccionado")
                    return True
            
            print(f"[!] '{nombre_almacen}' no encontrado")
            return False
        
        # Si llegamos aquí, no funcionó
        print("[ERROR] No se pudo seleccionar el almacén")
        
        # Screenshot
        try:
            screenshot_path = os.path.join(os.path.dirname(__file__), f"error_almacen_{nombre_almacen}.png")
            driver.save_screenshot(screenshot_path)
            print(f"[*] Screenshot: {screenshot_path}")
        except:
            pass
        
        return False
        
    except Exception as e:
        tiempo_transcurrido = time.time() - tiempo_inicio
        print(f"[ERROR] Error después de {tiempo_transcurrido:.1f}s: {str(e)}")
        
        try:
            screenshot_path = os.path.join(os.path.dirname(__file__), f"error_timeout_{nombre_almacen}.png")
            driver.save_screenshot(screenshot_path)
            print(f"[*] Screenshot: {screenshot_path}")
        except:
            pass
        
        return False

def generar_xlsx(driver):
    """Hace click en el botón 'Generar XLSX'"""
    print("[*] Haciendo click en 'Generar XLSX'...")
    wait = WebDriverWait(driver, 10)
    
    try:
        # Buscar el botón "Generar XLSX"
        selectores_boton = [
            "//button[contains(text(), 'Generar XLSX')]",
            "//button[contains(., 'Generar XLSX')]",
            "//button[contains(@class, 'btn-primary') and contains(., 'Generar')]",
        ]
        
        for selector in selectores_boton:
            try:
                print(f"[*] Buscando boton con: {selector[:60]}...")
                boton = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                boton.click()
                time.sleep(2)
                print("[OK] 'Generar XLSX' clickeado - Descarga iniciada")
                return True
            except:
                continue
        
        raise Exception("No se encontro el boton 'Generar XLSX'")
        
    except Exception as e:
        print(f"[ERROR] Error generando XLSX: {str(e)}")
        return False

def esperar_descarga_archivo(carpeta, timeout=300, archivos_existentes_previos=None):
    """Espera a que se complete la descarga del archivo"""
    print(f"[*] Esperando descarga del archivo (maximo {timeout//60} minutos)...")
    
    tiempo_inicio = time.time()
    ultimo_reporte = tiempo_inicio
    
    # Obtener lista de archivos ANTES de la descarga con sus timestamps
    # Si se pasó un snapshot previo, usarlo; si no, tomar uno ahora
    if archivos_existentes_previos is not None:
        archivos_existentes = archivos_existentes_previos
        print(f"[*] Usando snapshot previo de archivos existentes: {len(archivos_existentes)}")
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
        print(f"[*] Archivos existentes antes: {len(archivos_existentes)}")
    
    for archivo in archivos_existentes.keys():
        print(f"    - {os.path.basename(archivo)}")
    
    while time.time() - tiempo_inicio < timeout:
        # Reportar progreso cada 5 segundos
        tiempo_actual = time.time()
        if tiempo_actual - ultimo_reporte >= 5:
            segundos_transcurridos = int(tiempo_actual - tiempo_inicio)
            print(f"[*] Esperando... {segundos_transcurridos}s transcurridos")
            ultimo_reporte = tiempo_actual
        
        # PASO 1: Verificar si hay archivos temporales (descarga en progreso)
        archivos_temp = glob.glob(os.path.join(carpeta, "*.crdownload"))
        archivos_temp += glob.glob(os.path.join(carpeta, "*.tmp"))
        
        if archivos_temp:
            # Hay descarga en progreso, seguir esperando
            time.sleep(1)  # Verificar cada 1 segundo
            continue
        
        # PASO 2: Buscar CUALQUIER archivo .xlsx/.xls nuevo o modificado
        try:
            archivos_actuales = {}
            for archivo in glob.glob(os.path.join(carpeta, "*.xlsx")):
                if not os.path.basename(archivo).startswith("~$"):
                    archivos_actuales[archivo] = os.path.getmtime(archivo)
            for archivo in glob.glob(os.path.join(carpeta, "*.xls")):
                if not os.path.basename(archivo).startswith("~$"):
                    archivos_actuales[archivo] = os.path.getmtime(archivo)
            
            # Buscar archivos que NO existían antes O que fueron modificados
            for archivo, mtime_actual in archivos_actuales.items():
                # Caso 1: Archivo completamente nuevo
                if archivo not in archivos_existentes:
                    edad = time.time() - mtime_actual
                    print(f"[*] Archivo nuevo detectado: {os.path.basename(archivo)} (edad: {int(edad)}s)")
                    
                    # Aceptar archivos de hasta 90 segundos (1.5 minutos de margen)
                    if edad < 90:
                        print(f"[OK] Archivo descargado: {os.path.basename(archivo)}")
                        return archivo
                    else:
                        print(f"[!] Archivo muy antiguo, ignorando")
                
                # Caso 2: Archivo que fue modificado recientemente
                elif mtime_actual > archivos_existentes[archivo]:
                    edad = time.time() - mtime_actual
                    print(f"[*] Archivo modificado: {os.path.basename(archivo)} (edad: {int(edad)}s)")
                    
                    if edad < 90:
                        print(f"[OK] Archivo descargado (modificado): {os.path.basename(archivo)}")
                        return archivo
        
        except Exception as e:
            print(f"[!] Error verificando archivos: {e}")
        
        time.sleep(1)  # Verificar cada 1 segundo (más agresivo)
    
    # Timeout alcanzado
    print("[!] Timeout esperando la descarga")
    
    # ÚLTIMA OPORTUNIDAD: Verificar una vez más por si el archivo se terminó de descargar
    print("[*] Verificación final de archivos...")
    time.sleep(2)  # Esperar 2 segundos más por si acaso
    
    try:
        archivos_finales = {}
        for archivo in glob.glob(os.path.join(carpeta, "*.xlsx")):
            if not os.path.basename(archivo).startswith("~$"):
                archivos_finales[archivo] = os.path.getmtime(archivo)
        for archivo in glob.glob(os.path.join(carpeta, "*.xls")):
            if not os.path.basename(archivo).startswith("~$"):
                archivos_finales[archivo] = os.path.getmtime(archivo)
        
        # Buscar archivos nuevos o modificados en esta verificación final
        for archivo, mtime in archivos_finales.items():
            if archivo not in archivos_existentes:
                edad = time.time() - mtime
                if edad < 90:  # Hasta 90 segundos de edad
                    print(f"[OK] Archivo detectado en verificación final: {os.path.basename(archivo)}")
                    return archivo
            elif mtime > archivos_existentes[archivo]:
                edad = time.time() - mtime
                if edad < 90:
                    print(f"[OK] Archivo modificado detectado en verificación final: {os.path.basename(archivo)}")
                    return archivo
    except Exception as e:
        print(f"[!] Error en verificación final: {e}")
    
    # Listar TODOS los archivos para debugging
    print("[*] TODOS los archivos al finalizar:")
    try:
        archivos_finales = {}
        for archivo in glob.glob(os.path.join(carpeta, "*.xlsx")):
            if not os.path.basename(archivo).startswith("~$"):
                archivos_finales[archivo] = os.path.getmtime(archivo)
        for archivo in glob.glob(os.path.join(carpeta, "*.xls")):
            if not os.path.basename(archivo).startswith("~$"):
                archivos_finales[archivo] = os.path.getmtime(archivo)
        
        for archivo, mtime in archivos_finales.items():
            edad = time.time() - mtime
            if archivo in archivos_existentes:
                if mtime > archivos_existentes[archivo]:
                    estado = "MODIFICADO"
                else:
                    estado = "EXISTENTE"
            else:
                estado = "NUEVO"
            print(f"    - {os.path.basename(archivo)} (edad: {int(edad)}s) [{estado}]")
    except:
        pass
    
    return None

def renombrar_y_mover_archivo(archivo_original, nuevo_nombre, carpeta_destino):
    """Renombra el archivo y lo mueve a la carpeta de destino si es necesario"""
    print(f"[*] Renombrando archivo a: {nuevo_nombre}")
    
    try:
        nueva_ruta = os.path.join(carpeta_destino, nuevo_nombre)
        
        # Si el archivo ya existe en el destino, eliminarlo
        if os.path.exists(nueva_ruta):
            print(f"[*] Eliminando archivo existente: {nuevo_nombre}")
            os.remove(nueva_ruta)
        
        # Si el archivo original está en otra carpeta, moverlo
        if os.path.dirname(archivo_original) != carpeta_destino:
            print(f"[*] Moviendo archivo con shutil.move()...")
            import shutil
            shutil.move(archivo_original, nueva_ruta)
        else:
            # Si ya está en la carpeta correcta, solo renombrar
            print(f"[*] Renombrando archivo con os.rename()...")
            os.rename(archivo_original, nueva_ruta)
        
        print(f"[OK] Archivo guardado como: {nuevo_nombre}")
        return nueva_ruta
        
    except Exception as e:
        print(f"[ERROR] Error renombrando archivo: {str(e)}")
        import traceback
        traceback.print_exc()
        return archivo_original

def cerrar_modal(driver):
    """Cierra el modal para preparar la siguiente descarga"""
    print("[*] Cerrando modal...")
    
    try:
        # Buscar botón de cerrar (X)
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
                        time.sleep(2)
                        print("[OK] Modal cerrado")
                        return True
            except:
                continue
        
        # Si no encontró botón X, presionar ESC
        from selenium.webdriver.common.keys import Keys
        from selenium.webdriver.common.action_chains import ActionChains
        
        actions = ActionChains(driver)
        actions.send_keys(Keys.ESCAPE).perform()
        time.sleep(2)
        print("[OK] Modal cerrado con ESC")
        return True
        
    except Exception as e:
        print(f"[!] No se pudo cerrar modal: {str(e)}")
        return False

def enviar_email_notificacion(exito=True, archivos_descargados=None, tiempo_total=None, error=None):
    """Envía un email de notificación al finalizar el proceso"""
    if not EMAIL_CONFIG.get("enabled", False):
        print("[*] Envio de email desactivado")
        return False
    
    print("[*] Enviando notificacion por email...")
    
    try:
        msg = MIMEMultipart('alternative')
        msg['From'] = EMAIL_CONFIG['sender_email']
        msg['To'] = ', '.join(EMAIL_CONFIG['recipient_emails'])
        
        fecha_actual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        if exito:
            msg['Subject'] = f"✅ FERTRAC: Descarga de Valorizados Completada - {datetime.now().strftime('%d/%m/%Y')}"
            
            # Lista de archivos en HTML
            archivos_html = ""
            if archivos_descargados:
                for archivo in archivos_descargados:
                    archivos_html += f"<li>{os.path.basename(archivo)}</li>"
            
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
                    ul {{ margin: 10px 0; padding-left: 20px; }}
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="header">
                        <h2>✅ Descarga de Valorizados Completada</h2>
                    </div>
                    <div class="content">
                        <div class="info">
                            <span class="label">📅 Fecha y hora:</span>
                            <span class="value">{fecha_actual}</span>
                        </div>
                        <div class="info">
                            <span class="label">📁 Archivos descargados:</span>
                            <ul>{archivos_html}</ul>
                        </div>
                        <div class="info">
                            <span class="label">📂 Ubicación:</span>
                            <span class="value">{RUTA_DESCARGA}</span>
                        </div>
                        <div class="info">
                            <span class="label">⏱️ Tiempo total:</span>
                            <span class="value">{tiempo_total if tiempo_total else 'N/A'}</span>
                        </div>
                        <div class="info">
                            <span class="label">🎯 Proceso:</span>
                            <span class="value">Descarga de Informes Valorizados por Almacén</span>
                        </div>
                    </div>
                    <div class="footer">
                        <p>Este es un mensaje automático generado por el sistema de automatización FERTRAC.</p>
                        <p>Script: descargar_valorizado_almacen_fertrac.py</p>
                    </div>
                </div>
            </body>
            </html>
            """
        else:
            msg['Subject'] = f"❌ FERTRAC: Error en Descarga de Valorizados - {datetime.now().strftime('%d/%m/%Y')}"
            
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
                        <h2>❌ Error en Descarga de Valorizados</h2>
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
                            <span class="value">Descarga de Informes Valorizados por Almacén</span>
                        </div>
                    </div>
                    <div class="footer">
                        <p>Este es un mensaje automático generado por el sistema de automatización FERTRAC.</p>
                        <p>Script: descargar_valorizado_almacen_fertrac.py</p>
                    </div>
                </div>
            </body>
            </html>
            """
        
        part = MIMEText(html, 'html')
        msg.attach(part)
        
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
    """Función principal que ejecuta todo el proceso"""
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
        
        # Crear carpeta de destino
        carpeta_destino = crear_carpeta_destino()
        print(f"Carpeta de destino: {carpeta_destino}")
        print("-" * 70)
        
        # Inicializar driver
        driver = configurar_driver(carpeta_destino)
        
        # Login
        if not hacer_login(driver):
            raise Exception("Fallo en el login")
        
        # Navegar a inventario
        navegar_a_inventario(driver)
        
        # Abrir menú Informes
        if not abrir_menu_informes(driver):
            raise Exception("Fallo al abrir menu Informes")
        
        # Seleccionar Valorizado
        if not seleccionar_valorizado(driver):
            raise Exception("Fallo al seleccionar Valorizado")
        
        # Descargar cada informe
        print("\n" + "=" * 70)
        print("DESCARGANDO INFORMES")
        print("=" * 70)
        
        for i, config in enumerate(ALMACENES_CONFIG, 1):
            print(f"\n[*] Procesando {i}/{len(ALMACENES_CONFIG)}: {config['nombre_archivo']}")
            print("-" * 70)
            
            # PASO 1: Seleccionar tipo de consulta (solo si no es None)
            if config['tipo_consulta'] is not None:
                if not seleccionar_tipo_consulta(driver, config['tipo_consulta']):
                    print(f"[!] No se pudo seleccionar tipo consulta '{config['tipo_consulta']}', continuando...")
                    continue
            else:
                # Para el primer informe, Compañía ya está predeterminado
                print("[*] Usando tipo de consulta predeterminado (Compañía)")
                time.sleep(3)  # Esperar unos segundos antes de continuar
            
            # PASO 2: Si es Ubicación, seleccionar la ubicación específica
            if config['tipo_consulta'] == "Ubicación" and config['ubicacion']:
                if not seleccionar_ubicacion_dropdown(driver, config['ubicacion']):
                    print(f"[!] No se pudo seleccionar ubicacion '{config['ubicacion']}', continuando...")
                    continue
            
            # PASO 3: Tomar snapshot de archivos ANTES de iniciar la descarga
            print("[*] Tomando snapshot de archivos antes de descarga...")
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
            
            # PASO 4: Generar XLSX (inicia la descarga)
            if not generar_xlsx(driver):
                print(f"[!] No se pudo generar XLSX para '{config['nombre_archivo']}', continuando...")
                continue
            
            # PASO 5: Esperar descarga (usando el snapshot tomado ANTES)
            archivo_descargado = esperar_descarga_archivo(
                carpeta_destino, 
                timeout=TIMEOUT_DESCARGA,
                archivos_existentes_previos=archivos_antes_descarga
            )
            
            if archivo_descargado:
                # Esperar 3 segundos antes de renombrar para asegurar que el archivo
                # esté completamente cerrado por el navegador
                print("[*] Esperando 3 segundos antes de renombrar...")
                time.sleep(3)
                
                # Renombrar archivo
                archivo_final = renombrar_y_mover_archivo(
                    archivo_descargado,
                    config['nombre_archivo'],
                    carpeta_destino
                )
                archivos_descargados.append(archivo_final)
                print(f"[OK] {config['nombre_archivo']} descargado y renombrado")
                
            else:
                print(f"[!] Timeout esperando descarga de '{config['nombre_archivo']}'")
                
                # Screenshot para debugging
                try:
                    screenshot_path = os.path.join(os.path.dirname(__file__), f"error_descarga_{config['nombre_archivo'].replace('.xlsx', '')}.png")
                    driver.save_screenshot(screenshot_path)
                    print(f"[*] Screenshot guardado: {screenshot_path}")
                except:
                    pass
            
            # Si no es el último, cerrar modal y volver a abrir para siguiente descarga
            if i < len(ALMACENES_CONFIG):
                print("[*] Preparando para siguiente descarga...")
                cerrar_modal(driver)
                time.sleep(2)
                
                # Reabrir menú para siguiente descarga
                if not abrir_menu_informes(driver):
                    print("[!] No se pudo reabrir menu Informes")
                    break
                
                if not seleccionar_valorizado(driver):
                    print("[!] No se pudo reabrir Valorizado")
                    break
        
        # Calcular tiempo total
        tiempo_total_segundos = int(time.time() - tiempo_inicio_total)
        tiempo_total_minutos = tiempo_total_segundos // 60
        tiempo_total_texto = f"{tiempo_total_minutos} minutos {tiempo_total_segundos % 60} segundos"
        
        print("\n" + "=" * 70)
        print("PROCESO COMPLETADO")
        print("=" * 70)
        print(f"Archivos descargados: {len(archivos_descargados)}/{len(ALMACENES_CONFIG)}")
        for archivo in archivos_descargados:
            print(f"  - {os.path.basename(archivo)}")
        print(f"Ubicacion: {carpeta_destino}")
        print(f"Tiempo total: {tiempo_total_texto}")
        print("=" * 70)
        
        # Enviar email de notificación
        enviar_email_notificacion(
            exito=True,
            archivos_descargados=archivos_descargados,
            tiempo_total=tiempo_total_texto
        )
        
        print("\n[*] Cerrando navegador en 5 segundos...")
        time.sleep(5)
        
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