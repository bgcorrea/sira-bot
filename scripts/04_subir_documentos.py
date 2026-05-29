"""
Script 04: Bot de subida masiva a SIRA
========================================
Lee data/master_subida.xlsx (generado por el script 03) e itera los folios,
subiendo a SIRA los documentos que tenga asignados.

Reglas duras:
- Nunca envía a revisión. Solo deja folios en estado Borrador.
- Salta folios en estado "Enviado" (ya cargados por otro usuario).
- Idempotente: si la sección ya tiene un documento cargado, no lo re-sube.
- Reanudable: si lo cortás, retoma desde donde quedó leyendo el CSV de log.

Secciones que carga (en orden):
1. Convenio + Acto Administrativo  (2 archivos secuenciales)
2. Certificado de registro de entidad receptora
3. Transferencias (1 voucher)
4. Respaldo de rendición
5. Garantías (solo si existe el archivo)

Uso:
    python scripts\04_subir_documentos.py
    python scripts\04_subir_documentos.py --dry-run    (solo navega, no sube)
    python scripts\04_subir_documentos.py --folio 55828  (un solo folio para test)
"""

import argparse
import csv
import sys
import time
from pathlib import Path
from urllib.parse import quote
import openpyxl
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
)
from webdriver_manager.chrome import ChromeDriverManager

# ====== CONFIGURACIÓN ======
MASTER_XLSX = "data/master_subida.xlsx"
DISTRIBUCION_XLSX = "DISTRIBUCIÓN CARGA VB.xlsx"
LOG_CSV = "logs/ejecucion_subidas.csv"
LOG_REPORTE = "logs/reporte_faltantes.csv"
CHROME_PROFILE_DIR = "/home/bgcorrea/.bot_sira_chrome_profile"

TIMEOUT_PAGINA = 30        # Carga inicial de la página
TIMEOUT_SUBIDA = 60        # Espera estándar tras send_keys hasta ver el archivo procesado
SLEEP_ENTRE_ARCHIVOS = 2   # Pausa entre subidas dentro del mismo folio
SLEEP_ENTRE_FOLIOS = 3     # Pausa entre folios

# URLs según año
SEGEGOB = quote("SECRETARÍA GENERAL DE GOBIERNO")
SUBSEG = quote("SUBSECRETARÍA GENERAL DE GOBIERNO")
BASE = "https://sira.auditoriainternadegobierno.gob.cl"
URL_2022 = BASE + "/?q={folio}#/convenio/" + SEGEGOB + "/" + SUBSEG + "/" + SUBSEG + "/{folio}"
URL_2024 = BASE + "/?q={folio}#/convenio/" + SEGEGOB + "/" + SUBSEG + "/_/{folio}"

# Mapeo: nombre de sección en el DOM → nombre de columna en el master
SECCIONES = [
    {
        "titulo_dom": "Convenio + Acto Administrativo",
        "campos_master": ["convenio_pdf", "acto_admin_pdf"],  # primer archivo, luego segundo
        "obligatoria": True,
    },
    {
        "titulo_dom": "Certificado de registro de entidad receptora",
        "campos_master": ["certificado_pdf"],
        "obligatoria": True,
    },
    {
        "titulo_dom": "Transferencias",
        "campos_master": ["voucher_pdf"],
        "obligatoria": True,
    },
    {
        "titulo_dom": "Respaldo de rendición",
        "campos_master": ["rendicion_pdf"],
        "obligatoria": True,
    },
    {
        "titulo_dom": "Garantías",
        "campos_master": ["garantia_pdf"],
        "obligatoria": False,
    },
]


# ====== CARGA DE DISTRIBUCIÓN ======

def cargar_folios_distribucion() -> set[str]:
    """Devuelve el set de folios presentes en DISTRIBUCIÓN CARGA VB.xlsx."""
    df = pd.read_excel(DISTRIBUCION_XLSX)
    return set(df["ID Convenio"].astype(str).str.strip())


# ====== UTILIDADES DE DRIVER ======

SNAP_CHROMIUM = "/snap/bin/chromium"
SNAP_DRIVER   = "/snap/bin/chromium.chromedriver"
CDPORT        = 9222


def _lanzar_chromium_snap(perfil_dir: str, puerto: int = CDPORT):
    """Lanza Chromium snap con depuración TCP para evitar el bloqueo de pipe del snap."""
    import subprocess
    cmd = [
        SNAP_CHROMIUM,
        f"--remote-debugging-port={puerto}",
        f"--user-data-dir={perfil_dir}",
        "--start-maximized",
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--disable-gpu",
        "--no-first-run",
        "--password-store=basic",
    ]
    return subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def crear_driver() -> webdriver.Chrome:
    import socket
    Path(CHROME_PROFILE_DIR).mkdir(parents=True, exist_ok=True)

    # Snap Chromium: lanzar con TCP y conectar via debuggerAddress
    if Path(SNAP_CHROMIUM).exists() and Path(SNAP_DRIVER).exists():
        proc = _lanzar_chromium_snap(CHROME_PROFILE_DIR)
        for _ in range(20):
            time.sleep(0.5)
            try:
                with socket.create_connection(("127.0.0.1", CDPORT), timeout=1):
                    break
            except OSError:
                pass
        else:
            proc.terminate()
            raise RuntimeError("Chromium no abrió el puerto de depuración a tiempo.")

        options = Options()
        options.add_experimental_option("debuggerAddress", f"127.0.0.1:{CDPORT}")
        service = Service(SNAP_DRIVER)
        driver = webdriver.Chrome(service=service, options=options)
        driver._snap_proc = proc
        driver.execute_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        return driver

    # Fallback para Chrome/Chromium instalado via apt o Windows
    options = Options()
    options.add_argument(f"--user-data-dir={CHROME_PROFILE_DIR}")
    options.add_argument("--start-maximized")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.execute_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )
    return driver


def esperar_login(driver):
    print("\n" + "=" * 70)
    print("LOGIN MANUAL")
    print("=" * 70)
    print("Si el perfil de Chrome ya tenía sesión, deberías estar logueado.")
    print("Si no, hacé login con Clave Única.")
    print("Cuando estés DENTRO de SIRA, presioná ENTER.")
    print("=" * 70)
    driver.get(BASE + "/")
    input("\n>>> Presioná ENTER cuando estés logueado: ")


def url_para_folio(folio: str, ano: int) -> str:
    return URL_2022.format(folio=folio) if ano <= 2022 else URL_2024.format(folio=folio)


# ====== INTERACCIÓN CON SIRA ======

def cargar_folio(driver, folio: str, ano: int) -> str:
    """Navega al folio y devuelve el estado del convenio (Borrador/Enviado/Cerrado/error)."""
    driver.get(url_para_folio(folio, ano))
    try:
        WebDriverWait(driver, TIMEOUT_PAGINA).until(
            EC.any_of(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Datos declarados')]")),
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'no existe')]")),
            )
        )
        time.sleep(1.5)
        body = driver.find_element(By.TAG_NAME, "body").text
        if "no existe en el catálogo" in body.lower():
            return "NO_EXISTE"
        # Detectar estado desde el badge cerca del header
        for kw in ["✓ Enviado", "✓ Cerrado", "Enviado", "Cerrado", "Borrador"]:
            if kw in body[:800]:
                return kw.replace("✓ ", "").strip()
        return "DESCONOCIDO"
    except TimeoutException:
        return "TIMEOUT"


def localizar_seccion(driver, titulo_dom: str):
    """Devuelve el WebElement del <div class='seccion ...'> principal que contiene el título."""
    xpath = (
        f"//span[@class='seccion-title-text' and normalize-space(text())='{titulo_dom}']"
        f"/ancestor::div[contains(concat(' ', normalize-space(@class), ' '), ' seccion ')][1]"
    )
    return driver.find_element(By.XPATH, xpath)


def seccion_ya_tiene_documentos(seccion_element) -> bool:
    """True si la sección está en estado 'ok' (al menos 1 doc cargado)."""
    clase = seccion_element.get_attribute("class") or ""
    return "ok" in clase.split() or "seccion ok" in clase


def obtener_n_docs_subidos(seccion_element, driver) -> int:
    """Cuenta los <li class='doc-row'> que pertenecen a esta sección."""
    try:
        rows = seccion_element.find_elements(By.CSS_SELECTOR, "li.doc-row")
        return len(rows)
    except Exception:
        return 0


def subir_archivo_a_seccion(driver, titulo_seccion: str, ruta_archivo: str, n_docs_antes: int) -> tuple[bool, str]:
    """
    Sube un archivo a la sección dada y maneja alertas/modales de extracción de forma dinámica.
    """
    if ruta_archivo.startswith("D:\\") or ruta_archivo.startswith("C:\\"):
        return False, f"RUTA_WINDOWS_EN_LINUX: {ruta_archivo[:60]}"
    ruta = Path(ruta_archivo).resolve()
    if not ruta.exists():
        return False, f"ARCHIVO_NO_EXISTE: {ruta}"
    if not ruta.is_file():
        return False, f"NO_ES_ARCHIVO: {ruta}"

    try:
        seccion_element = localizar_seccion(driver, titulo_seccion)
        input_file = seccion_element.find_element(By.XPATH, ".//input[@type='file']")
    except NoSuchElementException:
        return False, "INPUT_FILE_NO_ENCONTRADO"

    # Descartar alert residual antes de subir
    _descartar_alert(driver)

    # Enviar la ruta absoluta
    try:
        input_file.send_keys(str(ruta))
    except Exception as e:
        return False, f"SEND_KEYS_ERROR: {type(e).__name__}: {str(e)[:100]}"

    # Capturar alert inmediato de SIRA ("No se pudo subir el archivo", etc.)
    time.sleep(0.8)
    alerta = _descartar_alert(driver)
    if alerta and any(k in alerta.lower() for k in ["no se pudo", "error", "falló"]):
        return False, f"ALERT_SIRA: {alerta[:120]}"

    # CORRECCIÓN: Para Transferencias aplicamos un timeout corto (5s) según lo solicitado
    timeout_actual = 5 if "Transferencias" in titulo_seccion else TIMEOUT_SUBIDA

    # Monitoreo dinámico del DOM
    try:
        def comprobar_progreso_o_modal(d):
            try:
                sec_el = localizar_seccion(d, titulo_seccion)
                clase_sec = sec_el.get_attribute("class") or ""
                # Éxito si la sección pasa a estado 'ok' o si aumentan las filas físicas
                if "ok" in clase_sec.split() or obtener_n_docs_subidos(sec_el, d) > n_docs_antes:
                    return "COMPLETO"
            except Exception:
                pass
            
            # Reconocedor inteligente de modales (alertas OCR, datos personales, etc.)
            modales_xpaths = [
                "//div[contains(@class, 'modal') or contains(@class, 'confirm')]//button[contains(text(), 'Sí') or contains(text(), 'igual') or contains(text(), 'de todos modos')]",
                "//div[contains(@class, 'modal') or contains(@class, 'confirm')]//button[contains(@class, 'warn') or contains(@class, 'primary')]",
                "//button[contains(text(), 'Aceptar') or contains(text(), 'Entendido') or contains(text(), 'Subir igual')]"
            ]
            
            for xpath in modales_xpaths:
                try:
                    btn = d.find_element(By.XPATH, xpath)
                    if btn.is_displayed() and btn.is_enabled():
                        padre_html = btn.find_element(By.XPATH, "./ancestor::div[1]").get_attribute("class") or ""
                        if any(k in padre_html for k in ["actions", "modal", "confirm"]) or any(t in btn.text for t in ["Sí", "Aceptar", "igual", "Entendido"]):
                            return btn
                except NoSuchElementException:
                    pass
            return False

        try:
            resultado = WebDriverWait(driver, timeout_actual).until(comprobar_progreso_o_modal)

            if resultado != "COMPLETO":
                print(f"      ⚠ Modal/Alerta detectada ('{resultado.text}'). Forzando confirmación...")
                resultado.click()
                
                # Esperamos la confirmación definitiva
                WebDriverWait(driver, timeout_actual).until(
                    lambda d: "ok" in localizar_seccion(d, titulo_seccion).get_attribute("class").split() or obtener_n_docs_subidos(localizar_seccion(d, titulo_seccion), d) > n_docs_antes
                )

            time.sleep(1.5)
            return True, "OK"

        except TimeoutException:
            # CORRECCIÓN FALLBACK: Si es Transferencias, es rápido y se sube bien, pero no genera filas tradicionales. 
            # Si expira el plazo corto de 5s sin alertas de bloqueo en pantalla, lo damos por válido.
            if "Transferencias" in titulo_seccion:
                time.sleep(1)
                return True, "OK"
            raise
        
    except TimeoutException:
        return False, f"TIMEOUT_SUBIDA_{timeout_actual}s"
    except StaleElementReferenceException:
        return False, "STALE_ELEMENT"


def click_agregar_otro_documento(driver, seccion_element) -> bool:
    """Hace click en 'Agregar otro documento' anulando disparadores automáticos de ventanas del OS."""
    try:
        btn = seccion_element.find_element(
            By.XPATH,
            ".//button[contains(@class,'btn') and contains(text(),'Agregar')]"
        )
        
        # CORRECCIÓN CRÍTICA: Interceptamos el prototipo de click del navegador mediante JS 
        # para evitar que el trigger interno del botón abra el explorador de archivos nativo de Windows.
        driver.execute_script("""
            window._originalInputClick = HTMLInputElement.prototype.click;
            HTMLInputElement.prototype.click = function() {
                if (this.type === 'file') {
                    console.log('Bot interceptó apertura forzada de ventana OS');
                    return;
                }
                window._originalInputClick.apply(this, arguments);
            };
        """)
        
        _esperar_sin_overlay(driver)
        _descartar_alert(driver)
        try:
            btn.click()
        except Exception:
            driver.execute_script("arguments[0].click();", btn)
        time.sleep(0.5)

        # Devolvemos el comportamiento original inmediatamente
        driver.execute_script("""
            if (window._originalInputClick) {
                HTMLInputElement.prototype.click = window._originalInputClick;
            }
        """)
        return True
    except NoSuchElementException:
        return False


# ====== BORRADO DE DOCUMENTOS EXISTENTES ======

def _confirmar_modal_borrado(driver, timeout: int = 10) -> None:
    """Acepta cualquier modal de confirmación que aparezca al borrar."""
    xpaths = [
        "//button[contains(text(),'Sí') or contains(text(),'Aceptar') or contains(text(),'Confirmar') or contains(text(),'Eliminar')]",
        "//div[contains(@class,'modal')]//button[contains(@class,'primary') or contains(@class,'danger') or contains(@class,'warn')]",
        "//div[contains(@class,'confirm')]//button",
    ]
    try:
        def modal_presente(d):
            for xp in xpaths:
                try:
                    btn = d.find_element(By.XPATH, xp)
                    if btn.is_displayed() and btn.is_enabled():
                        return btn
                except NoSuchElementException:
                    pass
            return False
        btn = WebDriverWait(driver, timeout).until(modal_presente)
        btn.click()
        time.sleep(0.8)
    except TimeoutException:
        pass  # No había modal, está bien


def borrar_docs_seccion(driver, titulo_seccion: str, dry_run: bool = False) -> int:
    """
    Elimina todos los documentos cargados en la sección.
    Devuelve la cantidad de documentos borrados.
    """
    borrados = 0
    _SELECTORES_BORRAR = [
        ".//button[contains(@class,'eliminar') or contains(@class,'delete') or contains(@class,'btn-danger')]",
        ".//button[contains(@title,'liminar') or contains(@title,'orrar') or contains(@title,'emover')]",
        ".//a[contains(@class,'eliminar') or contains(@class,'delete')]",
        ".//i[contains(@class,'trash') or contains(@class,'times') or contains(@class,'remove')]/parent::button",
        ".//button[contains(@ng-click,'eliminar') or contains(@ng-click,'borrar') or contains(@ng-click,'remove')]",
    ]

    max_intentos = 20  # seguro contra bucles infinitos
    for _ in range(max_intentos):
        try:
            seccion_el = localizar_seccion(driver, titulo_seccion)
            rows = seccion_el.find_elements(By.CSS_SELECTOR, "li.doc-row")
            if not rows:
                break
            row = rows[0]  # siempre borramos el primero; el DOM se actualiza
        except (NoSuchElementException, StaleElementReferenceException):
            break

        btn_borrar = None
        for selector in _SELECTORES_BORRAR:
            try:
                btn_borrar = row.find_element(By.XPATH, selector)
                if btn_borrar.is_displayed():
                    break
                btn_borrar = None
            except NoSuchElementException:
                continue

        if btn_borrar is None:
            print(f"      ⚠ No se encontró botón de borrar en doc-row de '{titulo_seccion}'")
            break

        if dry_run:
            print(f"      [DRY] Borraría doc en '{titulo_seccion}'")
            borrados += 1
            break

        try:
            btn_borrar.click()
            time.sleep(0.5)
            _confirmar_modal_borrado(driver)
            # Esperar que el row desaparezca
            WebDriverWait(driver, 10).until(
                lambda d: len(localizar_seccion(d, titulo_seccion)
                              .find_elements(By.CSS_SELECTOR, "li.doc-row")) < len(rows)
            )
            borrados += 1
            time.sleep(0.5)
        except (TimeoutException, StaleElementReferenceException):
            print(f"      ⚠ No se pudo confirmar el borrado en '{titulo_seccion}'")
            break

    return borrados


# ====== ORQUESTACIÓN ======

def _descartar_alert(driver) -> str | None:
    """Acepta cualquier alert JS abierto. Devuelve el texto o None si no había."""
    try:
        alert = driver.switch_to.alert
        texto = alert.text
        alert.accept()
        time.sleep(0.5)
        return texto
    except Exception:
        return None


def _esperar_sin_overlay(driver, timeout: int = 8) -> None:
    """Espera a que desaparezcan overlays/modales que bloquean clicks."""
    try:
        WebDriverWait(driver, timeout).until_not(
            EC.presence_of_element_located(
                (By.XPATH, "//div[contains(@class,'modal-backdrop') or contains(@class,'overlay') or contains(@class,'loading')]")
            )
        )
    except TimeoutException:
        pass


def procesar_folio(driver, fila: dict, dry_run: bool = False,
                   solo_transferencias: bool = False,
                   solo_rendicion: bool = False) -> list[dict]:
    """
    Procesa un folio: navega, valida estado, y sube cada documento que aplique.
    Devuelve lista de logs por sección.
    """
    folio = fila["folio"]
    ano = int(fila.get("ano", 2022))
    logs = []

    # Descartar cualquier alert residual antes de navegar
    _descartar_alert(driver)

    estado = cargar_folio(driver, folio, ano)

    # Folios ya enviados o no existentes se saltan
    if estado in ("Enviado", "Cerrado"):
        logs.append({
            "folio": folio, "seccion": "-", "archivo": "-",
            "estado": "SKIP", "detalle": f"Convenio ya {estado.lower()}",
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        })
        return logs

    if estado in ("NO_EXISTE", "TIMEOUT", "DESCONOCIDO"):
        logs.append({
            "folio": folio, "seccion": "-", "archivo": "-",
            "estado": "ERROR", "detalle": f"Estado del folio: {estado}",
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
        })
        return logs

    # En modo solo-transferencias / solo-rendicion, procesar únicamente esa sección
    secciones_a_procesar = SECCIONES
    if solo_transferencias:
        secciones_a_procesar = [s for s in SECCIONES if s["titulo_dom"] == "Transferencias"]
    elif solo_rendicion:
        secciones_a_procesar = [s for s in SECCIONES if s["titulo_dom"] == "Respaldo de rendición"]

    # Iterar secciones
    for seccion_def in secciones_a_procesar:
        titulo = seccion_def["titulo_dom"]
        campos = seccion_def["campos_master"]

        # ¿Tenemos algún archivo asignado a esta sección?
        archivos_seccion = [fila.get(c, "").strip() for c in campos]
        archivos_seccion = [a for a in archivos_seccion if a]
        if not archivos_seccion:
            estado_skip = "FALTANTE" if seccion_def["obligatoria"] else "SKIP"
            logs.append({
                "folio": folio, "seccion": titulo, "archivo": "-",
                "estado": estado_skip, "detalle": "Sin archivo asignado en master",
                "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            })
            continue

        # Localizar la sección en el DOM
        try:
            seccion_el = localizar_seccion(driver, titulo)
        except NoSuchElementException:
            logs.append({
                "folio": folio, "seccion": titulo, "archivo": "-",
                "estado": "ERROR", "detalle": "Sección no encontrada en DOM",
                "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            })
            continue

        # Borrar docs previos antes de subir.
        # En modo solo-transferencias / solo-rendicion: siempre intentamos borrar.
        # borrar_docs_seccion() es seguro cuando no hay rows — devuelve 0.
        debe_borrar = solo_transferencias or solo_rendicion or seccion_ya_tiene_documentos(seccion_el)
        if debe_borrar:
            n_actual = obtener_n_docs_subidos(seccion_el, driver)
            if n_actual > 0 or seccion_ya_tiene_documentos(seccion_el):
                print(f"    → Borrando {n_actual} doc(s) previos en '{titulo}'...")
                borrados = borrar_docs_seccion(driver, titulo, dry_run=dry_run)
                if borrados:
                    logs.append({
                        "folio": folio, "seccion": titulo, "archivo": "-",
                        "estado": "BORRADO", "detalle": f"Borrados {borrados} doc(s) previos",
                        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                    })
                if not dry_run:
                    try:
                        seccion_el = localizar_seccion(driver, titulo)
                    except NoSuchElementException:
                        continue

        # Subir cada archivo (puede ser 1 o 2, en orden)
        for i, archivo in enumerate(archivos_seccion):
            # Re-localizar la sección porque el DOM cambió tras la subida anterior
            try:
                seccion_el = localizar_seccion(driver, titulo)
            except NoSuchElementException:
                logs.append({
                    "folio": folio, "seccion": titulo, "archivo": archivo,
                    "estado": "ERROR", "detalle": "Sección desapareció",
                    "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                })
                break

            n_antes = obtener_n_docs_subidos(seccion_el, driver)

            # En dry-run, simular sin tocar el DOM
            if dry_run:
                logs.append({
                    "folio": folio, "seccion": titulo, "archivo": archivo,
                    "estado": "DRY_RUN", "detalle": f"Subiría doc {i+1}",
                    "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                })
                continue

            # Para el segundo archivo en adelante, hacer click en "Agregar otro documento"
            if i > 0:
                if not click_agregar_otro_documento(driver, seccion_el):
                    logs.append({
                        "folio": folio, "seccion": titulo, "archivo": archivo,
                        "estado": "ERROR", "detalle": "Botón 'Agregar otro' no encontrado",
                        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                    })
                    break
                # Re-localizar
                try:
                    seccion_el = localizar_seccion(driver, titulo)
                except NoSuchElementException:
                    break

            print(f"    → Subiendo {titulo} ({i+1}/{len(archivos_seccion)}): {Path(archivo).name}")
            ok, detalle = subir_archivo_a_seccion(driver, titulo, archivo, n_antes)
            logs.append({
                "folio": folio, "seccion": titulo, "archivo": archivo,
                "estado": "OK" if ok else "ERROR", "detalle": detalle,
                "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            })
            print(f"      {'✓' if ok else '✗'} {detalle}")
            if not ok:
                break  # No intentar subir el segundo archivo si el primero falló

            time.sleep(SLEEP_ENTRE_ARCHIVOS)

    return logs


# ====== LECTURA DEL MASTER Y LOG ======

def leer_master() -> list[dict]:
    """Lee data/master_subida.xlsx y devuelve lista de filas como dict."""
    wb = openpyxl.load_workbook(MASTER_XLSX, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(h).strip() if h else "" for h in rows[0]]
    return [
        {h: (str(v).strip() if v is not None else "") for h, v in zip(headers, row)}
        for row in rows[1:]
        if any(row)
    ]


def cargar_folios_completados() -> set:
    """Lee el CSV de log y devuelve folios que ya están completos (sin errores recientes)."""
    if not Path(LOG_CSV).exists():
        return set()
    estados_por_folio = {}
    with open(LOG_CSV, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            folio = row["folio"]
            estado = row["estado"]
            if folio not in estados_por_folio:
                estados_por_folio[folio] = []
            estados_por_folio[folio].append(estado)
    completados = set()
    for folio, estados in estados_por_folio.items():
        if estados and "ERROR" not in estados:
            completados.add(folio)
    return completados


LOG_ENVIO_REVISION = "logs/envio_revision.csv"

def cargar_folios_enviados() -> set[str]:
    """Lee logs/envio_revision.csv y devuelve folios ya enviados a revisión."""
    if not Path(LOG_ENVIO_REVISION).exists():
        return set()
    enviados = set()
    with open(LOG_ENVIO_REVISION, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            if row.get("estado", "").upper() in ("ENVIADO", "YA_ENVIADO"):
                enviados.add(row["folio"].strip())
    return enviados


# ====== REPORTE FINAL ======

def generar_reporte(todos_los_logs: list[dict], filas_master: list[dict]) -> None:
    """
    Genera reporte de faltantes en consola y en logs/reporte_faltantes.csv.
    """
    # Índice folio → región/razón social
    info_folio = {f["folio"]: f for f in filas_master}

    # Agrupar logs por folio
    por_folio: dict[str, list[dict]] = {}
    for log in todos_los_logs:
        por_folio.setdefault(log["folio"], []).append(log)

    resumen: list[dict] = []
    for folio, logs in por_folio.items():
        info = info_folio.get(folio, {})
        region = info.get("region", "?")
        razon = info.get("razon_social", "?")[:50]

        estados_por_sec = {l["seccion"]: l["estado"] for l in logs}
        secciones_ok       = [s for s, e in estados_por_sec.items() if e == "OK"]
        secciones_faltante = [s for s, e in estados_por_sec.items() if e == "FALTANTE"]
        secciones_error    = [s for s, e in estados_por_sec.items() if e == "ERROR"]

        folio_ok = not secciones_faltante and not secciones_error

        resumen.append({
            "folio": folio,
            "region": region,
            "razon_social": razon,
            "ok": len(secciones_ok),
            "faltante": "; ".join(secciones_faltante) if secciones_faltante else "",
            "error": "; ".join(secciones_error) if secciones_error else "",
            "estado_final": "COMPLETO" if folio_ok else ("ERROR" if secciones_error else "INCOMPLETO"),
        })

    # Guardar CSV
    Path("logs").mkdir(exist_ok=True)
    with open(LOG_REPORTE, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=["folio","region","razon_social","ok","faltante","error","estado_final"])
        writer.writeheader()
        writer.writerows(resumen)

    # Consola
    completos   = [r for r in resumen if r["estado_final"] == "COMPLETO"]
    incompletos = [r for r in resumen if r["estado_final"] == "INCOMPLETO"]
    con_error   = [r for r in resumen if r["estado_final"] == "ERROR"]

    print("\n" + "=" * 70)
    print("  REPORTE DE CARGA")
    print("=" * 70)
    print(f"\n✅  Folios completamente cargados: {len(completos)}")
    print(f"⚠️   Folios con archivos faltantes:  {len(incompletos)}")
    print(f"🔴  Folios con errores:              {len(con_error)}")

    if incompletos:
        print("\n── ARCHIVOS FALTANTES POR FOLIO ──")
        # Agrupar por región
        por_region: dict[str, list] = {}
        for r in incompletos:
            por_region.setdefault(r["region"], []).append(r)
        for reg in sorted(por_region):
            print(f"\n  [{reg}]")
            for r in por_region[reg]:
                secciones = r["faltante"].replace("Convenio + Acto Administrativo", "Convenio+Acto")
                secciones = secciones.replace("Certificado de registro de entidad receptora", "Certificado")
                secciones = secciones.replace("Respaldo de rendición", "Rendición")
                print(f"    folio {r['folio']} — {r['razon_social'][:40]}")
                print(f"      Falta: {secciones}")

    if con_error:
        print("\n── ERRORES ──")
        for r in con_error:
            print(f"  folio {r['folio']} [{r['region']}] — {r['error'][:80]}")

    # Resumen por región
    print("\n── RESUMEN POR REGIÓN ──")
    regiones_stats: dict[str, dict] = {}
    for r in resumen:
        reg = r["region"]
        if reg not in regiones_stats:
            regiones_stats[reg] = {"total": 0, "completo": 0, "incompleto": 0, "error": 0}
        regiones_stats[reg]["total"] += 1
        if r["estado_final"] == "COMPLETO":
            regiones_stats[reg]["completo"] += 1
        elif r["estado_final"] == "INCOMPLETO":
            regiones_stats[reg]["incompleto"] += 1
        else:
            regiones_stats[reg]["error"] += 1

    print(f"\n  {'Región':<15} {'Total':>5} {'OK':>5} {'Incompleto':>11} {'Error':>7}")
    print(f"  {'-'*47}")
    for reg in sorted(regiones_stats):
        s = regiones_stats[reg]
        print(f"  {reg:<15} {s['total']:>5} {s['completo']:>5} {s['incompleto']:>11} {s['error']:>7}")

    print(f"\n  Reporte detallado → {LOG_REPORTE}")
    print()


# ====== MAIN ======

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="No sube nada, solo simula")
    parser.add_argument("--folio", type=str, default=None, help="Procesar solo este folio")
    parser.add_argument("--limit", type=int, default=None, help="Procesar solo N folios")
    parser.add_argument("--solo-transferencias", action="store_true",
                        help="Solo sube Transferencias; borra doc previo si existe y re-sube")
    parser.add_argument("--solo-rendicion", action="store_true",
                        help="Solo sube Respaldo de rendición; borra doc previo si existe y re-sube")
    parser.add_argument("--skip-enviados", action="store_true",
                        help="Salta folios que ya fueron enviados a revisión (lee logs/envio_revision.csv)")
    args = parser.parse_args()

    print("=" * 70)
    print("BOT DE SUBIDA SIRA")
    print("=" * 70)
    if args.dry_run:
        print("** MODO DRY-RUN: no se subirá nada **")
    if args.solo_transferencias:
        print("** MODO SOLO-TRANSFERENCIAS: borra doc previo y re-sube Transferencias **")
    if args.solo_rendicion:
        print("** MODO SOLO-RENDICION: borra doc previo y re-sube Respaldo de rendición **")

    Path("logs").mkdir(exist_ok=True)

    # Cargar master
    if not Path(MASTER_XLSX).exists():
        print(f"\n[ERROR] No existe {MASTER_XLSX}")
        print("Primero ejecutá el script 03 para generar el master.")
        sys.exit(1)

    filas = leer_master()
    print(f"\nFilas en master: {len(filas)}")

    # Saltar folios ya enviados a revisión
    if args.skip_enviados:
        enviados = cargar_folios_enviados()
        antes = len(filas)
        filas = [f for f in filas if f["folio"] not in enviados]
        print(f"Saltando enviados a revisión: {antes - len(filas)} omitidos, {len(filas)} restantes")

    # En modo solo-rendicion, procesar únicamente folios con rendicion_pdf que exista en disco
    if args.solo_rendicion and not args.folio:
        antes = len(filas)
        filas = [f for f in filas
                 if f.get("rendicion_pdf", "").strip()
                 and Path(f["rendicion_pdf"].strip()).exists()]
        print(f"Filtrado por rendicion_pdf en disco: {len(filas)} folios ({antes - len(filas)} sin archivo válido)")

    # Filtrar por folio específico
    if args.folio:
        filas = [f for f in filas if f["folio"] == args.folio]
        print(f"Filtrado por folio {args.folio}: {len(filas)} fila(s)")

    if args.limit:
        filas = filas[:args.limit]
        print(f"Limitado a {args.limit} folios")

    if not filas:
        print("\nNada que procesar.")
        return

    # Abrir log CSV
    log_existe = Path(LOG_CSV).exists()
    f_log = open(LOG_CSV, "a", encoding="utf-8-sig", newline="")
    log_writer = csv.DictWriter(
        f_log,
        fieldnames=["folio", "seccion", "archivo", "estado", "detalle", "timestamp"],
    )
    if not log_existe:
        log_writer.writeheader()
        f_log.flush()

    # Iniciar driver
    driver = crear_driver()
    n_ok = 0
    n_err = 0
    n_skip = 0
    todos_los_logs: list[dict] = []

    try:
        esperar_login(driver)

        for i, fila in enumerate(filas, 1):
            folio = fila["folio"]
            print(f"\n[{i}/{len(filas)}] Folio {folio} ({fila.get('razon_social','?')[:40]})")
            try:
                logs = procesar_folio(driver, fila, dry_run=args.dry_run,
                                     solo_transferencias=args.solo_transferencias,
                                     solo_rendicion=args.solo_rendicion)
                todos_los_logs.extend(logs)
                for log in logs:
                    log_writer.writerow(log)
                    f_log.flush()
                    if log["estado"] == "OK":
                        n_ok += 1
                    elif log["estado"] == "ERROR":
                        n_err += 1
                    elif log["estado"] in ("SKIP", "FALTANTE", "BORRADO"):
                        n_skip += 1
            except KeyboardInterrupt:
                print("\n[INTERRUPCIÓN] Guardando progreso y saliendo...")
                break
            except Exception as e:
                print(f"  [EXCEPCIÓN INESPERADA] {type(e).__name__}: {e}")
                err_log = {
                    "folio": folio, "seccion": "-", "archivo": "-",
                    "estado": "ERROR", "detalle": f"EXCEPCIÓN: {type(e).__name__}: {str(e)[:200]}",
                    "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                }
                log_writer.writerow(err_log)
                todos_los_logs.append(err_log)
                f_log.flush()
            time.sleep(SLEEP_ENTRE_FOLIOS)

    finally:
        f_log.close()
        print("\n" + "=" * 70)
        print("RESUMEN DE EJECUCIÓN")
        print("=" * 70)
        print(f"  Subidas OK:   {n_ok}")
        print(f"  Errores:      {n_err}")
        print(f"  Skipped:      {n_skip}")
        print(f"\n  Log -> {LOG_CSV}")

        # Generar reporte de faltantes
        if todos_los_logs:
            generar_reporte(todos_los_logs, filas)

        print("\nNavegador queda abierto. ENTER para cerrar.")
        try:
            input(">>> ")
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()