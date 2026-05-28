"""
Script 00: Bot de registro de convenios a mano en SIRA
=======================================================
Lee Libro1.xlsx (raíz del proyecto) y registra cada convenio en SIRA
usando el modal "Registrar convenio a mano".

Campos que completa (todos desde Libro1.xlsx):
  Obligatorios:
    - ID del convenio           ← columna "ID Convenio"
    - Nombre entidad receptora  ← columna "NOMBRE OOSS/MEDIO"
    - Año del convenio          ← columna "AÑO"
  Opcionales:
    - RUT entidad receptora     ← columna "RUT"
    - Tipo de convenio          ← columna "ALCANCE"
    - Nombre de la transferencia← columna "NOMBRE TRANSFERENCIA"
    - Monto total (CLP)         ← columna "MONTO"
    - Cantidad de cuotas        ← columna "CUOTAS"
    - Fecha de inicio           ← columna "INICIO"
    - Fecha de término          ← columna "TERMIO"
    - Objeto / propósito        ← columna "OBJETIVO"

Reglas:
  - Idempotente: lee logs/registro_convenios.csv y salta folios ya OK.
  - Verifica en SIRA si el folio ya existe antes de intentar registrar.
  - Reanudable: si lo cortás, retoma desde donde quedó.

Uso:
    python scripts/00_registrar_convenios.py
    python scripts/00_registrar_convenios.py --dry-run
    python scripts/00_registrar_convenios.py --folio 73542
"""

import argparse
import csv
import sys
import time
from datetime import datetime

from selenium.webdriver.common.keys import Keys
from pathlib import Path

import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
)
from webdriver_manager.chrome import ChromeDriverManager

# ====== CONFIGURACIÓN ======
LIBRO1_XLSX = "Libro1.xlsx"
LOG_CSV = "logs/registro_convenios.csv"
CHROME_PROFILE_DIR = "/home/bgcorrea/.bot_sira_chrome_profile"

BASE = "https://sira.auditoriainternadegobierno.gob.cl"
TIMEOUT_PAGINA = 25
TIMEOUT_MODAL = 15
SLEEP_ENTRE_REGISTROS = 2.5


# ====== LECTURA DE DATOS ======

def _normalizar_fecha(valor) -> str:
    """Convierte cualquier formato de fecha a mm/dd/yyyy que acepta el input de SIRA."""
    if not valor or str(valor).strip() in ("", "None"):
        return ""
    if isinstance(valor, datetime):
        return valor.strftime("%m/%d/%Y")
    s = str(valor).strip()
    # Formato "m/d/yyyy" o "mm/dd/yyyy"
    partes = s.split("/")
    if len(partes) == 3:
        try:
            m, d, y = partes
            return f"{int(m):02d}/{int(d):02d}/{y.split()[0]}"
        except ValueError:
            pass
    # Formato ISO "yyyy-mm-dd ..."
    try:
        dt = datetime.strptime(s[:10], "%Y-%m-%d")
        return dt.strftime("%m/%d/%Y")
    except ValueError:
        pass
    return s


def leer_libro1() -> list[dict]:
    """Lee Libro1.xlsx de la raíz y devuelve lista de dicts normalizados."""
    if not Path(LIBRO1_XLSX).exists():
        print(f"[ERROR] No se encontró {LIBRO1_XLSX}")
        sys.exit(1)
    wb = openpyxl.load_workbook(LIBRO1_XLSX, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    raw_headers = rows[0]
    headers = [str(h).strip() if h else f"col_{i}" for i, h in enumerate(raw_headers)]

    result = []
    for row in rows[1:]:
        if not any(v is not None for v in row):
            continue
        fila = {}
        for h, v in zip(headers, row):
            fila[h] = str(v).strip() if v is not None else ""
        # Guardar fechas normalizadas en claves adicionales
        idx_inicio = headers.index("INICIO") if "INICIO" in headers else -1
        idx_termio = headers.index("TERMIO") if "TERMIO" in headers else -1
        raw_inicio = row[idx_inicio] if idx_inicio >= 0 else None
        raw_termio = row[idx_termio] if idx_termio >= 0 else None
        fila["_fecha_inicio"] = _normalizar_fecha(raw_inicio)
        fila["_fecha_termio"] = _normalizar_fecha(raw_termio)
        result.append(fila)
    return result


# ====== LOG CSV ======

def cargar_registros_completados() -> set[str]:
    """Lee el log CSV y devuelve folios que ya fueron registrados con éxito."""
    if not Path(LOG_CSV).exists():
        return set()
    completados = set()
    with open(LOG_CSV, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            if row.get("estado", "").upper() in ("OK", "YA_EXISTE"):
                completados.add(row["folio"].strip())
    return completados


# ====== DRIVER ======

def crear_driver() -> webdriver.Chrome:
    Path(CHROME_PROFILE_DIR).mkdir(parents=True, exist_ok=True)
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
    print("Si el perfil Chrome ya tiene sesión, deberías estar logueado.")
    print("Si no, hacé login con Clave Única.")
    print("Cuando veas 'Hola, ...' en SIRA, presioná ENTER.")
    print("=" * 70)
    driver.get(BASE + "/")
    input("\n>>> Presioná ENTER cuando estés logueado: ")


# ====== INTERACCIÓN CON SIRA ======

def _descartar_alert(driver) -> str | None:
    try:
        alert = driver.switch_to.alert
        texto = alert.text
        alert.accept()
        time.sleep(0.3)
        return texto
    except Exception:
        return None


def navegar_a_inicio(driver):
    """Navega a la página principal y espera que cargue."""
    driver.get(BASE + "/")
    try:
        WebDriverWait(driver, TIMEOUT_PAGINA).until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[contains(text(), 'Registrar convenio a mano') or contains(text(), 'Hola,')]")
            )
        )
        time.sleep(0.8)
    except TimeoutException:
        pass


def folio_ya_existe_en_sira(driver, folio_id: str) -> bool:
    """
    Busca el folio en SIRA usando el parámetro ?q= y comprueba si aparece
    en el listado de convenios de la página principal.
    """
    driver.get(f"{BASE}/?q={folio_id}")
    try:
        # Esperar a que cargue la página
        WebDriverWait(driver, TIMEOUT_PAGINA).until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[contains(@class, 'convenio') or contains(text(), 'Registrar convenio')]")
            )
        )
        time.sleep(1)
    except TimeoutException:
        pass

    body = driver.find_element(By.TAG_NAME, "body").text
    # El folio aparece en tarjetas como "Convenio 73542" o simplemente el número
    # Buscamos coincidencia exacta del número en el body visible
    lineas = body.split("\n")
    for linea in lineas:
        # Una tarjeta de convenio tendría el ID exacto
        if folio_id in linea and "Convenio" in linea:
            return True
    return False


def abrir_modal_registro(driver) -> bool:
    """Hace click en '+ Registrar convenio a mano' y espera que el modal aparezca."""
    try:
        btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(text(), 'Registrar convenio a mano')]")
            )
        )
        btn.click()
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[contains(text(), 'ID del convenio')]")
            )
        )
        time.sleep(0.5)
        return True
    except (TimeoutException, NoSuchElementException):
        return False


def _escribir_js(driver, elemento, valor: str) -> None:
    """Escribe en un input/textarea de Angular usando JavaScript nativo.
    Dispara los eventos que Angular necesita para registrar el cambio."""
    driver.execute_script("""
        var el = arguments[0];
        var val = arguments[1];
        var tag = el.tagName.toLowerCase();
        var proto = tag === 'textarea'
            ? window.HTMLTextAreaElement.prototype
            : window.HTMLInputElement.prototype;
        var setter = Object.getOwnPropertyDescriptor(proto, 'value').set;
        setter.call(el, val);
        el.dispatchEvent(new Event('input',  {bubbles: true}));
        el.dispatchEvent(new Event('change', {bubbles: true}));
    """, elemento, valor)


def _encontrar_elemento(driver, placeholder_frag: str = None, label_texto: str = None,
                        tipo: str = "input"):
    """
    Devuelve el WebElement buscando primero por placeholder (CSS),
    luego por atributo 'for' del label, luego por XPath adyacente.
    tipo puede ser 'input', 'textarea' o 'date'.
    """
    # 1. Por placeholder (CSS) — más estable
    if placeholder_frag:
        selectores = [
            f"input[placeholder*='{placeholder_frag}']",
            f"textarea[placeholder*='{placeholder_frag}']",
        ]
        for css in selectores:
            try:
                return driver.find_element(By.CSS_SELECTOR, css)
            except NoSuchElementException:
                pass

    if not label_texto:
        return None

    # 2. Por atributo for del label → id del input
    try:
        lbl = driver.find_element(By.XPATH, f"//label[contains(.,'{label_texto}')]")
        for_id = lbl.get_attribute("for")
        if for_id:
            return driver.find_element(By.ID, for_id)
    except (NoSuchElementException, StaleElementReferenceException):
        pass

    # 3. Por XPaths adyacentes al label
    if tipo == "date":
        xpaths = [
            f"//label[contains(.,'{label_texto}')]/following-sibling::input[@type='date'][1]",
            f"//label[contains(.,'{label_texto}')]/..//input[@type='date'][1]",
            f"//label[contains(.,'{label_texto}')]/following::input[@type='date'][1]",
        ]
    elif tipo == "textarea":
        xpaths = [
            f"//label[contains(.,'{label_texto}')]/following-sibling::textarea[1]",
            f"//label[contains(.,'{label_texto}')]/..//textarea[1]",
            f"//label[contains(.,'{label_texto}')]/following::textarea[1]",
        ]
    else:
        xpaths = [
            f"//label[contains(.,'{label_texto}')]/following-sibling::input[not(@type='hidden') and not(@type='date')][1]",
            f"//label[contains(.,'{label_texto}')]/..//input[not(@type='hidden') and not(@type='date')][1]",
            f"//label[contains(.,'{label_texto}')]/following::input[not(@type='hidden') and not(@type='date')][1]",
        ]
    for xpath in xpaths:
        try:
            return driver.find_element(By.XPATH, xpath)
        except (NoSuchElementException, StaleElementReferenceException):
            continue
    return None


def _llenar_campo(driver, valor: str, placeholder_frag: str = None,
                  label_texto: str = None) -> bool:
    """Llena un input de texto usando JavaScript (compatible con Angular)."""
    if not valor:
        return True
    el = _encontrar_elemento(driver, placeholder_frag=placeholder_frag,
                             label_texto=label_texto, tipo="input")
    if el is None:
        return False
    _escribir_js(driver, el, valor)
    return True


def _obtener_modal(driver):
    """Devuelve el elemento raíz del modal para acotar búsquedas a él."""
    xpaths = [
        "//div[contains(@class,'modal-dialog')]",
        "//div[contains(@class,'modal-content')]",
        "//*[contains(text(),'Registrar convenio a mano')]/ancestor::div[contains(@class,'modal')][1]",
        "//*[contains(text(),'ID del convenio')]/ancestor::div[4]",
    ]
    for xpath in xpaths:
        try:
            el = driver.find_element(By.XPATH, xpath)
            if el:
                return el
        except (NoSuchElementException, StaleElementReferenceException):
            continue
    return driver  # fallback: toda la página


def _llenar_opcionales_posicional(driver, rut, alcance, nombre_trans, monto, cuotas,
                                   fecha_inicio, fecha_termio) -> None:
    """
    Llena los campos opcionales por posición dentro del modal.

    Acotando al modal se evita capturar el buscador de la página principal
    que aparece como input extra y desplaza todos los índices.

    Inputs de texto dentro del modal (tras expansión):
      [0]=ID  [1]=Nombre  [2]=RUT  [3]=Tipo  [4]=Nombre_trans  [5]=Monto  [6]=Cuotas
    Inputs de fecha:
      [0]=Fecha de inicio  [1]=Fecha de término
    """
    modal = _obtener_modal(driver)

    # ── Campos de texto por índice (scoped al modal) ─────────────────────────
    # Orden dentro del modal tras expansión:
    #   [0]=ID  [1]=Nombre  [2]=RUT  [3]=Tipo de convenio  [4]=Nombre transf.
    #   [5]=Monto  [6]=Cuotas
    # "Tipo de convenio" (idx=3) recibe el texto largo hardcodeado.
    # "Nombre de la transferencia" (idx=4) se deja vacío.
    text_inputs = modal.find_elements(
        By.CSS_SELECTOR, "input:not([type='hidden']):not([type='date'])"
    )
    valores_texto = [rut, nombre_trans, "", monto, cuotas]
    for offset, valor in enumerate(valores_texto):
        idx = 2 + offset
        if valor and idx < len(text_inputs):
            _escribir_js(driver, text_inputs[idx], valor)

    # ── Fechas por índice (scoped al modal) ──────────────────────────────────
    # Se usa send_keys en vez de JS puro: Angular no registra el valor en el
    # form model si solo se modifica el DOM, pero sí lo hace con eventos reales
    # de teclado. Se re-busca la lista antes de cada escritura para evitar stale.
    for i, fecha_str in enumerate([fecha_inicio, fecha_termio]):
        if not fecha_str:
            continue
        date_inputs = modal.find_elements(By.CSS_SELECTOR, "input[type='date']")
        if i >= len(date_inputs):
            continue
        el = date_inputs[i]
        try:
            # Formato "07/17/2023" → "07172023" para el cursor del date picker de Chrome
            partes = fecha_str.split("/")
            valor_teclado = "".join(partes)  # "07172023"
            el.click()
            el.send_keys(Keys.CONTROL + "a")
            el.send_keys(Keys.DELETE)
            el.send_keys(valor_teclado)
            el.send_keys(Keys.TAB)
            time.sleep(0.4)
        except Exception:
            pass


def _seleccionar_anio(driver, anio: str) -> bool:
    """Selecciona el año en el dropdown del modal."""
    xpaths_select = [
        "//label[contains(.,'Año del convenio')]/following-sibling::select[1]",
        "//label[contains(.,'Año del convenio')]/..//select[1]",
        "//select[contains(@ng-model,'anio') or contains(@ng-model,'año') or contains(@name,'anio')]",
    ]
    for xpath in xpaths_select:
        try:
            sel_el = driver.find_element(By.XPATH, xpath)
            Select(sel_el).select_by_visible_text(anio)
            return True
        except (NoSuchElementException, StaleElementReferenceException):
            continue
    # Último recurso: seleccionar por valor
    try:
        sel_el = driver.find_element(By.XPATH, "//select")
        Select(sel_el).select_by_value(anio)
        return True
    except Exception:
        return False


def _expandir_campos_opcionales(driver) -> bool:
    """Hace click en 'Más campos opcionales' para mostrar el formulario extendido."""
    xpaths = [
        "//*[contains(text(),'Más campos opcionales')]",
        "//*[contains(text(),'campos opcionales')]",
        "//*[contains(@class,'expandir') or contains(@class,'toggle') or contains(@class,'collapse')]",
    ]
    for xpath in xpaths:
        try:
            el = driver.find_element(By.XPATH, xpath)
            driver.execute_script("arguments[0].click();", el)
            time.sleep(0.6)
            return True
        except NoSuchElementException:
            continue
    return False


def _confirmar_registro(driver) -> bool:
    """Hace click en el botón 'Registrar' del modal y espera confirmación."""
    try:
        btn = WebDriverWait(driver, 8).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[normalize-space(text())='Registrar' or (contains(text(),'Registrar') and not(contains(text(),'mano')))]")
            )
        )
        btn.click()
        # Esperar que el modal desaparezca
        WebDriverWait(driver, TIMEOUT_MODAL).until_not(
            EC.presence_of_element_located(
                (By.XPATH, "//*[contains(text(),'ID del convenio')]")
            )
        )
        time.sleep(1)
        return True
    except TimeoutException:
        return False
    except NoSuchElementException:
        return False


def _cerrar_modal_si_abierto(driver):
    """Cierra el modal si quedó abierto (por error), presionando Cancelar."""
    try:
        btn_cancelar = driver.find_element(
            By.XPATH, "//button[contains(text(),'Cancelar')]"
        )
        btn_cancelar.click()
        time.sleep(0.5)
    except NoSuchElementException:
        pass


# ====== ORQUESTACIÓN POR FILA ======

def procesar_fila(driver, fila: dict, dry_run: bool = False, no_submit: bool = False) -> tuple[str, str]:
    """
    Registra un convenio. Devuelve (estado, detalle).
    estados: OK | YA_EXISTE | SKIP | ERROR_... | DRY_RUN
    """
    folio = fila.get("ID Convenio", "").strip()
    razon = fila.get("NOMBRE OOSS/MEDIO", "").strip()
    anio = fila.get("AÑO", "").strip()
    rut = fila.get("RUT", "").strip()
    alcance = fila.get("ALCANCE", "").strip()
    nombre_trans = "CONVENIO DE TRANSFERENCIA CORRIENTE A ENTIDAD PRIVADA BENEFICIARIA"
    monto = fila.get("MONTO", "").strip()
    cuotas = "1"
    fecha_inicio = fila.get("_fecha_inicio", "")
    fecha_termio = fila.get("_fecha_termio", "")

    if not folio or not razon or not anio:
        return "ERROR_DATOS", f"Campos obligatorios vacíos: folio={folio!r} razon={razon!r} anio={anio!r}"

    if dry_run:
        return "DRY_RUN", f"Registraría folio={folio}, entidad={razon[:50]}, año={anio}"

    _descartar_alert(driver)

    # Verificar si ya existe en SIRA
    if folio_ya_existe_en_sira(driver, folio):
        return "YA_EXISTE", "El convenio ya figura en SIRA"

    # Ir al inicio y abrir modal
    navegar_a_inicio(driver)
    if not abrir_modal_registro(driver):
        return "ERROR_MODAL", "No se pudo abrir el modal de registro"

    # Campos obligatorios — por placeholder (único e inequívoco)
    if not _llenar_campo(driver, folio, placeholder_frag="Como aparece",
                         label_texto="ID del convenio"):
        _cerrar_modal_si_abierto(driver)
        return "ERROR_CAMPO_ID", "No se encontró el campo ID del convenio"

    if not _llenar_campo(driver, razon, placeholder_frag="Fundación",
                         label_texto="Nombre de la entidad receptora"):
        _cerrar_modal_si_abierto(driver)
        return "ERROR_CAMPO_NOMBRE", "No se encontró el campo nombre entidad"

    if not _seleccionar_anio(driver, anio):
        _cerrar_modal_si_abierto(driver)
        return "ERROR_CAMPO_ANIO", f"No se pudo seleccionar el año {anio}"

    # Expandir campos opcionales y esperar que Angular los renderice
    _expandir_campos_opcionales(driver)
    time.sleep(1.2)

    # Campos opcionales — por posición en el DOM
    _llenar_opcionales_posicional(
        driver, rut, alcance, nombre_trans, monto, cuotas, fecha_inicio, fecha_termio
    )

    # Sin submit: dejar el modal abierto para revisión visual
    if no_submit:
        input(f"\n  [NO-SUBMIT] Folio {folio} — revisá el modal y presioná ENTER para continuar con el siguiente: ")
        _cerrar_modal_si_abierto(driver)
        return "NO_SUBMIT", "Modal llenado pero no enviado"

    # Confirmar
    if not _confirmar_registro(driver):
        # Revisar si hay error visible en el modal
        try:
            body = driver.find_element(By.TAG_NAME, "body").text
        except Exception:
            body = ""
        if "ya existe" in body.lower() or "duplicado" in body.lower():
            _cerrar_modal_si_abierto(driver)
            return "YA_EXISTE", "SIRA reportó que el convenio ya existe"
        _cerrar_modal_si_abierto(driver)
        return "ERROR_CONFIRMACION", "Timeout esperando que el modal se cierre"

    return "OK", "Convenio registrado exitosamente"


# ====== MAIN ======

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dry-run", action="store_true", help="Solo muestra lo que haría, no registra nada")
    parser.add_argument("--no-submit", action="store_true", help="Llena el modal pero no hace click en Registrar")
    parser.add_argument("--folio", type=str, default=None, help="Procesar solo este folio (para test)")
    parser.add_argument("--limit", type=int, default=None, help="Procesar solo N registros")
    args = parser.parse_args()

    print("=" * 70)
    print("BOT REGISTRO DE CONVENIOS - SIRA")
    print("=" * 70)
    if args.dry_run:
        print("** MODO DRY-RUN: no se registrará nada **")
    if args.no_submit:
        print("** MODO NO-SUBMIT: llena el formulario pero no hace click en Registrar **")
    print()

    Path("logs").mkdir(exist_ok=True)

    # Cargar datos
    filas = leer_libro1()
    print(f"Filas en Libro1: {len(filas)}")

    # Filtrar por folio si se especificó
    if args.folio:
        filas = [f for f in filas if f.get("ID Convenio", "").strip() == args.folio]
        print(f"Filtrado por folio {args.folio}: {len(filas)} fila(s)")

    if args.limit:
        filas = filas[:args.limit]
        print(f"Limitado a {args.limit} registros")

    # Saltar folios ya completados
    completados = cargar_registros_completados()
    if completados:
        antes = len(filas)
        filas = [f for f in filas if f.get("ID Convenio", "").strip() not in completados]
        print(f"Ya registrados (log): {antes - len(filas)} omitidos, {len(filas)} pendientes")

    if not filas:
        print("\nNada que procesar. Saliendo.")
        return

    print(f"\nA procesar: {len(filas)} convenio(s)\n")

    # Abrir log CSV
    log_existe = Path(LOG_CSV).exists()
    f_log = open(LOG_CSV, "a", encoding="utf-8-sig", newline="")
    log_writer = csv.DictWriter(
        f_log,
        fieldnames=["folio", "razon_social", "estado", "detalle", "timestamp"],
    )
    if not log_existe:
        log_writer.writeheader()
        f_log.flush()

    if args.dry_run:
        # En dry-run no abrimos Chrome
        n_ok = n_err = n_skip = 0
        for i, fila in enumerate(filas, 1):
            folio = fila.get("ID Convenio", "?")
            razon = fila.get("NOMBRE OOSS/MEDIO", "?")[:45]
            estado, detalle = procesar_fila(None, fila, dry_run=True)
            print(f"  [{i:>3}] Folio {folio:>6} — {razon}")
            print(f"         {estado}: {detalle}")
            log_writer.writerow({
                "folio": folio, "razon_social": razon,
                "estado": estado, "detalle": detalle,
                "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            })
            f_log.flush()
            n_ok += 1
        f_log.close()
        print(f"\n  Total dry-run: {n_ok} filas")
        print(f"  Log -> {LOG_CSV}")
        return

    # Iniciar Chrome
    driver = crear_driver()
    n_ok = n_err = n_ya_existe = 0

    try:
        esperar_login(driver)

        for i, fila in enumerate(filas, 1):
            folio = fila.get("ID Convenio", "?").strip()
            razon = fila.get("NOMBRE OOSS/MEDIO", "?")[:45]
            print(f"\n[{i:>3}/{len(filas)}] Folio {folio} — {razon}")

            try:
                estado, detalle = procesar_fila(driver, fila, dry_run=False, no_submit=args.no_submit)
            except KeyboardInterrupt:
                print("\n[INTERRUPCIÓN] Guardando progreso y saliendo...")
                break
            except Exception as e:
                estado = "ERROR_EXCEPCION"
                detalle = f"{type(e).__name__}: {str(e)[:200]}"

            icono = "✓" if estado in ("OK", "YA_EXISTE") else ("→" if estado == "SKIP" else "✗")
            print(f"  {icono} {estado}: {detalle}")

            log_writer.writerow({
                "folio": folio,
                "razon_social": fila.get("NOMBRE OOSS/MEDIO", ""),
                "estado": estado,
                "detalle": detalle,
                "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            })
            f_log.flush()

            if estado == "OK":
                n_ok += 1
            elif estado == "YA_EXISTE":
                n_ya_existe += 1
            elif estado not in ("SKIP", "DRY_RUN"):
                n_err += 1

            time.sleep(SLEEP_ENTRE_REGISTROS)

    finally:
        f_log.close()
        print("\n" + "=" * 70)
        print("RESUMEN")
        print("=" * 70)
        print(f"  Registrados OK:    {n_ok}")
        print(f"  Ya existían:       {n_ya_existe}")
        print(f"  Errores:           {n_err}")
        print(f"\n  Log -> {LOG_CSV}")
        print("\nNavegador queda abierto. ENTER para cerrar.")
        try:
            input(">>> ")
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
