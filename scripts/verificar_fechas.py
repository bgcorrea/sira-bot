"""
Script: verificar_fechas.py
===========================
Compara las fechas de inicio y término registradas en SIRA contra las
fechas esperadas en Libro1.xlsx, para detectar casos donde se dieron
vuelta día y mes (ej: 01/07 → 07/01).

Uso:
    python scripts/verificar_fechas.py
    python scripts/verificar_fechas.py --folio 73542
    python scripts/verificar_fechas.py --limit 10

Output:
    logs/verificacion_fechas.csv
"""

import argparse
import csv
import re
import time
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote

import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# ====== CONFIG ======
LIBRO1_XLSX = "Libro1.xlsx"
LOG_REGISTRO_CSV = "logs/registro_convenios.csv"
OUTPUT_CSV = "logs/verificacion_fechas.csv"
CHROME_PROFILE_DIR = "/home/bgcorrea/.bot_sira_chrome_profile"
BASE = "https://sira.auditoriainternadegobierno.gob.cl"
TIMEOUT = 25
SLEEP_ENTRE_FOLIOS = 1.5

SEGEGOB = quote("SECRETARÍA GENERAL DE GOBIERNO")
SUBSEG = quote("SUBSECRETARÍA GENERAL DE GOBIERNO")
URL_FOLIO = BASE + "/?q={folio}#/convenio/" + SEGEGOB + "/" + SUBSEG + "/_/{folio}"

MESES_ES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
}

CAMPOS_SALIDA = [
    "folio", "razon_social",
    "inicio_esperado", "termio_esperado",
    "inicio_sira", "termio_sira",
    "inicio_ok", "termio_ok",
    "resultado", "detalle", "timestamp",
]


# ====== PARSERS ======

def normalizar_fecha_libro(valor) -> str | None:
    """Devuelve 'dd/mm/yyyy' desde datetime o string 'm/d/yyyy'."""
    if valor is None:
        return None
    if isinstance(valor, (datetime, date)):
        return f"{valor.day:02d}/{valor.month:02d}/{valor.year}"
    s = str(valor).strip()
    partes = s.split("/")
    if len(partes) == 3:
        m, d, y = partes
        return f"{int(d):02d}/{int(m):02d}/{y}"
    return s


def parsear_fecha_sira(texto: str) -> str | None:
    """
    Intenta parsear la fecha que muestra SIRA. Formatos posibles:
      - '17 de julio de 2023'      → '17/07/2023'
      - '17/07/2023'               → '17/07/2023'
      - '2023-07-17'               → '17/07/2023'
    Devuelve 'dd/mm/yyyy' o None si no puede parsear.
    """
    texto = texto.strip()

    # 'dd de mes de yyyy'
    m = re.match(r"(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})", texto, re.IGNORECASE)
    if m:
        dia, mes_str, anio = m.groups()
        mes = MESES_ES.get(mes_str.lower())
        if mes:
            return f"{int(dia):02d}/{mes:02d}/{anio}"

    # 'dd/mm/yyyy'
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", texto)
    if m:
        d, mo, y = m.groups()
        return f"{int(d):02d}/{int(mo):02d}/{y}"

    # 'yyyy-mm-dd'
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", texto)
    if m:
        y, mo, d = m.groups()
        return f"{int(d):02d}/{int(mo):02d}/{y}"

    return None


def extraer_despues_de(texto: str, label: str) -> str:
    try:
        idx = texto.upper().index(label.upper())
        resto = texto[idx + len(label):].strip().split("\n")
        for linea in resto:
            linea = linea.strip()
            if linea:
                return linea
        return ""
    except ValueError:
        return ""


# ====== SELENIUM ======

def crear_driver():
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
    print("Si el perfil Chrome ya tenía sesión, estarás logueado automáticamente.")
    print("Si no, hacé login con Clave Única.")
    driver.get(BASE + "/")
    input("\n>>> Presioná ENTER cuando estés dentro de SIRA: ")
    print("Continuando...\n")


def detectar_sesion_expirada(body_text: str) -> bool:
    keywords = ["iniciar sesión", "iniciar sesion", "ingresar con clave", "claveunica.gob.cl"]
    return any(kw in body_text.lower() for kw in keywords)


def extraer_fechas_folio(driver, folio: str) -> dict:
    """Navega al folio en SIRA y extrae FECHA DE SUSCRIPCIÓN y FECHA DE TÉRMINO."""
    url = URL_FOLIO.format(folio=folio)
    driver.get(url)

    resultado = {
        "inicio_sira_raw": "",
        "termio_sira_raw": "",
        "error": "",
    }

    try:
        WebDriverWait(driver, TIMEOUT).until(
            EC.any_of(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(text(), 'FECHA DE SUSCRIPCIÓN')]")
                ),
                EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(text(), 'Datos declarados')]")
                ),
                EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(text(), 'Iniciar sesión')]")
                ),
                EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(text(), 'no existe')]")
                ),
            )
        )
        time.sleep(1.2)
        body = driver.find_element(By.TAG_NAME, "body").text

        if detectar_sesion_expirada(body):
            resultado["error"] = "SESION_EXPIRADA"
            return resultado

        if "no existe en el catálogo" in body.lower():
            resultado["error"] = "FOLIO_NO_EXISTE"
            return resultado

        resultado["inicio_sira_raw"] = extraer_despues_de(body, "FECHA DE SUSCRIPCIÓN")
        resultado["termio_sira_raw"] = extraer_despues_de(body, "FECHA DE TÉRMINO")

    except TimeoutException:
        resultado["error"] = f"TIMEOUT_{TIMEOUT}s"
    except Exception as e:
        resultado["error"] = f"{type(e).__name__}: {str(e)[:150]}"

    return resultado


# ====== CARGA DE DATOS ======

def cargar_folios_ok() -> list[str]:
    """Lee el CSV de registro y devuelve folios con estado OK (único, sin duplicados)."""
    vistos = set()
    folios = []
    with open(LOG_REGISTRO_CSV, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            if row["estado"] == "OK" and row["folio"] not in vistos:
                vistos.add(row["folio"])
                folios.append(row["folio"])
    return folios


def cargar_datos_libro() -> dict[str, dict]:
    """Devuelve dict folio→{razon_social, inicio, termio} desde Libro1.xlsx."""
    wb = openpyxl.load_workbook(LIBRO1_XLSX, data_only=True)
    ws = wb.active
    datos = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        folio = row[8]
        if folio is None:
            continue
        folio = str(folio).strip()
        datos[folio] = {
            "razon_social": row[11] or "",
            "inicio": normalizar_fecha_libro(row[16]),
            "termio": normalizar_fecha_libro(row[17]),
        }
    return datos


# ====== MAIN ======

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--folio", help="Verificar solo este folio")
    parser.add_argument("--limit", type=int, help="Máximo N folios a verificar")
    args = parser.parse_args()

    Path("logs").mkdir(exist_ok=True)

    # Cargar datos esperados
    libro = cargar_datos_libro()

    # Determinar folios a verificar
    if args.folio:
        folios = [args.folio]
    else:
        folios = cargar_folios_ok()
        if args.limit:
            folios = folios[: args.limit]

    print(f"\n{'=' * 70}")
    print(f"VERIFICACIÓN DE FECHAS — {len(folios)} folios")
    print(f"{'=' * 70}\n")

    # Abrir CSV de salida
    f_out = open(OUTPUT_CSV, "w", encoding="utf-8-sig", newline="")
    writer = csv.DictWriter(f_out, fieldnames=CAMPOS_SALIDA)
    writer.writeheader()

    driver = crear_driver()
    n_ok = n_mismatch = n_error = 0

    try:
        esperar_login(driver)

        for i, folio in enumerate(folios, 1):
            info = libro.get(folio, {})
            razon = info.get("razon_social", "?")
            inicio_esp = info.get("inicio")
            termio_esp = info.get("termio")

            print(f"[{i:>3}/{len(folios)}] {folio} — {razon[:45]}...", end=" ", flush=True)

            datos_sira = extraer_fechas_folio(driver, folio)

            if datos_sira["error"] == "SESION_EXPIRADA":
                print("⚠️  SESIÓN EXPIRADA")
                input("Hacé login y presioná ENTER: ")
                datos_sira = extraer_fechas_folio(driver, folio)

            if datos_sira["error"]:
                n_error += 1
                fila = {
                    "folio": folio,
                    "razon_social": razon,
                    "inicio_esperado": inicio_esp or "",
                    "termio_esperado": termio_esp or "",
                    "inicio_sira": "",
                    "termio_sira": "",
                    "inicio_ok": "",
                    "termio_ok": "",
                    "resultado": "ERROR",
                    "detalle": datos_sira["error"],
                    "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                }
                writer.writerow(fila)
                f_out.flush()
                print(f"❌ {datos_sira['error']}")
                time.sleep(SLEEP_ENTRE_FOLIOS)
                continue

            # Parsear fechas SIRA
            inicio_sira = parsear_fecha_sira(datos_sira["inicio_sira_raw"]) if datos_sira["inicio_sira_raw"] else ""
            termio_sira = parsear_fecha_sira(datos_sira["termio_sira_raw"]) if datos_sira["termio_sira_raw"] else ""

            # Comparar
            inicio_ok = (inicio_sira == inicio_esp) if (inicio_sira and inicio_esp) else None
            termio_ok = (termio_sira == termio_esp) if (termio_sira and termio_esp) else None

            hay_mismatch = (inicio_ok is False) or (termio_ok is False)
            resultado = "MISMATCH" if hay_mismatch else "COINCIDE"

            if hay_mismatch:
                n_mismatch += 1
                marca = "⚠️ "
            else:
                n_ok += 1
                marca = "✓ "

            detalle_partes = []
            if inicio_ok is False:
                detalle_partes.append(f"inicio: esperado={inicio_esp} sira={inicio_sira}")
            if termio_ok is False:
                detalle_partes.append(f"termio: esperado={termio_esp} sira={termio_sira}")
            detalle = " | ".join(detalle_partes) if detalle_partes else ""

            fila = {
                "folio": folio,
                "razon_social": razon,
                "inicio_esperado": inicio_esp or "",
                "termio_esperado": termio_esp or "",
                "inicio_sira": inicio_sira or datos_sira["inicio_sira_raw"],
                "termio_sira": termio_sira or datos_sira["termio_sira_raw"],
                "inicio_ok": str(inicio_ok),
                "termio_ok": str(termio_ok),
                "resultado": resultado,
                "detalle": detalle,
                "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            }
            writer.writerow(fila)
            f_out.flush()

            print(f"{marca} inicio={inicio_sira or '?'} term={termio_sira or '?'} → {resultado}")
            if detalle:
                print(f"       {detalle}")

            time.sleep(SLEEP_ENTRE_FOLIOS)

    finally:
        f_out.close()
        print(f"\n{'=' * 70}")
        print("RESUMEN")
        print(f"{'=' * 70}")
        print(f"  Coinciden:  {n_ok}")
        print(f"  Mismatch:   {n_mismatch}")
        print(f"  Errores:    {n_error}")
        print(f"\n  Resultado → {OUTPUT_CSV}")
        input("\nENTER para cerrar Chrome: ")
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
