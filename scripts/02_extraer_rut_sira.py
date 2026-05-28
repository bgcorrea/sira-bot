"""
Script 02: Extraer datos de los 250 folios desde SIRA
======================================================
Login manual con Clave Única, después recorre los 250 folios del Excel
extrayendo RUT, razón social, fecha de suscripción, garantía, etc.

Features:
- Idempotente: si lo cortás, retoma desde donde quedó (lee CSV existente).
- Escribe cada fila inmediatamente (no pierde progreso si crashea).
- Detecta sesión expirada y te pide relogin sin perder lo procesado.
- Genera DOS outputs:
    logs/rut_por_folio.csv      -> CSV "limpio" para usar después
    logs/sira_raw_dump.jsonl    -> JSON Lines con todo el body crudo (debug)

Uso:
    python scripts\02_extraer_rut_sira.py
"""

import time
import csv
import json
from pathlib import Path
from urllib.parse import quote
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# ====== CONFIGURACIÓN ======
EXCEL_INPUT = "data/Libro1.xlsx"
CSV_OUTPUT = "logs/rut_por_folio.csv"
JSONL_OUTPUT = "logs/sira_raw_dump.jsonl"
CHROME_PROFILE_DIR = "/home/bgcorrea/.bot_sira_chrome_profile"
TIMEOUT = 20
SLEEP_ENTRE_FOLIOS = 1.2  # Segundos entre folios para no saturar

# URLs según año
SEGEGOB = quote("SECRETARÍA GENERAL DE GOBIERNO")
SUBSEG = quote("SUBSECRETARÍA GENERAL DE GOBIERNO")
BASE = "https://sira.auditoriainternadegobierno.gob.cl"
URL_2022 = BASE + "/?q={folio}#/convenio/" + SEGEGOB + "/" + SUBSEG + "/" + SUBSEG + "/{folio}"
URL_2024 = BASE + "/?q={folio}#/convenio/" + SEGEGOB + "/" + SUBSEG + "/_/{folio}"

# Campos del CSV de salida (en orden)
CAMPOS_CSV = [
    "folio", "ano", "fondo", "region_excel",
    "estado_sira", "razon_social", "nombre_fantasia",
    "rut", "tipo_entidad",
    "fecha_suscripcion", "modalidad", "tipo_convenio",
    "alcance", "cuotas_declaradas",
    "tipo_garantia", "monto_garantia",
    "toma_razon", "estados_financieros", "socios_beneficiarios",
    "url", "error", "timestamp",
]


def url_para_folio(folio: str, ano: int) -> str:
    if ano <= 2022:
        return URL_2022.format(folio=folio)
    return URL_2024.format(folio=folio)


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
    print("Si el perfil de Chrome ya tenía sesión activa, deberías estar")
    print("logueado. Si no, hacé login con Clave Única.")
    print("Cuando estés DENTRO de SIRA, presioná ENTER.")
    print("=" * 70)
    driver.get(BASE + "/")
    input("\n>>> Presioná ENTER cuando estés logueado: ")
    print("Continuando...\n")


def extraer_despues_de(texto: str, label: str) -> str:
    """Busca label y devuelve el siguiente valor no vacío."""
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


def detectar_sesion_expirada(body_text: str) -> bool:
    """Heurística para detectar que la sesión Clave Única expiró."""
    keywords = ["iniciar sesión", "iniciar sesion", "ingresar con clave", "claveunica.gob.cl"]
    return any(kw in body_text.lower() for kw in keywords)


def extraer_datos_folio(driver, folio: str, ano: int) -> dict:
    """Navega al folio y devuelve dict con todos los datos."""
    url = url_para_folio(folio, ano)
    driver.get(url)

    datos = {
        "folio": folio,
        "ano": ano,
        "url": url,
        "estado_sira": "",
        "razon_social": "",
        "nombre_fantasia": "",
        "rut": "",
        "tipo_entidad": "",
        "fecha_suscripcion": "",
        "modalidad": "",
        "tipo_convenio": "",
        "alcance": "",
        "cuotas_declaradas": "",
        "tipo_garantia": "",
        "monto_garantia": "",
        "toma_razon": "",
        "estados_financieros": "",
        "socios_beneficiarios": "",
        "error": "",
        "body_raw": "",  # solo para JSONL, no va al CSV
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
    }

    try:
        WebDriverWait(driver, TIMEOUT).until(
            EC.any_of(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Datos declarados')]")),
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'no existe')]")),
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'RUT ENTIDAD')]")),
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Iniciar sesión')]")),
            )
        )
        time.sleep(1.5)
        body_text = driver.find_element(By.TAG_NAME, "body").text
        datos["body_raw"] = body_text

        if detectar_sesion_expirada(body_text):
            datos["error"] = "SESION_EXPIRADA"
            return datos

        if "no existe en el catálogo" in body_text.lower():
            datos["error"] = "FOLIO_NO_EXISTE"
            return datos

        # Parsear campos
        datos["rut"] = extraer_despues_de(body_text, "RUT ENTIDAD RECEPTORA")
        datos["tipo_entidad"] = extraer_despues_de(body_text, "TIPO DE ENTIDAD RECEPTORA")
        datos["fecha_suscripcion"] = extraer_despues_de(body_text, "FECHA DE SUSCRIPCIÓN")
        datos["modalidad"] = extraer_despues_de(body_text, "MODALIDAD DE ASIGNACIÓN")
        datos["tipo_convenio"] = extraer_despues_de(body_text, "TIPO DE CONVENIO")
        datos["alcance"] = extraer_despues_de(body_text, "ALCANCE")
        datos["cuotas_declaradas"] = extraer_despues_de(body_text, "CUOTAS DECLARADAS")
        datos["tipo_garantia"] = extraer_despues_de(body_text, "TIPO DE GARANTÍA")
        datos["monto_garantia"] = extraer_despues_de(body_text, "MONTO GARANTÍA")
        datos["toma_razon"] = extraer_despues_de(body_text, "TOMA DE RAZÓN CGR")
        datos["estados_financieros"] = extraer_despues_de(body_text, "ESTADOS FINANCIEROS")
        datos["socios_beneficiarios"] = extraer_despues_de(body_text, "SOCIOS / BENEFICIARIOS FINALES")

        # Razón social: línea con " / " que no contenga el folio
        for linea in body_text.split("\n"):
            linea = linea.strip()
            if " / " in linea and len(linea) > 10 and folio not in linea:
                partes = [p.strip() for p in linea.split(" / ", 1)]
                if len(partes) == 2:
                    # Descartar "Socios: SI · Beneficiarios: SI" que también tiene " / "
                    if "Socios" not in partes[0] and "SUBSECRETARÍA" not in partes[0]:
                        datos["razon_social"] = partes[0]
                        datos["nombre_fantasia"] = partes[1]
                        break

        # Estado
        for kw in ["Borrador", "Cerrado", "Enviado", "Vigente"]:
            if kw in body_text[:1000]:
                datos["estado_sira"] = kw
                break

    except TimeoutException:
        datos["error"] = f"TIMEOUT_{TIMEOUT}s"
    except Exception as e:
        datos["error"] = f"{type(e).__name__}: {str(e)[:200]}"

    return datos


def leer_folios_excel() -> list[dict]:
    """Lee el Excel y devuelve lista de dicts con folio, año, fondo, región."""
    wb = openpyxl.load_workbook(EXCEL_INPUT, data_only=True)
    ws = wb.active
    folios = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[2]:
            continue
        folios.append({
            "fondo": row[0],
            "ano": int(row[1]) if row[1] else 2022,
            "folio": str(row[2]).strip(),
            "region_excel": row[5] or "",
        })
    return folios


def cargar_progreso() -> set[str]:
    """Lee el CSV existente y devuelve folios ya procesados sin error."""
    procesados = set()
    if Path(CSV_OUTPUT).exists():
        with open(CSV_OUTPUT, encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                # Reprocesamos folios con error de SESION_EXPIRADA o TIMEOUT
                err = row.get("error", "")
                if err in ("", "FOLIO_NO_EXISTE"):
                    procesados.add(row["folio"])
    return procesados


def main():
    print("=" * 70)
    print("EXTRACCIÓN DE DATOS SIRA - 250 FOLIOS")
    print("=" * 70)

    Path("logs").mkdir(exist_ok=True)

    # Cargar folios objetivo
    folios = leer_folios_excel()
    total = len(folios)
    print(f"\nFolios a procesar: {total}")

    # Cargar progreso previo
    procesados = cargar_progreso()
    pendientes = [f for f in folios if f["folio"] not in procesados]
    print(f"Ya procesados: {len(procesados)}")
    print(f"Pendientes: {len(pendientes)}")

    if not pendientes:
        print("\n[OK] No hay folios pendientes. Todo listo.")
        return

    # Abrir CSV en modo append (escribir header solo si es archivo nuevo)
    csv_existe = Path(CSV_OUTPUT).exists()
    f_csv = open(CSV_OUTPUT, "a", encoding="utf-8-sig", newline="")
    writer = csv.DictWriter(f_csv, fieldnames=CAMPOS_CSV, extrasaction="ignore")
    if not csv_existe:
        writer.writeheader()
        f_csv.flush()

    f_jsonl = open(JSONL_OUTPUT, "a", encoding="utf-8")

    driver = crear_driver()
    n_exitos = 0
    n_errores = 0
    n_no_existe = 0

    try:
        esperar_login(driver)

        for i, folio_info in enumerate(pendientes, 1):
            folio = folio_info["folio"]
            ano = folio_info["ano"]
            print(f"[{i:>3}/{len(pendientes)}] Folio {folio} (año {ano})...", end=" ", flush=True)

            datos = extraer_datos_folio(driver, folio, ano)
            # Enriquecer con metadata del Excel
            datos["fondo"] = folio_info["fondo"]
            datos["region_excel"] = folio_info["region_excel"]

            # Detectar sesión expirada y pedir relogin
            if datos["error"] == "SESION_EXPIRADA":
                print("⚠️  SESIÓN EXPIRADA")
                print("\nLa sesión Clave Única expiró. Hacé login de nuevo en Chrome,")
                print("después volvé y presioná ENTER para reanudar.")
                input(">>> ")
                # Reintentar este folio
                datos = extraer_datos_folio(driver, folio, ano)
                datos["fondo"] = folio_info["fondo"]
                datos["region_excel"] = folio_info["region_excel"]

            # Escribir resultado
            writer.writerow(datos)
            f_csv.flush()
            f_jsonl.write(json.dumps(datos, ensure_ascii=False) + "\n")
            f_jsonl.flush()

            # Print resumen
            if datos["error"] == "FOLIO_NO_EXISTE":
                n_no_existe += 1
                print(f"❌ NO EXISTE EN OG1")
            elif datos["error"]:
                n_errores += 1
                print(f"❌ {datos['error']}")
            else:
                n_exitos += 1
                rut = datos["rut"] or "?"
                razon = (datos["razon_social"] or "?")[:45]
                estado = datos["estado_sira"] or "?"
                print(f"✓ {rut:13s} | {estado:9s} | {razon}")

            time.sleep(SLEEP_ENTRE_FOLIOS)

    finally:
        f_csv.close()
        f_jsonl.close()
        print("\n" + "=" * 70)
        print(f"RESUMEN")
        print("=" * 70)
        print(f"  Procesados con éxito:  {n_exitos}")
        print(f"  No existen en OG1:     {n_no_existe}")
        print(f"  Errores:               {n_errores}")
        print(f"\n  CSV   -> {CSV_OUTPUT}")
        print(f"  JSONL -> {JSONL_OUTPUT}")
        print("\nDejé el navegador abierto. ENTER para cerrarlo.")
        input(">>> ")
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
