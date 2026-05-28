"""
Script 02-TEST v2: Prueba de extracción de SIRA (3 folios)
============================================================
Versión corregida con las URLs reales de SIRA.

Patrón detectado:
  - 2022:  ?q={folio}#/.../SUBSECRETARÍA GENERAL DE GOBIERNO/{folio}
  - 2024+: ?q={folio}#/.../_/{folio}

Uso:
    python scripts\02_test_3folios.py
"""

import time
import json
from pathlib import Path
from urllib.parse import quote
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# ====== CONFIGURACIÓN ======
CHROME_PROFILE_DIR = r"C:\bot_sira_chrome_profile"
TIMEOUT = 20

# Plantillas de URL según año
SEGEGOB = quote("SECRETARÍA GENERAL DE GOBIERNO")
SUBSEG = quote("SUBSECRETARÍA GENERAL DE GOBIERNO")
BASE = "https://sira.auditoriainternadegobierno.gob.cl"

URL_2022 = BASE + "/?q={folio}#/convenio/" + SEGEGOB + "/" + SUBSEG + "/" + SUBSEG + "/{folio}"
URL_2024 = BASE + "/?q={folio}#/convenio/" + SEGEGOB + "/" + SUBSEG + "/_/{folio}"


def url_para_folio(folio: str, ano: int) -> str:
    if ano <= 2022:
        return URL_2022.format(folio=folio)
    return URL_2024.format(folio=folio)


FOLIOS_TEST = [
    ("55828", 2022),
    ("54925", 2022),
    ("100126", 2024),
]


def crear_driver():
    Path(CHROME_PROFILE_DIR).mkdir(parents=True, exist_ok=True)
    options = Options()
    options.add_argument(f"--user-data-dir={CHROME_PROFILE_DIR}")
    options.add_argument("--start-maximized")
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
    print("PASO 1: LOGIN MANUAL")
    print("=" * 70)
    print("Si Chrome todavía tiene la sesión del intento anterior, deberías")
    print("estar logueado ya. Si no, hacé login con Clave Única.")
    print("Cuando estés en SIRA, presioná ENTER.")
    print("=" * 70)
    driver.get(BASE + "/")
    input("\n>>> Presioná ENTER cuando estés logueado: ")
    print("Continuando...\n")


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


def extraer_datos_folio(driver, folio: str, ano: int) -> dict:
    url = url_para_folio(folio, ano)
    print(f"\n[Cargando año={ano}] {url}")
    driver.get(url)

    datos = {
        "folio": folio,
        "ano": ano,
        "url": url,
        "current_url_despues_carga": "",
        "body_text_raw": "",
        "campos_parseados": {},
        "razon_social": "",
        "estado_detectado": "",
        "error": "",
    }

    try:
        WebDriverWait(driver, TIMEOUT).until(
            EC.any_of(
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Datos declarados')]")),
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'no existe')]")),
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'RUT ENTIDAD')]")),
            )
        )
        time.sleep(2)

        datos["current_url_despues_carga"] = driver.current_url
        body_text = driver.find_element(By.TAG_NAME, "body").text
        datos["body_text_raw"] = body_text

        if "no existe en el catálogo" in body_text.lower():
            datos["error"] = "FOLIO_NO_EXISTE"
            return datos

        campos = {
            "rut": extraer_despues_de(body_text, "RUT ENTIDAD RECEPTORA"),
            "tipo_entidad": extraer_despues_de(body_text, "TIPO DE ENTIDAD RECEPTORA"),
            "fecha_suscripcion": extraer_despues_de(body_text, "FECHA DE SUSCRIPCIÓN"),
            "modalidad": extraer_despues_de(body_text, "MODALIDAD DE ASIGNACIÓN"),
            "tipo_convenio": extraer_despues_de(body_text, "TIPO DE CONVENIO"),
            "alcance": extraer_despues_de(body_text, "ALCANCE"),
            "cuotas_declaradas": extraer_despues_de(body_text, "CUOTAS DECLARADAS"),
            "tipo_garantia": extraer_despues_de(body_text, "TIPO DE GARANTÍA"),
            "monto_garantia": extraer_despues_de(body_text, "MONTO GARANTÍA"),
            "toma_razon": extraer_despues_de(body_text, "TOMA DE RAZÓN CGR"),
            "estados_financieros": extraer_despues_de(body_text, "ESTADOS FINANCIEROS"),
            "socios_beneficiarios": extraer_despues_de(body_text, "SOCIOS / BENEFICIARIOS FINALES"),
        }
        datos["campos_parseados"] = campos

        # Razón social - buscar línea con " / " después del folio
        for linea in body_text.split("\n"):
            linea = linea.strip()
            if " / " in linea and len(linea) > 10 and folio not in linea:
                datos["razon_social"] = linea
                break

        for kw in ["Borrador", "Cerrado", "Enviado", "Vigente"]:
            if kw in body_text[:1000]:
                datos["estado_detectado"] = kw
                break

    except TimeoutException:
        datos["error"] = f"TIMEOUT_{TIMEOUT}s"
    except Exception as e:
        datos["error"] = f"{type(e).__name__}: {str(e)[:200]}"

    return datos


def main():
    print("=" * 70)
    print("PRUEBA DE EXTRACCIÓN SIRA v2 - 3 FOLIOS")
    print("=" * 70)

    Path("logs").mkdir(exist_ok=True)
    driver = crear_driver()

    try:
        esperar_login(driver)

        resultados = []
        for i, (folio, ano) in enumerate(FOLIOS_TEST, 1):
            print(f"\n{'#' * 70}")
            print(f"# FOLIO {i}/{len(FOLIOS_TEST)}: {folio} (año {ano})")
            print(f"{'#' * 70}")

            datos = extraer_datos_folio(driver, folio, ano)
            resultados.append(datos)

            if datos["error"]:
                print(f"\n[ERROR] {datos['error']}")
                print(f"[Body capturado, primeros 600 chars]")
                print("-" * 70)
                print(datos["body_text_raw"][:600])
                print("-" * 70)
            else:
                print(f"\n[URL final] {datos['current_url_despues_carga']}")
                print(f"[Razón social] {datos['razon_social']!r}")
                print(f"[Estado] {datos['estado_detectado']!r}")
                print("\n[Campos parseados]")
                for k, v in datos["campos_parseados"].items():
                    print(f"  {k:25s} = {v!r}")
                print("\n[Body text - primeros 1000 chars]")
                print("-" * 70)
                print(datos["body_text_raw"][:1000])
                print("-" * 70)

            time.sleep(2)

        with open("logs/test_3folios_resultado_v2.json", "w", encoding="utf-8") as f:
            json.dump(resultados, f, ensure_ascii=False, indent=2)
        print(f"\n[OK] Resultado guardado en logs/test_3folios_resultado_v2.json")

    finally:
        print("\n" + "=" * 70)
        print("FIN DE LA PRUEBA")
        print("=" * 70)
        print("Dejé el navegador abierto. ENTER para cerrarlo todo.")
        input(">>> ")
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
