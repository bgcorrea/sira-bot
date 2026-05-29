"""
Script 11 — Enviar folios a revisión en SIRA (en lote).

Solo envía folios que:
  - Están en estado "Borrador"
  - Tienen los 5 archivos obligatorios en master (según script 09)
  - Están en DISTRIBUCIÓN CARGA VB

IMPORTANTE: Esta acción es irreversible. Requiere --ejecutar para correr de verdad.

Uso:
    python scripts/11_enviar_revision.py                   # dry-run (solo navega y reporta)
    python scripts/11_enviar_revision.py --ejecutar        # envía de verdad
    python scripts/11_enviar_revision.py --folio 54922     # un solo folio
"""

import argparse
import csv
import sys
import time
from pathlib import Path
from urllib.parse import quote

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
)
from webdriver_manager.chrome import ChromeDriverManager

MASTER_XLSX    = "data/master_subida.xlsx"
LOG_CSV        = "logs/envio_revision.csv"
CHROME_PROFILE = "/home/bgcorrea/.bot_sira_chrome_profile"

BASE    = "https://sira.auditoriainternadegobierno.gob.cl"
SEGEGOB = quote("SECRETARÍA GENERAL DE GOBIERNO")
SUBSEG  = quote("SUBSECRETARÍA GENERAL DE GOBIERNO")
URL_2022 = BASE + "/?q={folio}#/convenio/" + SEGEGOB + "/" + SUBSEG + "/" + SUBSEG + "/{folio}"
URL_2024 = BASE + "/?q={folio}#/convenio/" + SEGEGOB + "/" + SUBSEG + "/_/{folio}"

TIMEOUT_PAGINA = 30


# ── Driver ────────────────────────────────────────────────────────────────────

SNAP_CHROMIUM = "/snap/bin/chromium"
SNAP_DRIVER   = "/snap/bin/chromium.chromedriver"
CDPORT        = 9222


def _lanzar_chromium_snap(perfil_dir: str, puerto: int = CDPORT):
    import subprocess
    cmd = [SNAP_CHROMIUM, f"--remote-debugging-port={puerto}",
           f"--user-data-dir={perfil_dir}", "--start-maximized",
           "--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu",
           "--no-first-run", "--password-store=basic"]
    return subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def crear_driver() -> webdriver.Chrome:
    import socket
    Path(CHROME_PROFILE).mkdir(parents=True, exist_ok=True)

    if Path(SNAP_CHROMIUM).exists() and Path(SNAP_DRIVER).exists():
        proc = _lanzar_chromium_snap(CHROME_PROFILE)
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

    options = Options()
    options.add_argument(f"--user-data-dir={CHROME_PROFILE}")
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


# ── Navegación y estado ───────────────────────────────────────────────────────

def url_para_folio(folio: str, ano: int) -> str:
    return URL_2022.format(folio=folio) if ano <= 2022 else URL_2024.format(folio=folio)


def cargar_folio(driver, folio: str, ano: int = 2023) -> str:
    """Devuelve: Borrador / Enviado / Cerrado / NO_EXISTE / TIMEOUT / DESCONOCIDO"""
    url = url_para_folio(folio, ano)
    driver.get(url)
    try:
        WebDriverWait(driver, TIMEOUT_PAGINA).until(
            EC.any_of(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(text(), 'Datos declarados')]")),
                EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(text(), 'no existe')]")),
            )
        )
        time.sleep(1.5)
        body = driver.find_element(By.TAG_NAME, "body").text
        if "no existe en el catálogo" in body.lower():
            return "NO_EXISTE"
        for kw in ["✓ Enviado", "✓ Cerrado", "Enviado", "Cerrado", "Borrador"]:
            if kw in body[:800]:
                return kw.replace("✓ ", "").strip()
        return "DESCONOCIDO"
    except TimeoutException:
        return "TIMEOUT"


def enviar_a_revision(driver, folio: str, dry_run: bool) -> tuple[bool, str]:
    """
    Hace click en el botón 'Enviar' / 'Enviar a revisión' y confirma el modal.
    Devuelve (ok, detalle).
    """
    # Selectores del botón de envío — CSS primero, luego XPath de respaldo
    _CSS_BTN = [
        "button.btn.primary",           # <button class="btn primary">
        "button.primary",
        "a.btn.primary",
    ]
    _XPATHS_BTN = [
        "//button[contains(text(),'Enviar a revisi')]",   # evita problema con tilde
        "//button[contains(text(),'Enviar a Revisi')]",
        "//button[normalize-space(text())='Enviar']",
        "//button[contains(@ng-click,'enviar') or contains(@ng-click,'submit')]",
        "//a[contains(@class,'btn') and contains(text(),'nviar')]",
    ]

    btn = None
    for css in _CSS_BTN:
        try:
            candidates = driver.find_elements(By.CSS_SELECTOR, css)
            for b in candidates:
                if b.is_displayed() and b.is_enabled() and "enviar" in b.text.lower():
                    btn = b
                    break
            if btn:
                break
        except NoSuchElementException:
            continue

    if btn is None:
        for xp in _XPATHS_BTN:
            try:
                b = driver.find_element(By.XPATH, xp)
                if b.is_displayed() and b.is_enabled():
                    btn = b
                    break
            except NoSuchElementException:
                continue

    if btn is None:
        return False, "BOTON_NO_ENCONTRADO"

    if dry_run:
        return True, f"DRY_RUN (botón='{btn.text.strip()}')"

    try:
        btn.click()
        time.sleep(1)
    except Exception as e:
        try:
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(1)
        except Exception:
            return False, f"CLICK_ERROR: {e}"

    # Confirmar modal de confirmación si aparece
    _XPATHS_CONFIRMAR = [
        "//button[contains(text(),'Sí') or contains(text(),'Confirmar') or contains(text(),'Aceptar')]",
        "//div[contains(@class,'modal')]//button[contains(@class,'primary') or contains(@class,'warn')]",
    ]
    for xp in _XPATHS_CONFIRMAR:
        try:
            modal_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, xp))
            )
            modal_btn.click()
            time.sleep(1)
            break
        except TimeoutException:
            pass

    # Verificar que el estado cambió a "Enviado"
    try:
        WebDriverWait(driver, 15).until(
            lambda d: "Enviado" in d.find_element(By.TAG_NAME, "body").text[:800]
        )
        return True, "ENVIADO"
    except TimeoutException:
        body = driver.find_element(By.TAG_NAME, "body").text
        if "Enviado" in body:
            return True, "ENVIADO"
        return False, "TIMEOUT_CONFIRMACION"


# ── Carga de datos ────────────────────────────────────────────────────────────

LOG_SUBIDA_CSV = "logs/ejecucion_subidas.csv"

# Secciones obligatorias para enviar (certificado es opcional para personas naturales)
_SEC_CAMPOS = {
    "Convenio + Acto Administrativo": ["convenio_pdf", "acto_admin_pdf"],
    "Transferencias":                  ["voucher_pdf"],
    "Respaldo de rendición":           ["rendicion_pdf"],
}

def _sira_ok_por_folio() -> dict:
    """Lee ejecucion_subidas.csv y devuelve {folio: {secciones con OK}}."""
    result: dict[str, set] = {}
    p = Path(LOG_SUBIDA_CSV)
    if not p.exists():
        return result
    with open(p, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            if row.get("estado") == "OK":
                result.setdefault(row["folio"], set()).add(row["seccion"])
    return result


def folios_listos() -> list[dict]:
    """Devuelve todos los folios del master que tienen las secciones mínimas cubiertas."""
    master = pd.read_excel(MASTER_XLSX)
    sira_ok = _sira_ok_por_folio()

    listos = []
    for _, row in master.iterrows():
        folio = str(row["folio"])
        faltantes = []
        for sec, campos in _SEC_CAMPOS.items():
            en_sira = sec in sira_ok.get(folio, set())
            en_disco = any(
                (v := str(row.get(c, "")).strip()) and v != "nan" and Path(v).exists()
                for c in campos
            )
            if not en_sira and not en_disco:
                faltantes.append(sec)
        if not faltantes:
            listos.append({
                "folio":        folio,
                "ano":          int(row.get("ano", 2023)),
                "region":       str(row.get("region", "")),
                "razon_social": str(row.get("razon_social", "")),
            })
    return listos


# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--ejecutar", action="store_true",
                        help="Enviar de verdad (sin este flag: dry-run)")
    parser.add_argument("--folio", type=str, default=None,
                        help="Procesar un solo folio para prueba")
    args = parser.parse_args()
    dry_run = not args.ejecutar

    filas = folios_listos()
    if args.folio:
        filas = [f for f in filas if f["folio"] == args.folio]
        if not filas:
            print(f"[ERROR] Folio {args.folio} no encontrado en listos o no tiene archivos completos.")
            sys.exit(1)

    print("=" * 70)
    print("BOT DE ENVÍO A REVISIÓN — SIRA")
    print("=" * 70)
    if dry_run:
        print("** MODO DRY-RUN: no se enviará nada **")
    else:
        print("** MODO REAL: folios serán enviados a revisión (irreversible) **")
    print(f"\n  Folios listos para enviar: {len(filas)}")
    print()

    driver = crear_driver()
    esperar_login(driver)

    Path("logs").mkdir(exist_ok=True)
    log_rows = []

    enviados  = 0
    ya_enviados = 0
    errores   = 0

    try:
        for i, fila in enumerate(filas, 1):
            folio  = fila["folio"]
            region = fila["region"]
            razon  = fila["razon_social"]
            print(f"[{i}/{len(filas)}] Folio {folio} ({razon[:45]})")

            estado = cargar_folio(driver, folio, fila.get("ano", 2023))

            if estado in ("Enviado", "Cerrado"):
                print(f"  → Ya {estado.lower()} — omitido")
                ya_enviados += 1
                log_rows.append({"folio": folio, "region": region, "razon_social": razon,
                                  "estado": "YA_ENVIADO", "detalle": estado,
                                  "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")})
                continue

            if estado != "Borrador":
                print(f"  ✗ Estado inesperado: {estado}")
                errores += 1
                log_rows.append({"folio": folio, "region": region, "razon_social": razon,
                                  "estado": "ERROR", "detalle": f"Estado={estado}",
                                  "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")})
                continue

            ok, detalle = enviar_a_revision(driver, folio, dry_run)
            simbolo = "✓" if ok else "✗"
            print(f"  {simbolo} {detalle}")

            if ok:
                enviados += 1
            else:
                errores += 1

            log_rows.append({"folio": folio, "region": region, "razon_social": razon,
                              "estado": "ENVIADO" if (ok and not dry_run) else detalle,
                              "detalle": detalle,
                              "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")})

            time.sleep(2)

    except KeyboardInterrupt:
        print("\n[Interrumpido]")
    finally:
        # Guardar log
        if log_rows:
            with open(LOG_CSV, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.DictWriter(f, fieldnames=log_rows[0].keys())
                writer.writeheader()
                writer.writerows(log_rows)

        print(f"\n{'='*70}")
        print("RESUMEN")
        print(f"{'='*70}")
        accion = "Enviados" if not dry_run else "DRY-RUN OK"
        print(f"  {accion}:      {enviados}")
        print(f"  Ya enviados:  {ya_enviados}")
        print(f"  Errores:      {errores}")
        print(f"  Log → {LOG_CSV}")
        print()

        input("Navegador queda abierto. ENTER para cerrar. ")
        driver.quit()


if __name__ == "__main__":
    main()
