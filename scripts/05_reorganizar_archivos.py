"""
Script 05: Reorganizar archivos al estándar Arica + auditoría de completitud
=============================================================================
1. Crea carpetas por folio en cada región.
2. Mueve garantía y rendición desde carpetas compartidas hacia la carpeta
   individual del folio con nombres estándar.
3. Audita cada carpeta contra los 4 documentos OBLIGATORIOS:
      convenio | resolución | egreso/transferencia | rendición
   y genera un reporte de qué le falta a cada folio.

Estructura objetivo por folio:
  Archivos/{REGION}/{folio} - {NOMBRE}/
    Convenio   - {folio}.pdf    ← obligatorio
    Resolución - {folio}.pdf    ← obligatorio
    Egreso     - {folio}.pdf    ← obligatorio (también acepta "Transferencia")
    Rendición  - {folio}.pdf    ← obligatorio
    Garantía   - {folio}.pdf    ← opcional

Modo por defecto: DRY-RUN (muestra qué haría sin tocar nada).
Usar --ejecutar para mover archivos realmente.

Uso:
    python scripts/05_reorganizar_archivos.py
    python scripts/05_reorganizar_archivos.py --ejecutar
    python scripts/05_reorganizar_archivos.py --region ATACAMA --ejecutar
"""

import argparse
import csv
import re
import shutil
import unicodedata
from pathlib import Path

# ====== CONFIGURACIÓN ======
BASE_ARCHIVOS = Path("/home/bgcorrea/personal/workspace/caigg/Archivos")
MASTER_CSV    = Path("logs/master_subida.csv")
LOG_MOVIDOS   = Path("logs/reorg_movidos.csv")
LOG_AUDITORIA = Path("logs/reorg_auditoria.csv")

REGION_DIRS = {
    "ATACAMA":   "04. ATACAMA - FFOIP",
    "OHIGGINS":  "08. O'HIGGINS - FFOIP",
    "NUBLE":     "10. ÑUBLE - FFOIP",
    "AYSEN":     "15. AYSEN - FFOIP",
    "ARICA":     "01. ARICA - FFOIP",
    "RM":        "07. METROPOLITANA - FFOIP",
}

# Palabras clave para detectar cada tipo de documento por nombre de archivo
KEYWORDS_DOC = {
    "convenio":   ["convenio"],
    "resolucion": ["resolución", "resolucion", "res.", "rex", "adjudicación", "adjudicacion"],
    "egreso":     ["egreso", "transferencia", "recepción", "recepcion", "voucher"],
    "rendicion":  ["rendición", "rendicion", "cfc", "fiel cumplimiento", "memo daf"],
    "garantia":   ["garantía", "garantia", "letra de cambio"],
}

FOLIO_PATRON = re.compile(r"\b(5\d{4}|6\d{4}|7\d{4}|8\d{4}|9\d{4}|1\d{5})\b")


# ====== UTILIDADES ======

def normalizar(texto: str) -> str:
    if not texto:
        return ""
    nfd = unicodedata.normalize("NFD", texto)
    sin_tildes = "".join(c for c in nfd if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", sin_tildes.upper().strip())


def detectar_tipo(nombre_archivo: str) -> str:
    """Detecta qué tipo de documento es según palabras clave en el nombre."""
    nombre_lower = nombre_archivo.lower()
    for tipo, keywords in KEYWORDS_DOC.items():
        if any(kw in nombre_lower for kw in keywords):
            return tipo
    return "otro"


def auditar_carpeta(carpeta: Path) -> dict[str, bool]:
    """
    Revisa qué documentos obligatorios existen en la carpeta.
    Devuelve {convenio, resolucion, egreso, rendicion, garantia} → bool
    """
    estado = {k: False for k in KEYWORDS_DOC}
    if not carpeta.exists():
        return estado
    for f in carpeta.iterdir():
        if not f.is_file():
            continue
        tipo = detectar_tipo(f.name)
        if tipo in estado:
            estado[tipo] = True
    return estado


def nombre_carpeta(folio: str, razon_social: str) -> str:
    return f"{folio} - {razon_social}"


def mover_archivo(src: Path, dst: Path, dry_run: bool) -> tuple[bool, str]:
    if dst.exists():
        return False, "YA_EXISTE"
    if not src.exists():
        return False, "FUENTE_NO_EXISTE"
    if not dry_run:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(src), str(dst))
    return True, "DRY_RUN" if dry_run else "OK"


def copiar_archivo(src: Path, dst: Path, dry_run: bool) -> tuple[bool, str]:
    if dst.exists():
        return False, "YA_EXISTE"
    if not src.exists():
        return False, "FUENTE_NO_EXISTE"
    if not dry_run:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(src), str(dst))
    return True, "DRY_RUN" if dry_run else "OK"


# ====== CARGA DEL MASTER ======

def cargar_master() -> dict[str, dict]:
    resultado = {}
    with open(MASTER_CSV, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            resultado[r["folio"]] = r
    return resultado


def region_canonica(region_raw: str) -> str:
    r = normalizar(region_raw)
    if "ARICA" in r:      return "ARICA"
    if "ATACAMA" in r:    return "ATACAMA"
    if "VALPARAISO" in r: return "VALPARAISO"
    if "HIGGINS" in r:    return "OHIGGINS"
    if "NUBLE" in r:      return "NUBLE"
    if "AYSEN" in r:      return "AYSEN"
    if r == "RM":         return "RM"
    return r


# ====== INDEXADORES (extraen folio → archivo desde carpetas compartidas) ======

def indexar_atacama(region_dir: Path) -> tuple[dict, dict]:
    """FOLIO {n}- nombre.pdf"""
    def _idx(carpeta):
        out = {}
        for f in carpeta.iterdir():
            if not f.is_file():
                continue
            m = re.search(r"FOLIO\s*(\d+)", f.name, re.IGNORECASE)
            folio = m.group(1) if m else None
            if folio == "588222":   # typo conocido
                folio = "58822"
            if folio:
                out[folio] = f
        return out
    gar = _idx(region_dir / "01. GARANTÍAS")
    ren = _idx(region_dir / "02. RENDICIÓN")
    return gar, ren


def indexar_ohiggins(region_dir: Path) -> tuple[dict, dict]:
    """Garantías: {n} - nombre.pdf | Rendición: carpetas {n} - nombre/ con PDF adentro"""
    gar = {}
    for f in (region_dir / "01. GARANTÍAS").iterdir():
        if not f.is_file():
            continue
        m = re.match(r"^(\d+)\s*-", f.name)
        if m:
            gar[m.group(1)] = f

    ren = {}
    for sub in (region_dir / "02. RENDICIÓN").iterdir():
        if not sub.is_dir():
            continue
        m = re.match(r"^(\d+)\s*-", sub.name)
        if not m:
            continue
        pdfs = list(sub.glob("*.pdf")) + list(sub.glob("*.PDF"))
        if pdfs:
            ren[m.group(1)] = pdfs[0]
    return gar, ren


def indexar_nuble(region_dir: Path) -> tuple[dict, dict]:
    """{n} nombre.pdf (sin guión)"""
    def _idx(carpeta):
        out = {}
        for f in carpeta.iterdir():
            if not f.is_file():
                continue
            m = re.match(r"^(\d{5,6})\s", f.name)
            if m:
                out[m.group(1)] = f
        return out
    return _idx(region_dir / "01. GARANTÍAS"), _idx(region_dir / "02. RENDICIÓN")


def indexar_aysen(region_dir: Path) -> tuple[dict, dict]:
    """Garantías en Eliminar/: '{idx} {n} nombre.pdf' | Rendición: 'FRR {n} nombre.pdf'"""
    gar = {}
    eliminar_dir = region_dir / "01. GARANTÍAS" / "Eliminar"
    if eliminar_dir.exists():
        for f in eliminar_dir.iterdir():
            if not f.is_file():
                continue
            m = re.search(r"\b(\d{5,6})\b", f.name)
            if m:
                gar[m.group(1)] = f

    ren = {}
    for f in (region_dir / "02. RENDICIÓN").iterdir():
        if not f.is_file():
            continue
        m = re.search(r"FRR\s+(\d+)", f.name, re.IGNORECASE)
        if m:
            ren[m.group(1)] = f
    return gar, ren


INDEXADORES = {
    "ATACAMA":  indexar_atacama,
    "OHIGGINS": indexar_ohiggins,
    "NUBLE":    indexar_nuble,
    "AYSEN":    indexar_aysen,
}


# ====== PROCESADOR GENÉRICO (Atacama, O'Higgins, Ñuble, Aysén) ======

def procesar_region(region: str, master: dict, dry_run: bool) -> tuple[list, list]:
    """
    Devuelve (logs_movidos, logs_auditoria).
    logs_movidos: qué se movió o intentó mover.
    logs_auditoria: estado de completitud de cada carpeta tras el proceso.
    """
    logs_mov = []
    region_dir = BASE_ARCHIVOS / REGION_DIRS[region]
    folios_region = {f: d for f, d in master.items() if region_canonica(d["region"]) == region}

    gar_idx, ren_idx = INDEXADORES[region](region_dir)

    for folio, datos in folios_region.items():
        razon = datos["razon_social"]
        carpeta_folio = region_dir / nombre_carpeta(folio, razon)

        for tipo, idx, prefijo in [
            ("garantia",  gar_idx, "Garantía"),
            ("rendicion", ren_idx, "Rendición"),
        ]:
            src = idx.get(folio)
            if not src:
                logs_mov.append({
                    "region": region, "folio": folio, "tipo": tipo,
                    "accion": "SIN_ARCHIVO",
                    "detalle": f"No encontrado en carpeta compartida"
                })
                continue
            ext = src.suffix.lower()
            dst = carpeta_folio / f"{prefijo} - {folio}{ext}"
            ok, estado = mover_archivo(src, dst, dry_run)
            logs_mov.append({
                "region": region, "folio": folio, "tipo": tipo,
                "accion": estado,
                "detalle": f"{src.name} → {dst.name}"
            })

        # Limpiar carpeta de rendición (O'Higgins: subcarpetas vacías)
        if not dry_run and region == "OHIGGINS":
            ren_dir = region_dir / "02. RENDICIÓN"
            if ren_dir.exists():
                for sub in ren_dir.iterdir():
                    if sub.is_dir() and not any(sub.iterdir()):
                        sub.rmdir()

    # Auditoría post-proceso (o simulada en dry-run)
    logs_aud = auditar_region(region, folios_region, region_dir, dry_run, gar_idx, ren_idx)
    return logs_mov, logs_aud


def auditar_region(region: str, folios_region: dict, region_dir: Path,
                   dry_run: bool, gar_idx: dict, ren_idx: dict) -> list:
    """
    Simula o verifica el estado final de cada carpeta.
    En dry-run, combina lo que ya hay en disco con lo que SE MOVERÍA.
    En ejecución real, lee directamente la carpeta.
    """
    logs = []
    for folio, datos in folios_region.items():
        razon = datos["razon_social"]
        carpeta_folio = region_dir / nombre_carpeta(folio, razon)

        if dry_run:
            # Simular: lo que ya hay + lo que se movería
            estado = auditar_carpeta(carpeta_folio)
            if folio in gar_idx and not estado["garantia"]:
                estado["garantia"] = True
            if folio in ren_idx and not estado["rendicion"]:
                estado["rendicion"] = True
        else:
            estado = auditar_carpeta(carpeta_folio)

        obligatorios = ["convenio", "resolucion", "egreso", "rendicion"]
        faltantes = [doc for doc in obligatorios if not estado[doc]]
        tiene = [doc for doc in obligatorios if estado[doc]]

        logs.append({
            "region":   region,
            "folio":    folio,
            "razon_social": razon,
            "convenio":   "✓" if estado["convenio"]   else "FALTA",
            "resolucion": "✓" if estado["resolucion"] else "FALTA",
            "egreso":     "✓" if estado["egreso"]     else "FALTA",
            "rendicion":  "✓" if estado["rendicion"]  else "FALTA",
            "garantia":   "✓" if estado["garantia"]   else "-",
            "completo":   "SI" if not faltantes else "NO",
            "faltantes":  ", ".join(faltantes) if faltantes else "",
        })
    return logs


# ====== ARICA (tratamiento especial) ======

def procesar_arica(master: dict, dry_run: bool) -> tuple[list, list]:
    logs_mov = []
    region = "ARICA"
    region_dir = BASE_ARCHIVOS / REGION_DIRS[region]
    ren_dir = region_dir / "02. RENDICIÓN"
    folios_region = {f: d for f, d in master.items() if region_canonica(d["region"]) == region}

    # Mapear carpetas existentes por folio (puede haber duplicados)
    carpetas: dict[str, Path] = {}
    for sub in region_dir.iterdir():
        if not sub.is_dir():
            continue
        m = re.match(r"^(\d+)", sub.name)
        if m:
            folio = m.group(1)
            if folio not in carpetas:
                carpetas[folio] = sub
            else:
                logs_mov.append({
                    "region": region, "folio": folio, "tipo": "carpeta",
                    "accion": "DUPLICADO",
                    "detalle": f"{carpetas[folio].name} Y {sub.name}"
                })

    # Intentar mover los archivos de 02. RENDICIÓN por similitud de nombre
    if ren_dir.exists():
        for f in ren_dir.iterdir():
            if not f.is_file():
                continue
            nombre_norm = normalizar(f.stem)
            mejor_folio, mejor_score = None, 0
            for folio, carpeta in carpetas.items():
                razon = normalizar(folios_region.get(folio, {}).get("razon_social", ""))
                palabras = [p for p in razon.split() if len(p) >= 4]
                score = sum(1 for p in palabras if p in nombre_norm)
                if score > mejor_score:
                    mejor_score = score
                    mejor_folio = folio
            if mejor_folio and mejor_score >= 2:
                carpeta_dst = carpetas[mejor_folio]
                dst = carpeta_dst / f"Rendición - {mejor_folio}{f.suffix.lower()}"
                ok, estado = mover_archivo(f, dst, dry_run)
                logs_mov.append({
                    "region": region, "folio": mejor_folio, "tipo": "rendicion",
                    "accion": estado,
                    "detalle": f"{f.name} → {dst.name} (score={mejor_score})"
                })
            else:
                logs_mov.append({
                    "region": region, "folio": "?", "tipo": "rendicion",
                    "accion": "SIN_MATCH",
                    "detalle": f"{f.name} sin coincidencia"
                })

    # Auditoría de todas las carpetas de Arica
    logs_aud = []
    for folio, datos in folios_region.items():
        razon = datos["razon_social"]
        carpeta = carpetas.get(folio)
        if not carpeta:
            logs_aud.append({
                "region": region, "folio": folio, "razon_social": razon,
                "convenio": "FALTA", "resolucion": "FALTA",
                "egreso": "FALTA", "rendicion": "FALTA", "garantia": "-",
                "completo": "NO", "faltantes": "SIN CARPETA",
            })
            continue

        estado = auditar_carpeta(carpeta)
        obligatorios = ["convenio", "resolucion", "egreso", "rendicion"]
        faltantes = [doc for doc in obligatorios if not estado[doc]]
        logs_aud.append({
            "region":   region,
            "folio":    folio,
            "razon_social": razon,
            "convenio":   "✓" if estado["convenio"]   else "FALTA",
            "resolucion": "✓" if estado["resolucion"] else "FALTA",
            "egreso":     "✓" if estado["egreso"]     else "FALTA",
            "rendicion":  "✓" if estado["rendicion"]  else "FALTA",
            "garantia":   "✓" if estado["garantia"]   else "-",
            "completo":   "SI" if not faltantes else "NO",
            "faltantes":  ", ".join(faltantes) if faltantes else "",
        })

    return logs_mov, logs_aud


# ====== RM (tratamiento especial) ======

STOP_RAZON_RM = {
    "JUNTA", "VECINOS", "JJVV", "CLUB", "ADULTO", "MAYOR",
    "AGRUPACION", "AGRUPACION", "FUNDACION", "CORPORACION",
    "CENTRO", "ASOCIACION", "COMITE", "UNION", "COMUNAL",
    "COMUNIDAD", "SOCIAL", "CULTURAL", "DEPORTIVO", "DEPORTIVA",
    "DE", "DEL", "LA", "EL", "LAS", "LOS", "EN", "POR", "PARA",
    "CON", "SIN", "Y", "ORG", "CORP", "FUND",
}


def palabras_clave_rm(razon: str, min_largo: int = 5) -> list[str]:
    norm = normalizar(razon)
    return [p for p in re.findall(r"\w+", norm)
            if len(p) >= min_largo and p not in STOP_RAZON_RM]


def procesar_rm(master: dict, dry_run: bool) -> tuple[list, list]:
    """
    RM: distribuye archivos desde carpetas compartidas a carpetas de folio.
    - Egreso: 10. Egresos (TT)/{RUT}.pdf → match por RUT body
    - Convenio PDF: 08/Convenios Escaneados/ → match por razón social (fuzzy)
    - Resolución: 08/Res. Aprueba Convenios (*)/{n}-{m}.pdf → match por acto_admin_numero
    - Las carpetas de folio están en 07. Firma de Convenios/
    """
    logs_mov = []
    region = "RM"
    region_dir    = BASE_ARCHIVOS / "07. METROPOLITANA - FFOIP"
    firma_dir     = region_dir / "07. Firma de Convenios"
    conv_scan_dir = (region_dir / "08. Res. aprueba convenio (Res+ Memo TT)"
                     / "Convenios Escaneados")
    egresos_dir   = region_dir / "10. Egresos (TT)"
    res_local_dir = (region_dir / "08. Res. aprueba convenio (Res+ Memo TT)"
                     / "Res. Aprueba Convenios (Locales y Regionales)")
    res_nac_dir   = (region_dir / "08. Res. aprueba convenio (Res+ Memo TT)"
                     / "Res. Aprueba Convenios (Nacionales)")

    folios_region = {f: d for f, d in master.items()
                     if region_canonica(d["region"]) == region}

    # Indexar egresos: clave = RUT body (dígitos antes del guión)
    egresos_idx: dict[str, Path] = {}
    if egresos_dir.exists():
        for f in egresos_dir.iterdir():
            if f.is_file() and f.suffix.lower() == ".pdf":
                body = f.stem.replace(".", "").split("-")[0].strip()
                if body:
                    egresos_idx[body] = f

    # Indexar resoluciones: clave = número normalizado (solo dígitos)
    resoluciones_idx: dict[str, Path] = {}
    for res_dir in [res_local_dir, res_nac_dir]:
        if not res_dir.exists():
            continue
        for f in res_dir.iterdir():
            if f.is_file() and f.suffix.lower() == ".pdf":
                num_key = re.sub(r"[^0-9]", "", f.stem)
                if num_key:
                    resoluciones_idx[num_key] = f

    # Lista de convenios escaneados para fuzzy matching
    conv_escaneados = (
        [f for f in conv_scan_dir.iterdir()
         if f.is_file() and f.suffix.lower() == ".pdf"]
        if conv_scan_dir.exists() else []
    )

    for folio, datos in folios_region.items():
        razon    = datos["razon_social"]
        rut      = datos.get("rut", "")
        acto_num = datos.get("acto_admin_numero", "")

        carpeta_folio = firma_dir / nombre_carpeta(folio, razon)
        if not carpeta_folio.exists():
            logs_mov.append({
                "region": region, "folio": folio, "tipo": "carpeta",
                "accion": "SIN_CARPETA", "detalle": carpeta_folio.name
            })
            continue

        # --- Egreso por RUT ---
        rut_body = rut.replace(".", "").split("-")[0].strip() if rut else ""
        egreso_src = egresos_idx.get(rut_body) if rut_body else None
        if egreso_src:
            dst = carpeta_folio / f"Egreso - {folio}.pdf"
            ok, estado = copiar_archivo(egreso_src, dst, dry_run)
            logs_mov.append({
                "region": region, "folio": folio, "tipo": "egreso",
                "accion": estado,
                "detalle": f"{egreso_src.name} → {dst.name}"
            })
        else:
            logs_mov.append({
                "region": region, "folio": folio, "tipo": "egreso",
                "accion": "SIN_MATCH_RUT", "detalle": f"RUT={rut}"
            })

        # --- Convenio PDF desde Convenios Escaneados (fuzzy razón social) ---
        palabras = palabras_clave_rm(razon)
        mejor_conv, mejor_score = None, 0
        for f in conv_escaneados:
            nombre_norm = normalizar(f.stem)
            score = sum(1 for p in palabras if p in nombre_norm)
            if score > mejor_score:
                mejor_score = score
                mejor_conv = f
        if mejor_conv and mejor_score >= 2:
            dst = carpeta_folio / f"Convenio - {folio}.pdf"
            ok, estado = copiar_archivo(mejor_conv, dst, dry_run)
            logs_mov.append({
                "region": region, "folio": folio, "tipo": "convenio",
                "accion": estado,
                "detalle": f"{mejor_conv.name} → {dst.name} (score={mejor_score})"
            })
        else:
            logs_mov.append({
                "region": region, "folio": folio, "tipo": "convenio",
                "accion": "SIN_MATCH_RAZON",
                "detalle": f"mejor_score={mejor_score} razon={razon[:40]}"
            })

        # --- Resolución por número de acto administrativo ---
        if acto_num:
            acto_key = re.sub(r"[^0-9]", "", str(acto_num))
            res_src = resoluciones_idx.get(acto_key)
            if res_src:
                dst = carpeta_folio / f"Resolución - {folio}.pdf"
                ok, estado = copiar_archivo(res_src, dst, dry_run)
                logs_mov.append({
                    "region": region, "folio": folio, "tipo": "resolucion",
                    "accion": estado,
                    "detalle": f"{res_src.name} → {dst.name}"
                })
            else:
                logs_mov.append({
                    "region": region, "folio": folio, "tipo": "resolucion",
                    "accion": "SIN_MATCH_NUMERO",
                    "detalle": f"acto_num={acto_num} key={acto_key}"
                })

    # Auditoría
    logs_aud = []
    for folio, datos in folios_region.items():
        razon   = datos["razon_social"]
        carpeta = firma_dir / nombre_carpeta(folio, razon)
        estado  = auditar_carpeta(carpeta)
        oblig   = ["convenio", "resolucion", "egreso", "rendicion"]
        falt    = [d for d in oblig if not estado[d]]
        logs_aud.append({
            "region": region, "folio": folio, "razon_social": razon,
            "convenio":   "✓" if estado["convenio"]   else "FALTA",
            "resolucion": "✓" if estado["resolucion"] else "FALTA",
            "egreso":     "✓" if estado["egreso"]     else "FALTA",
            "rendicion":  "✓" if estado["rendicion"]  else "FALTA",
            "garantia":   "✓" if estado["garantia"]   else "-",
            "completo":   "SI" if not falt else "NO",
            "faltantes":  ", ".join(falt) if falt else "",
        })
    return logs_mov, logs_aud


# ====== MAIN ======

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--ejecutar", action="store_true",
                        help="Mover archivos realmente (por defecto: dry-run)")
    parser.add_argument("--region", type=str, default=None,
                        help="Solo procesar esta región: ATACAMA, OHIGGINS, NUBLE, AYSEN, ARICA")
    args = parser.parse_args()
    dry_run = not args.ejecutar

    print("=" * 70)
    print("REORGANIZACIÓN DE ARCHIVOS AL ESTÁNDAR ARICA")
    print("=" * 70)
    if dry_run:
        print("** MODO DRY-RUN: no se moverá nada — resultado simulado **")
    print()

    master = cargar_master()

    regiones = (
        [args.region.upper()] if args.region
        else ["ATACAMA", "OHIGGINS", "NUBLE", "AYSEN", "ARICA", "RM"]
    )

    todos_movidos = []
    toda_auditoria = []

    for region in regiones:
        print(f"Procesando {region}...")
        if region == "ARICA":
            logs_mov, logs_aud = procesar_arica(master, dry_run)
        elif region == "RM":
            logs_mov, logs_aud = procesar_rm(master, dry_run)
        elif region in INDEXADORES:
            logs_mov, logs_aud = procesar_region(region, master, dry_run)
        else:
            print(f"  [WARN] región no implementada: {region}")
            continue

        todos_movidos.extend(logs_mov)
        toda_auditoria.extend(logs_aud)

        # Resumen de movimientos
        conteo_mov = {}
        for log in logs_mov:
            conteo_mov[log["accion"]] = conteo_mov.get(log["accion"], 0) + 1
        for accion, n in sorted(conteo_mov.items()):
            print(f"  {accion:20s} {n:>3d}")

        # Resumen de completitud
        completos = sum(1 for r in logs_aud if r["completo"] == "SI")
        total     = len(logs_aud)
        print(f"  Completitud: {completos}/{total} folios tienen los 4 docs")
        print()

    # Guardar logs
    Path("logs").mkdir(exist_ok=True)

    with open(LOG_MOVIDOS, "w", encoding="utf-8-sig", newline="") as f:
        campos = ["region", "folio", "tipo", "accion", "detalle"]
        w = csv.DictWriter(f, fieldnames=campos)
        w.writeheader()
        w.writerows(todos_movidos)

    campos_aud = ["region", "folio", "razon_social",
                  "convenio", "resolucion", "egreso", "rendicion", "garantia",
                  "completo", "faltantes"]
    with open(LOG_AUDITORIA, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=campos_aud)
        w.writeheader()
        w.writerows(toda_auditoria)

    # Resumen global de auditoría
    print("=" * 70)
    print("AUDITORÍA DE COMPLETITUD POR REGIÓN")
    print("=" * 70)
    from collections import defaultdict
    por_region = defaultdict(lambda: {"total": 0, "completos": 0,
                                      "falta_convenio": 0, "falta_resolucion": 0,
                                      "falta_egreso": 0, "falta_rendicion": 0})
    for r in toda_auditoria:
        reg = r["region"]
        por_region[reg]["total"] += 1
        if r["completo"] == "SI":
            por_region[reg]["completos"] += 1
        if r["convenio"]   == "FALTA": por_region[reg]["falta_convenio"]   += 1
        if r["resolucion"] == "FALTA": por_region[reg]["falta_resolucion"] += 1
        if r["egreso"]     == "FALTA": por_region[reg]["falta_egreso"]     += 1
        if r["rendicion"]  == "FALTA": por_region[reg]["falta_rendicion"]  += 1

    header = f"{'REGIÓN':12s} {'TOT':>4s} {'OK':>4s} {'fConv':>6s} {'fRes':>5s} {'fEgr':>5s} {'fRend':>6s}"
    print(header)
    print("-" * 50)
    for reg, s in sorted(por_region.items()):
        print(f"{reg:12s} {s['total']:>4d} {s['completos']:>4d} "
              f"{s['falta_convenio']:>6d} {s['falta_resolucion']:>5d} "
              f"{s['falta_egreso']:>5d} {s['falta_rendicion']:>6d}")

    total_global     = sum(s["total"]     for s in por_region.values())
    completos_global = sum(s["completos"] for s in por_region.values())
    print(f"\nTotal: {completos_global}/{total_global} folios completos")
    print(f"\nReporte detallado: {LOG_AUDITORIA}")
    if dry_run:
        print("\nPara ejecutar: python scripts/05_reorganizar_archivos.py --ejecutar")


if __name__ == "__main__":
    main()
