"""
Script 06: Validar estructura de archivos y estandarizar nombres
=================================================================
Dos funciones principales:

  MODO VALIDAR (por defecto):
    Revisa todas las carpetas por folio y reporta completitud.
    Detecta problemas: archivos faltantes, extensiones incorrectas,
    archivos duplicados del mismo tipo, nombres no estándar.

  MODO RENOMBRAR (--renombrar):
    Renombra los archivos que siguen el criterio de detección pero
    no tienen el nombre estándar `{Tipo} - {folio}.pdf`.
    Ejemplo: "Resolución Exenta N° 106-2022.pdf" → "Resolución - 59641.pdf"
    Solo actúa si el tipo es inequívoco (un solo match en la carpeta).
    Por defecto es dry-run; usar --ejecutar para aplicar.

Uso:
    python scripts/06_validar_estructura.py
    python scripts/06_validar_estructura.py --renombrar
    python scripts/06_validar_estructura.py --renombrar --ejecutar
    python scripts/06_validar_estructura.py --region ARICA
"""

import argparse
import csv
import re
import unicodedata
from pathlib import Path

# ====== CONFIGURACIÓN ======
BASE_ARCHIVOS  = Path("/home/bgcorrea/personal/workspace/caigg/Archivos")
MASTER_CSV     = Path("logs/master_subida.csv")
LOG_VALIDACION = Path("logs/validacion_estructura.csv")

REGION_DIRS = {
    "ARICA":      "01. ARICA - FFOIP",
    "ATACAMA":    "04. ATACAMA - FFOIP",
    "VALPARAISO": "06. VALPARAÍSO - FFOIP",
    "RM":         "07. METROPOLITANA - FFOIP",
    "OHIGGINS":   "08. O'HIGGINS - FFOIP",
    "NUBLE":      "10. ÑUBLE - FFOIP",
    "AYSEN":      "15. AYSEN - FFOIP",
}

# Regiones cuyas carpetas de folio están en una subcarpeta (no en la raíz)
REGION_SUBFOLDER = {
    "RM": "07. Firma de Convenios",
}

# Tipos obligatorios para SIRA
TIPOS_OBLIGATORIOS = ["convenio", "resolucion", "egreso", "rendicion"]
TIPOS_OPCIONALES   = ["garantia"]
TIPOS_TODOS        = TIPOS_OBLIGATORIOS + TIPOS_OPCIONALES

# Extensiones aceptadas por tipo
EXT_PDF_SOLO  = {".pdf"}
EXT_PDF_IMG   = {".pdf", ".jpg", ".jpeg", ".png"}

EXTENSIONES_VALIDAS = {
    "convenio":   EXT_PDF_SOLO,
    "resolucion": EXT_PDF_SOLO,
    "egreso":     EXT_PDF_IMG,
    "rendicion":  EXT_PDF_SOLO,
    "garantia":   EXT_PDF_IMG,
}

# Palabras clave para detectar tipo (orden importa: más específico primero)
KEYWORDS = {
    "convenio":   ["convenio"],
    "resolucion": ["resolución", "resolucion", "res. adj", "rex", "exenta", "adjudicación", "adjudicacion"],
    "egreso":     ["egreso", "transferencia", "recepción de recursos", "recepcion de recursos",
                   "voucher", "certificado bancario", "certificado recepción", "certificado recepcion"],
    "rendicion":  ["rendición", "rendicion", "cfc", "fiel cumplimiento", "memo daf"],
    "garantia":   ["garantía", "garantia", "letra de cambio"],
}

# Nombre estándar: "Tipo - {folio}.ext"
NOMBRE_ESTANDAR = {
    "convenio":   "Convenio",
    "resolucion": "Resolución",
    "egreso":     "Egreso",
    "rendicion":  "Rendición",
    "garantia":   "Garantía",
}

FOLIO_PATRON = re.compile(r"^(\d{5,6})\s*-")


def normalizar(texto: str) -> str:
    nfd = unicodedata.normalize("NFD", texto)
    sin_tildes = "".join(c for c in nfd if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", sin_tildes.upper().strip())


def detectar_tipo(nombre: str) -> str | None:
    """
    Detecta el tipo de documento. Prioriza el nombre estándar exacto,
    luego busca por palabras clave. Devuelve None si no puede determinar.
    """
    nombre_lower = nombre.lower()

    # 1. Nombre estándar exacto: "Convenio - 12345.pdf"
    for tipo, prefijo in NOMBRE_ESTANDAR.items():
        patron = re.compile(rf"^{re.escape(prefijo.lower())}\s*-\s*\d+", re.IGNORECASE)
        if patron.match(nombre_lower):
            return tipo

    # 2. Palabras clave (más específicas primero)
    for tipo, keywords in KEYWORDS.items():
        if any(kw in nombre_lower for kw in keywords):
            return tipo

    return None


def es_nombre_estandar(nombre: str, tipo: str, folio: str) -> bool:
    prefijo = NOMBRE_ESTANDAR[tipo]
    stem = Path(nombre).stem
    return stem == f"{prefijo} - {folio}"


def cargar_master() -> dict[str, dict]:
    with open(MASTER_CSV, encoding="utf-8-sig") as f:
        return {r["folio"]: r for r in csv.DictReader(f)}


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


# ====== VALIDACIÓN ======

def validar_carpeta(folio: str, carpeta: Path) -> list[dict]:
    """
    Analiza una carpeta de folio y devuelve una lista de hallazgos.
    Un hallazgo por tipo de documento (o por problema).
    """
    hallazgos = []

    if carpeta is None or not carpeta.exists():
        for tipo in TIPOS_OBLIGATORIOS:
            hallazgos.append({
                "folio": folio, "tipo": tipo,
                "estado": "SIN_CARPETA", "archivo": "", "problema": "Carpeta no existe"
            })
        return hallazgos

    # Agrupar archivos por tipo detectado
    por_tipo: dict[str, list[Path]] = {t: [] for t in TIPOS_TODOS}
    sin_tipo: list[Path] = []

    for f in sorted(carpeta.iterdir()):
        if not f.is_file():
            continue
        if f.suffix.lower() not in {".pdf", ".jpg", ".jpeg", ".png", ".docx"}:
            continue
        tipo = detectar_tipo(f.name)
        if tipo in por_tipo:
            por_tipo[tipo].append(f)
        else:
            sin_tipo.append(f)

    # Evaluar cada tipo
    for tipo in TIPOS_TODOS:
        archivos = por_tipo[tipo]
        obligatorio = tipo in TIPOS_OBLIGATORIOS

        if not archivos:
            hallazgos.append({
                "folio": folio, "tipo": tipo,
                "estado": "FALTA" if obligatorio else "OPCIONAL_AUSENTE",
                "archivo": "", "problema": ""
            })
            continue

        if len(archivos) > 1:
            # Si hay exactamente un archivo con nombre estándar, usarlo como canónico
            # y tratar el resto como suplementarios (no bloquean la validación).
            estandares = [f for f in archivos if es_nombre_estandar(f.name, tipo, folio)]
            if len(estandares) == 1:
                archivos = estandares  # continúa con el archivo canónico
            else:
                nombres = ", ".join(f.name for f in archivos)
                hallazgos.append({
                    "folio": folio, "tipo": tipo,
                    "estado": "DUPLICADO",
                    "archivo": nombres,
                    "problema": f"Hay {len(archivos)} archivos del mismo tipo"
                })
                continue

        archivo = archivos[0]
        ext = archivo.suffix.lower()
        exts_validas = EXTENSIONES_VALIDAS[tipo]
        problemas = []

        if ext not in exts_validas:
            if ext == ".docx":
                problemas.append("DOCX_NO_PDF: convertir a PDF antes de subir")
            else:
                problemas.append(f"EXT_INVALIDA: {ext}")

        estandar = es_nombre_estandar(archivo.name, tipo, folio)

        estado = "OK" if not problemas else "PROBLEMA"
        if not estandar and not problemas:
            estado = "NOMBRE_NO_ESTANDAR"

        hallazgos.append({
            "folio": folio, "tipo": tipo,
            "estado": estado,
            "archivo": archivo.name,
            "problema": " | ".join(problemas) if problemas else (
                f"Renombrar a '{NOMBRE_ESTANDAR[tipo]} - {folio}{ext}'" if not estandar else ""
            )
        })

    # Archivos no identificados (pueden ser documentos válidos sin nombre claro)
    for f in sin_tipo:
        hallazgos.append({
            "folio": folio, "tipo": "desconocido",
            "estado": "SIN_TIPO",
            "archivo": f.name,
            "problema": "No se pudo determinar el tipo"
        })

    return hallazgos


# ====== RENOMBRADO ======

def renombrar_carpeta(folio: str, carpeta: Path, dry_run: bool) -> list[dict]:
    """
    Renombra archivos al estándar `{Tipo} - {folio}.ext`.
    Solo actúa cuando el tipo es inequívoco (exactamente un archivo de ese tipo).
    """
    cambios = []
    if not carpeta.exists():
        return cambios

    por_tipo: dict[str, list[Path]] = {t: [] for t in TIPOS_TODOS}

    for f in sorted(carpeta.iterdir()):
        if not f.is_file():
            continue
        tipo = detectar_tipo(f.name)
        if tipo in por_tipo:
            por_tipo[tipo].append(f)

    for tipo, archivos in por_tipo.items():
        if len(archivos) == 0:
            continue
        if len(archivos) == 2 and tipo == "resolucion":
            # Caso Arica: adjudicación (suplementario) + exenta (Acto Administrativo).
            # Renombramos solo la exenta; la adjudicación se queda como referencia.
            exentas = [f for f in archivos if "exenta" in f.name.lower()]
            adj    = [f for f in archivos if "adjudicaci" in f.name.lower()]
            if len(exentas) == 1 and len(adj) == 1:
                archivos = exentas  # procesar solo la exenta
            else:
                continue  # ambigüedad real: no tocar
        elif len(archivos) != 1:
            continue  # ambigüedad o ausencia: no tocar
        archivo = archivos[0]
        if es_nombre_estandar(archivo.name, tipo, folio):
            continue  # ya tiene el nombre correcto

        ext = archivo.suffix.lower()
        nuevo_nombre = f"{NOMBRE_ESTANDAR[tipo]} - {folio}{ext}"
        nuevo_path = carpeta / nuevo_nombre

        if nuevo_path.exists():
            cambios.append({
                "folio": folio, "tipo": tipo,
                "original": archivo.name, "nuevo": nuevo_nombre,
                "accion": "OMITIDO", "detalle": "Ya existe el destino"
            })
            continue

        if not dry_run:
            archivo.rename(nuevo_path)

        cambios.append({
            "folio": folio, "tipo": tipo,
            "original": archivo.name, "nuevo": nuevo_nombre,
            "accion": "DRY_RUN" if dry_run else "OK", "detalle": ""
        })

    return cambios


# ====== MAIN ======

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--renombrar", action="store_true",
                        help="Renombrar archivos al nombre estándar")
    parser.add_argument("--ejecutar", action="store_true",
                        help="Aplicar cambios (por defecto solo muestra)")
    parser.add_argument("--region", type=str, default=None,
                        help="Procesar solo esta región")
    args = parser.parse_args()

    dry_run = not args.ejecutar
    modo = "RENOMBRAR" if args.renombrar else "VALIDAR"

    print("=" * 70)
    print(f"SCRIPT 06 — MODO: {modo}")
    if dry_run and args.renombrar:
        print("** DRY-RUN: no se renombrará nada **")
    print("=" * 70)

    master = cargar_master()

    regiones = (
        {args.region.upper(): REGION_DIRS[args.region.upper()]}
        if args.region and args.region.upper() in REGION_DIRS
        else REGION_DIRS
    )

    todos_hallazgos = []
    todos_cambios   = []

    for region_key, region_dir_name in regiones.items():
        region_dir = BASE_ARCHIVOS / region_dir_name
        if not region_dir.exists():
            print(f"[SKIP] {region_key}: carpeta no encontrada")
            continue

        # Algunas regiones tienen sus carpetas de folio en una subcarpeta
        subfolder = REGION_SUBFOLDER.get(region_key)
        if subfolder:
            region_dir = region_dir / subfolder

        folios_region = {f: d for f, d in master.items()
                         if region_canonica(d["region"]) == region_key}
        if not folios_region:
            continue

        print(f"\n{region_key} ({len(folios_region)} folios)")

        # Indexar carpetas existentes por folio
        carpetas: dict[str, Path] = {}
        for sub in region_dir.iterdir():
            if not sub.is_dir():
                continue
            m = FOLIO_PATRON.match(sub.name)
            if m:
                folio = m.group(1)
                if folio not in carpetas:
                    carpetas[folio] = sub

        for folio, datos in sorted(folios_region.items()):
            razon = datos["razon_social"]
            carpeta = carpetas.get(folio)

            if args.renombrar:
                cambios = renombrar_carpeta(folio, carpeta, dry_run) if carpeta else []
                todos_cambios.extend(cambios)
            else:
                hallazgos = validar_carpeta(folio, carpeta)
                todos_hallazgos.extend(hallazgos)

    if args.renombrar:
        # Reporte de renombrado
        print("\n" + "=" * 70)
        print("RESUMEN DE RENOMBRADO")
        print("=" * 70)
        conteo = {}
        for c in todos_cambios:
            conteo[c["accion"]] = conteo.get(c["accion"], 0) + 1
        for accion, n in sorted(conteo.items()):
            print(f"  {accion:20s} {n:>4d}")
        for c in todos_cambios:
            if c["accion"] in ("OK", "DRY_RUN"):
                print(f"  [{c['accion']}] {c['folio']} {c['tipo']}: {c['original']!r} → {c['nuevo']!r}")
        if dry_run:
            print("\nUsar --ejecutar para aplicar los cambios.")
    else:
        # Reporte de validación
        from collections import defaultdict
        print("\n" + "=" * 70)
        print("AUDITORÍA DE COMPLETITUD")
        print("=" * 70)

        # Resumen por región
        por_region: dict[str, dict] = defaultdict(lambda: {
            "total": 0, "completos": 0,
            "falta_convenio": 0, "falta_resolucion": 0,
            "falta_egreso": 0, "falta_rendicion": 0,
            "problemas": 0
        })

        folios_vistos: dict[str, set] = defaultdict(set)
        folios_completos: dict[str, set] = defaultdict(set)

        for h in todos_hallazgos:
            folio = h["folio"]
            # Determinar región desde master
            region = region_canonica(master.get(folio, {}).get("region", "?"))
            folios_vistos[region].add(folio)

            if h["estado"] in ("PROBLEMA", "DUPLICADO", "SIN_TIPO", "DOCX"):
                por_region[region]["problemas"] += 1
            falta = h["estado"] in ("FALTA", "SIN_CARPETA")
            if h["tipo"] == "convenio"   and falta:
                por_region[region]["falta_convenio"]   += 1
            if h["tipo"] == "resolucion" and falta:
                por_region[region]["falta_resolucion"] += 1
            if h["tipo"] == "egreso"     and falta:
                por_region[region]["falta_egreso"]     += 1
            if h["tipo"] == "rendicion"  and falta:
                por_region[region]["falta_rendicion"]  += 1

        # Calcular completos: folios sin ningún FALTA/SIN_CARPETA obligatorio
        folios_con_falta: dict[str, set] = defaultdict(set)
        for h in todos_hallazgos:
            if h["estado"] in ("FALTA", "SIN_CARPETA") and h["tipo"] in TIPOS_OBLIGATORIOS:
                region = region_canonica(master.get(h["folio"], {}).get("region", "?"))
                folios_con_falta[region].add(h["folio"])

        header = f"{'REGIÓN':12s} {'TOT':>4s} {'OK':>4s} {'fConv':>6s} {'fRes':>5s} {'fEgr':>5s} {'fRend':>6s} {'Prob':>5s}"
        print(header)
        print("-" * 55)
        total_g, ok_g = 0, 0
        for region in sorted(folios_vistos):
            total = len(folios_vistos[region])
            con_falta = len(folios_con_falta.get(region, set()))
            ok = total - con_falta
            s = por_region[region]
            total_g += total
            ok_g += ok
            print(f"{region:12s} {total:>4d} {ok:>4d} "
                  f"{s['falta_convenio']:>6d} {s['falta_resolucion']:>5d} "
                  f"{s['falta_egreso']:>5d} {s['falta_rendicion']:>6d} {s['problemas']:>5d}")
        print(f"\nTotal: {ok_g}/{total_g} folios completos (4 docs obligatorios presentes)")

        # Guardar CSV detallado
        Path("logs").mkdir(exist_ok=True)
        with open(LOG_VALIDACION, "w", encoding="utf-8-sig", newline="") as f:
            campos = ["folio", "tipo", "estado", "archivo", "problema"]
            w = csv.DictWriter(f, fieldnames=campos)
            w.writeheader()
            w.writerows(todos_hallazgos)
        print(f"\nDetalle completo: {LOG_VALIDACION}")

        # Mostrar problemas que no son simples faltantes
        problemas = [h for h in todos_hallazgos
                     if h["estado"] not in ("OK", "FALTA", "OPCIONAL_AUSENTE", "NOMBRE_NO_ESTANDAR")]
        if problemas:
            print(f"\n{'='*70}")
            print("PROBLEMAS QUE REQUIEREN ATENCIÓN:")
            for h in problemas:
                print(f"  [{h['estado']}] folio {h['folio']} / {h['tipo']}: {h['archivo']} — {h['problema']}")

        # Mostrar nombres no estándar
        no_estandar = [h for h in todos_hallazgos if h["estado"] == "NOMBRE_NO_ESTANDAR"]
        if no_estandar:
            print(f"\n{'='*70}")
            print(f"ARCHIVOS CON NOMBRE NO ESTÁNDAR ({len(no_estandar)}) — ejecutar con --renombrar:")
            for h in no_estandar[:20]:
                print(f"  folio {h['folio']} / {h['tipo']}: {h['archivo']!r} → {h['problema']}")
            if len(no_estandar) > 20:
                print(f"  ... y {len(no_estandar)-20} más (ver {LOG_VALIDACION})")


if __name__ == "__main__":
    main()
