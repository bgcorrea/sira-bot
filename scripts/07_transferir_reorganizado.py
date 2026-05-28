"""
Script 07 — Transferir archivos de FFOIP 2022 (REORGANIZADO) a carpetas por folio.

Lógica:
- Lee master_subida.xlsx para obtener folio → carpeta_folio destino
- Itera todos los archivos en REORGANIZADO (solo regiones relevantes)
- Extrae folio del nombre de archivo (ej. "Convenio - 54706.pdf" → 54706)
- Copia a carpeta destino si existe; reporta si no existe o ya hay archivo
- Si la carpeta no existe pero el folio está en el master, la crea con --crear-carpetas
- Desduplicación: si el mismo archivo ya fue procesado para ese folio, omitir

Uso:
    python scripts/07_transferir_reorganizado.py                              # dry-run
    python scripts/07_transferir_reorganizado.py --crear-carpetas             # dry-run + crea carpetas
    python scripts/07_transferir_reorganizado.py --ejecutar                   # copia real (no crea carpetas)
    python scripts/07_transferir_reorganizado.py --ejecutar --crear-carpetas  # todo
"""

import argparse
import re
import shutil
from pathlib import Path

import pandas as pd

# ── Paths ────────────────────────────────────────────────────────────────────
BASE    = Path("/home/bgcorrea/personal/workspace/caigg")
REORG   = BASE / "FFOIP 2022 (REORGANIZADO)"
MASTER  = BASE / "bot_sira" / "data" / "master_subida.xlsx"
ARCHIVOS = BASE / "Archivos"

# Carpetas REORGANIZADO → clave región en master
REGION_FOLDERS = {
    "03. Arica y Parinacota": "ARICA",
    "06. Atacama":            "ATACAMA",
    "08. Valparaíso":         "VALPARAÍSO",
    "09. O'Higgins":          "O'HIGGINS",
    "11. Ñuble":              "ÑUBLE",
    "16. Aysén":              "AYSEN",
    "18. RM":                 "RM",
}

# Carpeta base en Archivos/ para cada región
REGION_BASE = {
    "ARICA":      "01. ARICA - FFOIP",
    "ATACAMA":    "04. ATACAMA - FFOIP",
    "VALPARAÍSO": "06. VALPARAÍSO - FFOIP",
    "O'HIGGINS":  "08. O'HIGGINS - FFOIP",
    "ÑUBLE":      "10. ÑUBLE - FFOIP",
    "AYSEN":      "15. AYSEN - FFOIP",
    "RM":         "07. METROPOLITANA - FFOIP/07. Firma de Convenios",
}

NOMBRE_STD = re.compile(r"^.+\s-\s(\d{5,6})\.(pdf|PDF)$")


def cargar_master() -> tuple[dict[str, Path], dict[str, dict]]:
    """
    Devuelve:
      folio_map: {folio_str: carpeta_folio_path}   (solo folios con carpeta existente)
      master_all: {folio_str: {'region', 'razon_social'}}  (todos los 250)
    """
    df = pd.read_excel(MASTER)
    folio_map: dict[str, Path] = {}
    master_all: dict[str, dict] = {}

    for _, row in df.iterrows():
        folio = str(row["folio"])
        region = str(row["region"])
        razon = str(row["razon_social"]) if pd.notna(row["razon_social"]) else ""
        master_all[folio] = {"region": region, "razon_social": razon}

        if pd.notna(row["carpeta_folio"]) and str(row["carpeta_folio"]).strip():
            folio_map[folio] = Path(str(row["carpeta_folio"]))

    return folio_map, master_all


def construir_carpeta(folio: str, region: str, razon_social: str) -> Path | None:
    """Construye el path esperado para un folio sin carpeta existente."""
    base_key = region.upper()
    if base_key not in REGION_BASE:
        return None
    nombre_carpeta = f"{folio} - {razon_social}"
    return ARCHIVOS / REGION_BASE[base_key] / nombre_carpeta


def transferir(dry_run: bool, crear_carpetas: bool) -> None:
    folio_map, master_all = cargar_master()
    folios_validos = set(master_all.keys())

    resultados = {
        "copiado":          [],
        "ya_existe":        [],
        "creada_carpeta":   [],
        "sin_carpeta":      [],   # en master pero sin carpeta y --crear-carpetas no activo
        "folio_no_master":  [],
        "nombre_invalido":  [],
    }

    # folio → set de nombres de archivo ya procesados (deduplicación)
    ya_procesado: dict[str, set] = {}

    for folder_name, region in REGION_FOLDERS.items():
        region_dir = REORG / folder_name
        if not region_dir.exists():
            print(f"[WARN] No encontrado: {region_dir}")
            continue

        archivos = sorted(region_dir.rglob("*.pdf")) + sorted(region_dir.rglob("*.PDF"))

        for archivo in archivos:
            m = NOMBRE_STD.match(archivo.name)
            if not m:
                resultados["nombre_invalido"].append(
                    {"region": region, "archivo": str(archivo.relative_to(REORG))}
                )
                continue

            folio = m.group(1)

            if folio not in folios_validos:
                resultados["folio_no_master"].append(
                    {"region": region, "folio": folio,
                     "archivo": str(archivo.relative_to(REORG))}
                )
                continue

            # Deduplicación por folio + nombre de archivo
            clave = archivo.name.lower()
            if clave in ya_procesado.get(folio, set()):
                continue
            ya_procesado.setdefault(folio, set()).add(clave)

            # Determinar carpeta destino
            if folio in folio_map:
                destino_dir = folio_map[folio]
            else:
                # Folio en master pero sin carpeta_folio
                info = master_all[folio]
                destino_dir = construir_carpeta(folio, info["region"], info["razon_social"])
                if destino_dir is None:
                    resultados["sin_carpeta"].append(
                        {"region": region, "folio": folio, "archivo": archivo.name,
                         "nota": "región no reconocida"}
                    )
                    continue

                if not destino_dir.exists():
                    if crear_carpetas:
                        resultados["creada_carpeta"].append(
                            {"region": region, "folio": folio, "carpeta": str(destino_dir)}
                        )
                        if not dry_run:
                            destino_dir.mkdir(parents=True, exist_ok=True)
                    else:
                        resultados["sin_carpeta"].append(
                            {"region": region, "folio": folio, "archivo": archivo.name,
                             "destino": str(destino_dir)}
                        )
                        continue

            destino_archivo = destino_dir / archivo.name

            if destino_archivo.exists():
                resultados["ya_existe"].append(
                    {"region": region, "folio": folio, "archivo": archivo.name}
                )
                continue

            resultados["copiado"].append(
                {"region": region, "folio": folio,
                 "archivo": archivo.name, "destino": str(destino_dir)}
            )

            if not dry_run:
                shutil.copy2(archivo, destino_archivo)

    # ── Reporte ───────────────────────────────────────────────────────────────
    modo = "DRY-RUN" if dry_run else "EJECUTADO"
    crear_txt = " + CREAR CARPETAS" if crear_carpetas else ""
    print(f"\n{'='*65}")
    print(f"  TRANSFERENCIA REORGANIZADO → ARCHIVOS  [{modo}{crear_txt}]")
    print(f"{'='*65}")

    print(f"\n✅  COPIADOS: {len(resultados['copiado'])}")
    por_region: dict[str, int] = {}
    for r in resultados["copiado"]:
        por_region[r["region"]] = por_region.get(r["region"], 0) + 1
    for reg, cnt in sorted(por_region.items()):
        print(f"      {reg}: {cnt}")

    if resultados["creada_carpeta"]:
        carpetas_unicas = {r["carpeta"] for r in resultados["creada_carpeta"]}
        print(f"\n📁  CARPETAS NUEVAS: {len(carpetas_unicas)}")
        por_region_c: dict[str, int] = {}
        vistas: set[str] = set()
        for r in resultados["creada_carpeta"]:
            if r["carpeta"] not in vistas:
                vistas.add(r["carpeta"])
                por_region_c[r["region"]] = por_region_c.get(r["region"], 0) + 1
        for reg, cnt in sorted(por_region_c.items()):
            print(f"      {reg}: {cnt}")

    print(f"\n⚠️   YA EXISTÍAN (no sobrescritos): {len(resultados['ya_existe'])}")
    if resultados["ya_existe"]:
        for r in resultados["ya_existe"][:5]:
            print(f"      [{r['region']}] {r['archivo']}")
        if len(resultados["ya_existe"]) > 5:
            print(f"      ... y {len(resultados['ya_existe'])-5} más")

    print(f"\n🔴  SIN CARPETA (folio en master, carpeta no existe): {len(resultados['sin_carpeta'])}")
    if resultados["sin_carpeta"]:
        for r in resultados["sin_carpeta"][:10]:
            print(f"      [{r['region']}] folio {r['folio']} — {r['archivo']}")
        if len(resultados["sin_carpeta"]) > 10:
            print(f"      ... y {len(resultados['sin_carpeta'])-10} más")
        if not crear_carpetas:
            print("      → Usa --crear-carpetas para crearlas automáticamente")

    print(f"\n🔸  FOLIO NO EN MASTER (omitido): {len(resultados['folio_no_master'])}")
    if resultados["folio_no_master"]:
        # Agrupar por región
        por_reg: dict[str, list] = {}
        for r in resultados["folio_no_master"]:
            por_reg.setdefault(r["region"], []).append(r["folio"])
        for reg, folios in sorted(por_reg.items()):
            unicos = sorted(set(folios))
            print(f"      [{reg}] folios: {', '.join(unicos)}")

    print(f"\n🔹  NOMBRE NO ESTÁNDAR (omitido): {len(resultados['nombre_invalido'])}")
    if resultados["nombre_invalido"]:
        for r in resultados["nombre_invalido"]:
            print(f"      [{r['region']}] {r['archivo']}")

    total_utiles = len(resultados["copiado"]) + len(resultados["ya_existe"])
    if resultados["creada_carpeta"]:
        total_utiles += len(resultados["creada_carpeta"])
    print(f"\n  Archivos útiles procesados: {total_utiles}")
    print()

    if dry_run:
        if not crear_carpetas:
            print("  → Para copiar + crear carpetas faltantes:")
            print("    python scripts/07_transferir_reorganizado.py --ejecutar --crear-carpetas")
        else:
            print("  → Para ejecutar:")
            print("    python scripts/07_transferir_reorganizado.py --ejecutar --crear-carpetas")
    print()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--ejecutar", action="store_true",
                        help="Ejecutar copia real (sin este flag: dry-run)")
    parser.add_argument("--crear-carpetas", action="store_true",
                        help="Crear carpetas faltantes para folios en master sin carpeta")
    args = parser.parse_args()
    transferir(dry_run=not args.ejecutar, crear_carpetas=args.crear_carpetas)
