"""
Script 08 — Distribuir resoluciones a carpetas por folio.

Lógica:
- Lee DISTRIBUCIÓN CARGA VB.xlsx: folio → N° Acto Administrativo (número de resolución)
- Complementa con asignaciones manuales para 6 folios no presentes en el Excel
- Para cada folio: copia el PDF de resolución regional a su carpeta de destino
  como "Resolución - {folio}.pdf"
- La misma resolución se copia una vez por cada folio que la referencia

Uso:
    python scripts/08_distribuir_resoluciones.py             # dry-run
    python scripts/08_distribuir_resoluciones.py --ejecutar  # copia real
"""

import argparse
import shutil
from pathlib import Path

import pandas as pd

MASTER   = Path("data/master_subida.xlsx")
EXCEL    = Path("data/DISTRIBUCIÓN CARGA VB.xlsx")

# ── Mapeo (región, acto_num) → PDF source ────────────────────────────────────
# Ruta relativa a la carpeta de la región en FFOIP 2022 (REGIONES)
# La clave acto_num es exactamente como aparece en el Excel (columna N° Acto Administrativo)

REGION_FOLDER = {
    "ARICA":      "03. Arica y Parinacota",
    "ATACAMA":    "06. Atacama",
    "VALPARAÍSO": "08. Valparaíso",
    "VALAPARAISO": "08. Valparaíso",   # typo en Excel
    "O'HIGGINS":  "09. O'Higgins",
    "ÑUBLE":      "11. Ñuble",
    "AYSEN":      "16. Aysén",
    "RM":         "18. RM",
}

# Carpeta base en Archivos/ para cada región (para crear carpetas faltantes)
ARCHIVOS = BASE / "Archivos"
ARCHIVOS_BASE = {
    "ARICA":      "01. ARICA - FFOIP",
    "ATACAMA":    "04. ATACAMA - FFOIP",
    "VALPARAÍSO": "06. VALPARAÍSO - FFOIP",
    "VALAPARAISO": "06. VALPARAÍSO - FFOIP",
    "O'HIGGINS":  "08. O'HIGGINS - FFOIP",
    "ÑUBLE":      "10. ÑUBLE - FFOIP",
    "AYSEN":      "15. AYSEN - FFOIP",
    "RM":         "07. METROPOLITANA - FFOIP/07. Firma de Convenios",
}

RES_PDF: dict[tuple[str, str], str] = {
    # ARICA
    ("ARICA", "106/2022"): "08. Res. aprueba convenio (Res+ Memo TT)/Rex 106 Aprueba convenios.pdf",
    ("ARICA", "130/2022"): "08. Res. aprueba convenio (Res+ Memo TT)/REX. 130 Aprueba convenios JV NACE CHILE Y JV PORTAL DEL SOL.pdf",
    ("ARICA", "141/2022"): "08. Res. aprueba convenio (Res+ Memo TT)/Res. Exenta N°141.pdf",
    ("ARICA", "155/2022"): "08. Res. aprueba convenio (Res+ Memo TT)/REX 155.pdf",

    # ATACAMA
    ("ATACAMA", "72"): "08. Res. aprueba convenio (Res+ Memo TT)/Resolución y memo aprueba convenio22-07-2022-141611.pdf",
    ("ATACAMA", "74"): "/home/bgcorrea/personal/workspace/caigg/res faltantes/Resolución Exenta Nro 7407-11-2022-145422.pdf",
    ("ATACAMA", "79"): "/home/bgcorrea/personal/workspace/caigg/res faltantes/Resolución Exenta Nro 7907-11-2022-145610.pdf",
    ("ATACAMA", "80"): "/home/bgcorrea/personal/workspace/caigg/res faltantes/Resolución Exenta Nro 8007-11-2022-145722.pdf",

    # VALPARAÍSO (incluye ambas variantes del nombre de región del Excel)
    ("VALPARAÍSO", "272/109"):  "08. Res. aprueba convenio (Res+ Memo TT)/RES Y OFICIO 1 (28 PROYECTOS)/RES.APRUEBA CONVENIOS 272 109.pdf",
    ("VALPARAÍSO", "272/110"):  "08. Res. aprueba convenio (Res+ Memo TT)/RES Y OFICIO 2/res 272-110 (1) (2).pdf",
    ("VALPARAÍSO", "272/118"):  "09. Readjudicación- Renuncias/docs firmados/of 117 - res. ex272-118.pdf",
    ("VALPARAÍSO", "272/163"):  "/home/bgcorrea/personal/workspace/caigg/res faltantes/oficio 133_y res272-163.pdf",
    ("VALPARAÍSO", "272/164"):  "/home/bgcorrea/personal/workspace/caigg/res faltantes/oficio 135 Y res 272-164.pdf",
    ("VALPARAÍSO", "272/172"):  "08. Res. aprueba convenio (Res+ Memo TT)/resolucion 17223-09-2022-155043.pdf",
    ("VALAPARAISO", "272/109"): "08. Res. aprueba convenio (Res+ Memo TT)/RES Y OFICIO 1 (28 PROYECTOS)/RES.APRUEBA CONVENIOS 272 109.pdf",

    # O'HIGGINS (113-2022 → sin PDF)
    ("O'HIGGINS", "073-2022"): "08. Res. aprueba convenio (Res+ Memo TT)/RES EX 073-2022 APRUEBA CONVENIOS LOCALES Y PROVINCIALES FFOIP 2022 OHIGGINS.pdf",
    ("O'HIGGINS", "082-2022"): "08. Res. aprueba convenio (Res+ Memo TT)/RES EX 082-2022 APRUEBA 2 CONVENIOS LOCALES FFOIP 2022.pdf",
    ("O'HIGGINS", "084-2022"): "08. Res. aprueba convenio (Res+ Memo TT)/ORD 128 Y RES EX 84-2022 SOLICITA RECURSOS PROYECTO LOCAL FFOIP.pdf",

    # ÑUBLE
    ("ÑUBLE", "88"):  "08. Res. aprueba convenio (Res+ Memo TT)/1. Res. N° 088 Aprueba Convenios Locales y Regionales FFOIP 2022.pdf",
    ("ÑUBLE", "94"):  "08. Res. aprueba convenio (Res+ Memo TT)/2. Res. N° 094 Aprueba Convenios Locales y Regionales FFOIP 2022.pdf",
    ("ÑUBLE", "111"): "08. Res. aprueba convenio (Res+ Memo TT)/3. RES EX. N° 111 Aprueba Convenios  Locales.pdf",
    ("ÑUBLE", "135"): "08. Res. aprueba convenio (Res+ Memo TT)/4. Res. Ex N° 135 Aprueba Convenio Regional FFOIP 2022.pdf",
    ("ÑUBLE", "175"): "08. Res. aprueba convenio (Res+ Memo TT)/5. Res. Ex N° 175 Aprueba Convenio Regional FFOIP 2022.pdf",

    # AYSÉN
    ("AYSEN", "97"): "08. Res. aprueba convenio (Res+ Memo TT)/Res. 97 aprueba convenio FFOIP 2022.pdf",

    # RM
    ("RM", "272/714"): "08. Res. aprueba convenio (Res+ Memo TT)/Res. Aprueba Convenios (Locales y Regionales)/272-714.pdf",
    ("RM", "272/715"): "08. Res. aprueba convenio (Res+ Memo TT)/Res. Aprueba Convenios (Nacionales)/272-715.pdf",
}

# Folios sin entrada en el Excel → asignación manual
MANUAL: dict[str, tuple[str, str]] = {
    # folio: (region_key, acto_num)
    "54445": ("RM",    "272/714"),
    "54480": ("RM",    "272/714"),
    "54706": ("RM",    "272/714"),
    "54743": ("RM",    "272/714"),
    "54622": ("ÑUBLE", "111"),
    "54872": ("ÑUBLE", "135"),
}


def cargar_datos() -> tuple[dict[str, tuple[str, str]], dict[str, Path], dict[str, dict]]:
    """
    Devuelve:
      asig: {folio: (region_key, acto_num)}
      folio_map: {folio: carpeta_folio_path}   (solo los que tienen carpeta)
      master_all: {folio: {'region', 'razon_social'}}  (todos los 250)
    """
    excel = pd.read_excel(EXCEL)
    excel["folio"]    = excel["ID Convenio"].astype(str).str.strip()
    excel["region"]   = excel["REGION"].astype(str).str.strip().str.upper()
    excel["acto_num"] = excel["N° Acto Administrativo"].astype(str).str.strip()

    asig: dict[str, tuple[str, str]] = {}
    for _, row in excel.iterrows():
        asig[row["folio"]] = (row["region"], row["acto_num"])

    asig.update(MANUAL)

    master_df = pd.read_excel(MASTER)
    folio_map: dict[str, Path] = {}
    master_all: dict[str, dict] = {}
    for _, row in master_df.iterrows():
        folio = str(row["folio"])
        region = str(row["region"]).strip().upper()
        razon = str(row["razon_social"]) if pd.notna(row["razon_social"]) else ""
        master_all[folio] = {"region": region, "razon_social": razon}
        if pd.notna(row["carpeta_folio"]) and str(row["carpeta_folio"]).strip():
            folio_map[folio] = Path(str(row["carpeta_folio"]))

    return asig, folio_map, master_all


def resolver_pdf(region_key: str, acto_num: str) -> Path | None:
    """Devuelve el path absoluto del PDF de resolución, o None si no está mapeado."""
    key = (region_key, acto_num)
    val = RES_PDF.get(key)
    if val is None:
        return None
    p = Path(val)
    if p.is_absolute():
        return p
    folder_name = REGION_FOLDER.get(region_key)
    if folder_name is None:
        return None
    return REGIONES / folder_name / val


def distribuir(dry_run: bool, crear_carpetas: bool) -> None:
    asig, folio_map, master_all = cargar_datos()

    resultados = {
        "copiado":        [],
        "ya_existe":      [],
        "sin_pdf":        [],
        "sin_carpeta":    [],   # tiene PDF pero no hay carpeta destino y no se pidió crear
        "creada_carpeta": [],   # carpeta creada por --crear-carpetas
        "sin_asignacion": [],
    }

    todos_folios = list(master_all.keys())

    for folio in todos_folios:
        if folio not in asig:
            resultados["sin_asignacion"].append({"folio": folio})
            continue

        region_key, acto_num = asig[folio]
        pdf_src = resolver_pdf(region_key, acto_num)

        if pdf_src is None or not pdf_src.exists():
            resultados["sin_pdf"].append({
                "folio": folio, "region": region_key, "acto_num": acto_num,
            })
            continue

        if folio not in folio_map:
            info = master_all[folio]
            base_key = info["region"]
            archivos_base = ARCHIVOS_BASE.get(base_key) or ARCHIVOS_BASE.get(region_key)
            if archivos_base and crear_carpetas:
                nueva = ARCHIVOS / archivos_base / f"{folio} - {info['razon_social']}"
                resultados["creada_carpeta"].append({"folio": folio, "region": base_key, "carpeta": str(nueva)})
                if not dry_run:
                    nueva.mkdir(parents=True, exist_ok=True)
                folio_map[folio] = nueva  # usar en el mismo pase
            else:
                resultados["sin_carpeta"].append({"folio": folio, "region": base_key or region_key})
                continue

        destino_dir  = folio_map[folio]
        destino_file = destino_dir / f"Resolución - {folio}.pdf"

        if destino_file.exists():
            resultados["ya_existe"].append({"folio": folio, "region": region_key})
            continue

        resultados["copiado"].append({
            "folio": folio, "region": region_key,
            "src": pdf_src.name, "destino": str(destino_dir)
        })

        if not dry_run:
            shutil.copy2(pdf_src, destino_file)

    # ── Reporte ───────────────────────────────────────────────────────────────
    modo = "DRY-RUN" if dry_run else "EJECUTADO"
    crear_txt = " + CREAR CARPETAS" if crear_carpetas else ""
    print(f"\n{'='*62}")
    print(f"  DISTRIBUCIÓN DE RESOLUCIONES  [{modo}{crear_txt}]")
    print(f"{'='*62}")

    print(f"\n✅  COPIADOS: {len(resultados['copiado'])}")
    por_reg: dict[str, int] = {}
    for r in resultados["copiado"]:
        por_reg[r["region"]] = por_reg.get(r["region"], 0) + 1
    for reg, cnt in sorted(por_reg.items()):
        print(f"      {reg}: {cnt}")

    if resultados["creada_carpeta"]:
        print(f"\n📁  CARPETAS NUEVAS: {len(resultados['creada_carpeta'])}")
        for r in resultados["creada_carpeta"]:
            print(f"      [{r['region']}] folio {r['folio']}")

    print(f"\n⚠️   YA EXISTÍAN: {len(resultados['ya_existe'])}")
    if resultados["ya_existe"]:
        for r in resultados["ya_existe"][:5]:
            print(f"      [{r['region']}] folio {r['folio']}")
        if len(resultados["ya_existe"]) > 5:
            print(f"      ... y {len(resultados['ya_existe'])-5} más")

    print(f"\n🔴  SIN PDF DISPONIBLE: {len(resultados['sin_pdf'])}")
    if resultados["sin_pdf"]:
        prev_reg = None
        for r in sorted(resultados["sin_pdf"], key=lambda x: (x["region"], x["acto_num"])):
            if r["region"] != prev_reg:
                print(f"      [{r['region']}]")
                prev_reg = r["region"]
            print(f"        folio {r['folio']} — resolución {r['acto_num']}")

    print(f"\n🔸  SIN CARPETA DESTINO: {len(resultados['sin_carpeta'])}")
    if resultados["sin_carpeta"]:
        for r in resultados["sin_carpeta"]:
            print(f"      [{r['region']}] folio {r['folio']}")
        if not crear_carpetas:
            print("      → Usa --crear-carpetas para crearlas automáticamente")

    print(f"\n🔹  SIN ASIGNACIÓN EN EXCEL/MANUAL: {len(resultados['sin_asignacion'])}")
    if resultados["sin_asignacion"]:
        for r in resultados["sin_asignacion"]:
            print(f"      folio {r['folio']}")

    total = len(resultados["copiado"]) + len(resultados["ya_existe"])
    print(f"\n  Resoluciones procesables: {total} / {len(todos_folios)}")
    print()
    if dry_run:
        if not crear_carpetas:
            print("  → Para ejecutar con creación de carpetas faltantes:")
            print("    python scripts/08_distribuir_resoluciones.py --ejecutar --crear-carpetas")
        else:
            print("  → Para ejecutar:")
            print("    python scripts/08_distribuir_resoluciones.py --ejecutar --crear-carpetas")
    print()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--ejecutar", action="store_true")
    parser.add_argument("--crear-carpetas", action="store_true",
                        help="Crear carpetas para folios en master sin carpeta destino")
    args = parser.parse_args()
    distribuir(dry_run=not args.ejecutar, crear_carpetas=args.crear_carpetas)
