"""
Normaliza FFOIP hacia ARCHIVOS-NORMALIZADOS para los 41 folios sociales.
Reemplaza el OFICIO 108 provisional con las resoluciones reales de FFOIP.

Uso:
  python scripts/C_normalizar_ffoip.py           # dry-run
  python scripts/C_normalizar_ffoip.py --ejecutar
"""

import os, re, shutil, sys
import openpyxl
from pathlib import Path

DRY_RUN = "--ejecutar" not in sys.argv

BASE  = Path(__file__).parent.parent
FFOIP = BASE / "FFOIP"
NORM  = BASE / "ARCHIVOS-NORMALIZADOS"

# ── Mapeo EGRESOS (por nombre de archivo) ────────────────────────────────────
EGRESO_MAP = {
    "ACCION INTEGRA.pdf":                           75164,
    "AGRUPACION COMUNITARIA PURA VIDA.pdf":         73021,
    "AGRUPACION TERCERA EDAD.pdf":                  75192,
    "AGRUPACIONES DE ADULTO MAYOR DE PLACILLA.pdf": 72057,
    "ALMA JOVEN.pdf":                               75064,
    "APODERADOS Y AMIGOS EL GRUPO.pdf":             76761,
    "ASOC PADRES Y AMIGOS DE LOS.pdf":              72964,
    "ASOCIACION DEPORTIVA PEDRO AGUIRRE CERDA.pdf": 72537,
    "CHILE WANDERERS.pdf":                          76437,
    "CLUB ADULTO MAYOR ARMONIA.pdf":                76284,
    "CLUB ADULTO MAYOR SONRISA.pdf":                72754,
    "CLUB DE ADULTO MAYOR ACTIVA Y JOV.pdf":        73834,
    "CLUB DEPORTIVO DE BASQUETBOL.pdf":             72112,
    "CLUB DEPORTIVO VILLARRICA.pdf":                70373,
    "CLUB SOCIAL Y DEPORTIVO OROMPE.pdf":           73166,
    "CON CON NATIONAL.pdf":                         73504,
    "CORPORACION MAYOR VIDA.pdf":                   70703,
    "CORPORACION PILARES.pdf":                      70311,
    "EL CHAGUAL.pdf":                               69928,
    "FUNDACION CADENA DE FAVORES.pdf":              75822,
    "FUNDACION EDUCACIONAL CRUZ.pdf":               72936,
    "FUNDACION IMPULSA 21.pdf":                     69453,
    "FUNDACION KURMI.pdf":                          75082,
    "FUNDACION PAPUDO.pdf":                         70196,
    "FUNDACION PUNTO ESPORA.pdf":                   74567,
    "FUNDACION SALUD.pdf":                          70863,
    "FUNDACION TRUEKE.pdf":                         75127,
    "GIRASOL.pdf":                                  74631,
    "JJVV JOSE MIGUEL CARRERA.pdf":                 76171,
    "JJVV Villa Sto Domingo.pdf":                   73305,
    "JUNTA DE VECINOS EL MIRADOR.pdf":              74903,
    "JUNTA DE VECINOS N11 PHILLIPI.pdf":            75772,
    "JUNTA DE VECINOS RECREO.pdf":                  75284,
    "NEW CRUSADERS.pdf":                            75741,
    "ONG DE DESARROLLO AMUN.pdf":                   73472,
    "ORGANIZACION NO GUBERNAMENTAL 65060696-5.pdf": 74874,
    "ORGANIZACION NO GUBERNAMENTAL 65167392-5.pdf": 72159,
    "SOMOS HUMEDAL CORDOVA.pdf":                    70641,
    "TIERRAS ROJAS.pdf":                            75606,
    "TRAMA RURAL.pdf":                              75331,
    "UNION COMUNAL DE JUNTAS DE VECINOS.pdf":       76359,
}

# ── Mapeo GARANTÍAS (por nombre, opcional) ───────────────────────────────────
GARANTIA_MAP = {
    "AG. DE CLUBES DE ADULTO MAYOR DE PLACILLA.pdf":                    72057,
    "AG.COMUNITARIA PURA VIDA MADREBEBE.pdf":                           73021,
    "AG.TERCERA EDAD Y PACIENTES CRONICOS VERDE ESPERANZA.pdf":         75192,
    "AMIGOS DEL GRUPO SCOUT CAPITAN AQUILES RAMIREZ.pdf":               76761,
    "AS.DE FUTBOL PEDRO AGUIRRE CERDA.pdf":                             72537,
    "AS.DE PADRES Y AMIGOS DE LOS AUTISTAS V REGION.pdf":               72964,
    "CASA CULTURAL EL CHAGUAL.pdf":                                     69928,
    "CLUB ADULTO MAYOR ACTIVA Y JOVIALES.pdf":                          73834,
    "CLUB ADULTO MAYOR SONRISA Y CORAZONES.pdf":                        72754,
    "CLUB DE ADULTO MAYOR ARMONIA.pdf":                                 76284,
    "CLUB DEPORTIVO CHILE WANDERERS.pdf":                               76437,
    "CLUB DEPORTIVO DE BASQUETBO SENIORS DAMAS MINERVA.pdf":            72112,
    "CLUB DEPORTIVO NEW CRUSADERS.pdf":                                 75741,
    "CLUB SOCIAL Y DEPORTIVO OROMPELLO.pdf":                            73166,
    "COMITE TIERRAS ROJAS LAGUNA VERDE.pdf":                            75606,
    "CONCON NATIONAL.pdf":                                              73504,
    "CORPORACION MAYOR VIDA.pdf":                                       70703,
    "CORPORACION PILARES.pdf":                                          70311,
    "FUNDACION CADENA DE FAVORES.pdf":                                  75822,
    "FUNDACION EDUCACIONAL CRUZ DEL SUR.pdf":                           72936,
    "FUNDACION IMPULSA 21.pdf":                                         69453,
    "FUNDACION PAPUDO.pdf":                                             70196,
    "FUNDACION TRAMA RURAL.pdf":                                        75331,
    "FUNDACION TRUEKE.pdf":                                             75127,
    "GRUPO FOCLORICO ALMA JOVEN.pdf":                                   75064,
    "ORG.GRUPO FOLCLORICO ALMA JOVEN.pdf":                              None,  # duplicado
    "JJVV JOSE MIGUEL CARRERA.pdf":                                     76171,
    "JJVV NR.11 PHILLIPI.pdf":                                          75772,
    "JJVV RECREO Y RODRIGUEZ.pdf":                                      75284,
    "JJVV VILLA EL MIRADOR.pdf":                                        74903,
    "ORG.NO GUBERNAMENTAL DE DESARROLLO AMUN.pdf":                      73472,
    "ORG.NO GUBERNAMENTAL DE DESARROLLO COMUNIDAD LA ESCUELA IMAGINARIA.pdf": 72159,
    "ORG.NO GUBERNAMENTAL DE DESARROLLO LA MATRIZ.pdf":                 74874,
    "ORG.PROYECTO GIRASOL.pdf":                                         74631,
    "UNCO OLMUE.pdf":                                                   76359,
}

# ── Resoluciones: acto → archivo FFOIP ───────────────────────────────────────
RESOLUCION_FFOIP = {
    "272-074": "188-71 1TT FFOIP Valparaíso.pdf",
    "272-083": "188-87.pdf",
    "272-097": "certificsdo25-07-2023-103724.pdf",
    "272-100": "OFICIO 110.pdf",
}

# ── helpers ───────────────────────────────────────────────────────────────────

def extraer_folio(nombre):
    m = re.match(r'^(\d{5,6})', nombre)
    return int(m.group(1)) if m else None

def copiar(src, dst):
    src, dst = Path(src), Path(dst)
    if not src.exists():
        return False
    if not DRY_RUN:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
    return True

def carpeta_destino(folio, nombre):
    nombre_safe = re.sub(r'[<>:"/\\|?*]', '-', nombre).strip()
    return NORM / f"{folio} - {nombre_safe}"

# ── leer Libro1 ───────────────────────────────────────────────────────────────

def leer_libro1():
    wb = openpyxl.load_workbook(BASE / "Libro1.xlsx")
    ws = wb.active
    folios = {}
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0: continue
        if row[8]:
            folios[int(row[8])] = {"nombre": str(row[11]), "acto": str(row[9]) if row[9] else None}
    return folios

# ── indices de archivos FFOIP ─────────────────────────────────────────────────

def idx_convenios():
    idx = {}
    for f in (FFOIP / "CONVENIOS").glob("*.pdf"):
        if f.name.startswith("BORRAR"):
            continue
        folio = extraer_folio(f.name)
        if folio:
            idx[folio] = f
    return idx

def idx_egresos():
    idx = {}
    for f in (FFOIP / "EGRESOS").glob("*.pdf"):
        folio = EGRESO_MAP.get(f.name)
        if folio:
            idx[folio] = f
    return idx

def idx_rendiciones():
    idx = {}
    for f in (FFOIP / "GARANTIAS Y RENDICIONES" / "02. RENDICIÓN").glob("*.pdf"):
        folio = extraer_folio(f.name)
        if folio:
            idx[folio] = f
    return idx

def idx_garantias():
    idx = {}
    for f in (FFOIP / "GARANTIAS Y RENDICIONES" / "01. GARANTÍAS").glob("*.pdf"):
        folio = GARANTIA_MAP.get(f.name)
        if folio and folio not in idx:
            idx[folio] = f
    return idx

def idx_resoluciones():
    idx = {}
    for acto, nombre in RESOLUCION_FFOIP.items():
        ruta = FFOIP / "RESOLUCIONES" / nombre
        if ruta.exists():
            idx[acto] = ruta
    return idx

# ── normalización ─────────────────────────────────────────────────────────────

def normalizar(folios_libro1):
    conv  = idx_convenios()
    egr   = idx_egresos()
    rend  = idx_rendiciones()
    garan = idx_garantias()
    resol = idx_resoluciones()

    FOLIOS_SOCIALES = set(EGRESO_MAP.values()) | {70373, 70863, 73305}

    print("=" * 65)
    print(f"  Modo: {'DRY-RUN' if DRY_RUN else 'EJECUTANDO'}")
    print("=" * 65)

    completos = parciales = 0

    for folio in sorted(FOLIOS_SOCIALES):
        info = folios_libro1.get(folio, {})
        nombre = info.get("nombre", f"FOLIO {folio}")
        acto   = info.get("acto")
        dst    = carpeta_destino(folio, nombre)

        estado = {}

        # Convenio
        estado["Conv"] = "OK" if copiar(conv.get(folio), dst / f"Convenio - {folio}.pdf") else "FALTA"

        # Resolución (reemplaza OFICIO 108 anterior)
        resol_src = resol.get(acto) if acto else None
        estado["Resol"] = "OK" if copiar(resol_src, dst / f"Resolución - {folio}.pdf") else "FALTA"

        # Egreso
        estado["Egr"] = "OK" if copiar(egr.get(folio), dst / f"Egreso - {folio}.pdf") else "FALTA"

        # Rendición
        estado["Rend"] = "OK" if copiar(rend.get(folio), dst / f"Rendición - {folio}.pdf") else "FALTA"

        # Garantía (opcional)
        if folio in garan:
            estado["Gar"] = "OK" if copiar(garan[folio], dst / f"Garantía - {folio}.pdf") else "FALTA"
        else:
            estado["Gar"] = "-"

        obligatorios = [estado["Conv"], estado["Resol"], estado["Egr"], estado["Rend"]]
        ok = all(v == "OK" for v in obligatorios)
        if ok:
            completos += 1
            print(f"[✓] {folio} - {nombre[:50]}")
        else:
            parciales += 1
            detalle = "  ".join(f"{k}:{v}" for k, v in estado.items())
            print(f"[!] {folio} - {nombre[:45]}")
            print(f"    {detalle}")

    print()
    print("=" * 65)
    print(f"COMPLETOS : {completos}/41")
    print(f"PARCIALES : {parciales}/41")
    print("=" * 65)
    if DRY_RUN:
        print("\nEjecutar: python scripts/C_normalizar_ffoip.py --ejecutar")

if __name__ == "__main__":
    folios = leer_libro1()
    normalizar(folios)
