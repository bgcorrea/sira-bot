"""
Normaliza ARCHIVOS-SUBIR según CRITERIOS.md.
Genera estructura: ARCHIVOS-NORMALIZADOS/{folio} - {RAZON_SOCIAL}/
  Convenio   - {folio}.pdf
  Resolución - {folio}.pdf
  Egreso     - {folio}.pdf
  Rendición  - {folio}.pdf
  Garantía   - {folio}.pdf  (si existe)

Uso:
  python scripts/B_normalizar.py           # dry-run (solo muestra)
  python scripts/B_normalizar.py --ejecutar # copia archivos
"""

import os
import re
import shutil
import sys
import openpyxl
from pathlib import Path

DRY_RUN = "--ejecutar" not in sys.argv

BASE      = Path(__file__).parent.parent
ARCHIVOS  = BASE / "ARCHIVOS-SUBIR"
SALIDA    = BASE / "ARCHIVOS-NORMALIZADOS"
LIBRO1    = BASE / "Libro1.xlsx"
LOG_CSV   = BASE / "logs" / "normalizacion_resultado.csv"

# ── Mapeo definitivo de CONVENIOS (archivo origen → folio) ───────────────────
# Solo el archivo definitivo/firmado por folio.
# Archivos marcados None = descartar (duplicados o fuera de Libro1).

CONVENIO_MAP = {
    # nombre exacto del archivo (en Convenios/) → folio
    "adrian escobar27-07-2023-150711.pdf":      77492,
    "amiga27-07-2023-160350.pdf":               None,   # terminado, no en Libro1
    "amor27-07-2023-150906.pdf":                74926,
    "bahia on line firmado10-08-2023-090855.pdf": 77482,
    "caleta higuerillas27-07-2023-150536.pdf":  74253,
    "calle larga27-07-2023-152316.pdf":         77550,
    "canal 227-07-2023-151820.pdf":             73739,
    "canal local27-07-2023-153246.pdf":         72589,
    "congreso27-07-2023-160141.pdf":            75758,
    "costa magazine27-07-2023-151211.pdf":      77176,
    "crystal27-07-2023-151727.pdf":             77412,
    "de la costa27-07-2023-151538.pdf":         77181,
    "eclipse27-07-2023-151417.pdf":             70905,
    "fabiola tello27-07-2023-161034.pdf":       77337,  # = Multitudes Periféricas
    "global cosmos27-07-2023-152658.pdf":       77821,
    "humberto lopez  (crecer)firmado10-08-2023-090802.pdf": 75516,
    "ilusion27-07-2023-153050.pdf":             72476,
    "interferencia ok27-07-2023-160310.pdf":    69805,
    "kudell27-07-2023-155130.pdf":              76841,
    "la merced27-07-2023-151003.pdf":           71801,
    "la quinta emprende27-07-2023-152906.pdf":  74148,
    "latina27-07-2023-155312.pdf":              73226,
    "lovely27-07-2023-160430.pdf":              77125,
    "mata o te rapa nui27-07-2023-160535.pdf":  73944,
    "montealegre27-07-2023-155604.pdf":         77380,
    "multitudes perifericas27-07-2023-160626.pdf": None,  # mismo folio que fabiola tello
    "pagina1227-07-2023-150815.pdf":            77305,
    "para dios27-07-2023-155423.pdf":           74649,
    "portales27-07-2023-155750.pdf":            77562,
    # preludio: solo los definitivos firmados
    "preludio fm ok27-07-2023-161004.pdf":      70236,  # Preludio Comunicaciones SPA (Radio FM)
    "preludio fm ok27-07-2023-160839.pdf":      None,   # versión anterior del mismo
    "preludio fm27-07-2023-160808.pdf":         None,   # borrador FM
    "preludio ok27-07-2023-160052.pdf":         70245,  # Preludio Comunicaciones E.I.R.L. (web)
    "preludio27-07-2023-160006.pdf":            None,   # borrador web
    "proa27-07-2023-153141.pdf":                76889,
    "quilpue on line27-07-2023-155838.pdf":     73951,
    "raudal27-07-2023-151302.pdf":              72170,
    "rd aconcagua27-07-2023-152419.pdf":        74989,
    "rd buena onda27-07-2023-152523.pdf":       77437,
    "rd he27-07-2023-155654.pdf":               71909,
    "rd quintay firmado10-08-2023-090959.pdf":  74637,
    "rd somos27-07-2023-152954.pdf":            77263,
    "rd valparaiso27-07-2023-155516.pdf":       77257,
    "somos la ligua27-07-2023-152119.pdf":      77798,
    "super andina27-07-2023-152217.pdf":        70793,
    "tu opinas27-07-2023-151053.pdf":           71206,
    "tv lanligua27-07-2023-155924.pdf":         70037,
    "vistamar27-07-2023-155223.pdf":            72699,
}

# ── Mapeo de EGRESOS (cuando el folio NO está en el nombre) ──────────────────
# Para archivos con folio en el nombre, se detecta automáticamente.
EGRESO_MANUAL = {
    "AGRUPACION PENIEL LA LIGUA.pdf":           74649,
    "BAHIA ONLINE.pdf":                         77482,  # duplicado de F 77482
    "JJVV Villa Sto Domingo.pdf":               73305,
    "RADIO CRECER FM.pdf":                      75516,  # duplicado de F75516
    "RADIO QUINTAY.pdf":                        74637,  # duplicado de F 74637
    "RADIO SOMOS.pdf":                          77798,  # duplicado de F 77798
    "RODRIGO BERNAL CACERES.pdf":               72589,  # duplicado de F 72589
    "F74253 RADIO COMUN. CALETA HUGUERI.pdf":   74253,
    "F75516 RADIO CRECER.pdf":                  75516,
}
# Egresos fuera de Libro1 (ignorar)
EGRESO_IGNORAR = {"77693 RADIO 88.7.pdf"}

# ── Mapeo de GARANTÍAS ───────────────────────────────────────────────────────
GARANTIA_MAP = {
    "88.7.pdf":                                         None,   # folio 77693 fuera de Libro1
    "AGRUACIÓN PENIEL LA LIGUA.pdf":                    74649,
    "AGRUPACION SOCIAL, CULTURAL Y COMUNITARIA.pdf":    69805,
    "BOX COMUNICACIONES SPA.pdf":                       77798,
    "CANAL LOCAL.pdf":                                  72589,
    "CENTRO CULTURAL REMA COMUNICACIONES DE EL QUISCO.pdf": None,  # no en Libro1
    "CENTRO JUVENIL CULTURAL SOCIAL Y DE COMUNICACIONES RATEM.pdf": 71876,
    "COMUNICACIONES ACONCAGUA SPA.pdf":                 77482,
    "COMUNICACIONES CONGRESO SPA.pdf":                  75758,
    "COMUNICACIONES JULIO HARDOY BAYLAUCQ.pdf":         None,   # no en Libro1
    "COMUNICACIONES PACIFICO SPA.pdf":                  77257,
    "ESTUDIO TV LA LIGUA.pdf":                          70037,
    "FABIOLA HERRERA LEIVA.pdf":                        None,   # no en Libro1
    "HERNAN PULGAR AGUILERA SPA.pdf":                   71909,
    "HUMBERTO LOPEZ VERGARA.pdf":                       75516,
    "JUNTA DE VECINOS CALETA HIGUERILLAS.pdf":          74253,
    "KUDELL TV.pdf":                                    76841,
    "LEONARDO PAKARATI.pdf":                            73542,
    "LOVELY FM.pdf":                                    77125,
    "MANUEL DIAZ VILLAGRAN.pdf":                        74919,
    "MIGUEL ANGEL JARA MALDONADO.pdf":                  None,   # no en Libro1
    "ORGANIZACION PARROQUIA NUESTRA SEÑORA DE LA MERCED.pdf": 71801,
    "PERIODICO DE LA COSTA.pdf":                        77181,
    "PRELUDIO COMUNICACIONES SPA.pdf":                  70236,
    "QUILPUE ONLINE.pdf":                               73951,
    "RADIO CALLE LARGA.pdf":                            77550,
    "RADIO ECLIPSE.pdf":                                70905,
    "RADIO PORTALES.pdf":                               77562,
    "RADIO RECREO SPA.pdf":                             77003,
    "RADIO SOMOS.pdf":                                  77798,
    "RADIO VISTAMAR.pdf":                               72699,
    "RADIODIFUSION ADRIAN ESCOBAR ARAVENA.pdf":         77492,
    "RADIODIFUSION VERONICA DEL CARMEN.pdf":            77821,
    "REVISTA COSTA MAGAZINE.pdf":                       77176,
    "SEMINARIO PAGINA 12.pdf":                          77305,
    "SOCIEDAD COMUNICACIONES RAUDAL LTDA.pdf":          72170,
    "SOCIEDAD DIFUSORA DE RADIO Y TV SAN ANTONI.pdf":   73739,
    "SOCIEDAD PUBLIEVENTOS LTDA.pdf":                   72009,
    "SOCIEDAD RADIO ACONCAGUA LIMITADA.pdf":            74989,
    "SUPERANDINA.pdf":                                  70793,
    "TU OPINAS.CL.pdf":                                 71206,
    "UNCO OLMUE.pdf":                                   None,   # no en Libro1 (UNION COMUNAL OLMUE no tiene garantía registrada)
    "VICTORIA CALDERÓN QUEVEDO.pdf":                    None,   # no en Libro1
}

# ── Resoluciones: acto administrativo → archivo de resolución ────────────────
RESOLUCION_ARCHIVO = {
    "1483":    "180-74.pdf",
    "1562":    "180-88.pdf",   # mismo que "2 APRUEBA CONVENIOS..."
    "272-074": "OFICIO 108.pdf",
    "272-083": "OFICIO 108.pdf",
    "272-097": "OFICIO 108.pdf",
    "272-100": "OFICIO 108.pdf",
}

# ── helpers ──────────────────────────────────────────────────────────────────

def sanitizar_nombre_carpeta(texto):
    """Reemplaza caracteres inválidos para nombres de carpeta en Windows/Linux."""
    invalidos = r'[<>:"/\\|?*]'
    return re.sub(invalidos, '-', texto).strip()


def extraer_folio_del_nombre(nombre):
    m = re.search(r'\b(\d{5,6})\b', nombre)
    if m:
        return int(m.group(1))
    m = re.search(r'[Ff](\d{5,6})', nombre)
    if m:
        return int(m.group(1))
    return None


def copiar(src, dst, dry_run):
    if src is None:
        return False
    src, dst = Path(src), Path(dst)
    if not src.exists():
        return False
    if not dry_run:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
    return True


# ── leer Libro1 ───────────────────────────────────────────────────────────────

def leer_libro1():
    wb = openpyxl.load_workbook(LIBRO1)
    ws = wb.active
    folios = {}
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        folio  = row[8]
        rut    = row[5]
        nombre = row[11]
        acto   = row[9]
        if folio:
            folios[int(folio)] = {
                "rut": str(rut),
                "nombre": str(nombre),
                "acto": str(acto) if acto else None,
            }
    return folios


# ── construir índices de archivos fuente ─────────────────────────────────────

def indice_convenios():
    """folio → ruta del archivo de convenio definitivo."""
    idx = {}
    carpeta = ARCHIVOS / "Convenios"
    for f in os.listdir(carpeta):
        folio = CONVENIO_MAP.get(f)
        if folio is not None:
            if folio not in idx:
                idx[folio] = carpeta / f
    return idx


def indice_egresos():
    """folio → ruta del egreso (prefiere archivos con folio explícito)."""
    idx = {}
    carpeta = ARCHIVOS / "EGRESOS"
    for f in sorted(os.listdir(carpeta)):
        if f in EGRESO_IGNORAR:
            continue
        # Intentar extraer folio del nombre
        folio = extraer_folio_del_nombre(Path(f).stem)
        # Si no, buscar en mapeo manual
        if folio is None:
            folio = EGRESO_MANUAL.get(f)
        if folio is None:
            continue
        # Solo sobreescribir si el nuevo archivo tiene folio explícito en el nombre
        # (preferir F NNNNN sobre nombres sin folio)
        tiene_folio_en_nombre = bool(extraer_folio_del_nombre(Path(f).stem))
        if folio not in idx or tiene_folio_en_nombre:
            idx[folio] = carpeta / f
    return idx


def indice_rendiciones():
    """folio → ruta de la rendición."""
    idx = {}
    carpeta = ARCHIVOS / "GARANTIAS Y RENDICIONES" / "02. RENDICIONES"
    ignorar = {72374, 77693}
    for f in os.listdir(carpeta):
        folio = extraer_folio_del_nombre(Path(f).stem)
        if folio and folio not in ignorar:
            idx[folio] = carpeta / f
    return idx


def indice_garantias():
    """folio → ruta de la garantía."""
    idx = {}
    carpeta = ARCHIVOS / "GARANTIAS Y RENDICIONES" / "01. GARANTÍAS"
    for f in os.listdir(carpeta):
        folio = GARANTIA_MAP.get(f)
        if folio is not None and folio not in idx:
            idx[folio] = carpeta / f
    return idx


def indice_resoluciones():
    """acto (str o int) → ruta del archivo de resolución."""
    carpeta = ARCHIVOS / "RESOLUCIONES"
    idx = {}
    for acto, nombre_archivo in RESOLUCION_ARCHIVO.items():
        if nombre_archivo:
            ruta = carpeta / nombre_archivo
            if ruta.exists():
                idx[str(acto)] = ruta
    return idx


# ── normalización principal ───────────────────────────────────────────────────

def normalizar(folios_libro1, dry_run):
    conv  = indice_convenios()
    egr   = indice_egresos()
    rend  = indice_rendiciones()
    garan = indice_garantias()
    resol = indice_resoluciones()

    print("=" * 70)
    modo = "DRY-RUN (simulación)" if dry_run else "EJECUTANDO (copiando archivos)"
    print(f"  Modo: {modo}")
    print(f"  Salida: {SALIDA}")
    print("=" * 70)

    filas_log = []
    resumen = {"completos": 0, "parciales": 0, "sin_documentos": 0}

    for folio, info in sorted(folios_libro1.items()):
        nombre     = info["nombre"]
        acto_raw   = info["acto"]

        # Determinar clave de resolución
        try:
            acto_key = str(int(float(acto_raw))) if acto_raw else None
        except (ValueError, TypeError):
            acto_key = str(acto_raw) if acto_raw else None

        # Determinar si hay resolución disponible
        resol_src = resol.get(acto_key) if acto_key else None

        # ¿Tiene algún documento (incluyendo resolución)?
        tiene_algo = any([
            folio in conv, folio in egr,
            folio in rend, folio in garan,
            resol_src is not None,
        ])
        if not tiene_algo:
            resumen["sin_documentos"] += 1
            filas_log.append({
                "folio": folio, "razon_social": nombre,
                "convenio": "", "resolucion": "", "egreso": "",
                "rendicion": "", "garantia": "", "estado": "SIN DOCUMENTOS"
            })
            continue

        # Crear carpeta destino
        nombre_carpeta = sanitizar_nombre_carpeta(nombre)
        carpeta_destino = SALIDA / f"{folio} - {nombre_carpeta}"
        if not dry_run:
            carpeta_destino.mkdir(parents=True, exist_ok=True)

        estado_docs = {}

        # Convenio
        src = conv.get(folio)
        dst = carpeta_destino / f"Convenio - {folio}.pdf"
        estado_docs["convenio"] = "OK" if copiar(src, dst, dry_run) else "FALTA"

        # Resolución
        if resol_src:
            dst_res = carpeta_destino / f"Resolución - {folio}.pdf"
            ok = copiar(resol_src, dst_res, dry_run)
            estado_docs["resolucion"] = "OK" if ok else "FALTA"
        else:
            estado_docs["resolucion"] = "SIN ARCHIVO"

        # Egreso
        src = egr.get(folio)
        dst = carpeta_destino / f"Egreso - {folio}.pdf"
        estado_docs["egreso"] = "OK" if copiar(src, dst, dry_run) else "FALTA"

        # Rendición
        src = rend.get(folio)
        dst = carpeta_destino / f"Rendición - {folio}.pdf"
        estado_docs["rendicion"] = "OK" if copiar(src, dst, dry_run) else "FALTA"

        # Garantía (opcional)
        src = garan.get(folio)
        if src:
            dst = carpeta_destino / f"Garantía - {folio}.pdf"
            estado_docs["garantia"] = "OK" if copiar(src, dst, dry_run) else "FALTA"
        else:
            estado_docs["garantia"] = "-"

        # Estado global
        obligatorios = [estado_docs["convenio"], estado_docs["resolucion"],
                        estado_docs["egreso"], estado_docs["rendicion"]]
        if all(v == "OK" for v in obligatorios):
            estado_global = "COMPLETO"
            resumen["completos"] += 1
        else:
            faltantes = [k for k, v in estado_docs.items() if v == "FALTA"]
            sin_arch  = [k for k, v in estado_docs.items() if v == "SIN ARCHIVO"]
            partes = []
            if faltantes:
                partes.append("FALTA: " + ",".join(faltantes))
            if sin_arch:
                partes.append("SIN ARCHIVO: " + ",".join(sin_arch))
            estado_global = " | ".join(partes) if partes else "INCOMPLETO"
            resumen["parciales"] += 1

        # Imprimir
        icono = "✓" if estado_global == "COMPLETO" else "!"
        print(f"[{icono}] {folio} - {nombre[:45]:<45}")
        if estado_global != "COMPLETO":
            print(f"    Conv:{estado_docs['convenio']} Resol:{estado_docs['resolucion']} "
                  f"Egr:{estado_docs['egreso']} Rend:{estado_docs['rendicion']} "
                  f"Gar:{estado_docs['garantia']}")

        filas_log.append({
            "folio": folio, "razon_social": nombre,
            "convenio": estado_docs["convenio"],
            "resolucion": estado_docs["resolucion"],
            "egreso": estado_docs["egreso"],
            "rendicion": estado_docs["rendicion"],
            "garantia": estado_docs["garantia"],
            "estado": estado_global,
        })

    # Resumen final
    print()
    print("=" * 70)
    print(f"COMPLETOS    : {resumen['completos']}")
    print(f"PARCIALES    : {resumen['parciales']}")
    print(f"SIN DOCUMENTOS: {resumen['sin_documentos']}")
    print("=" * 70)

    # Guardar log CSV
    import csv
    with open(LOG_CSV, "w", newline="", encoding="utf-8") as f:
        campos = ["folio", "razon_social", "convenio", "resolucion",
                  "egreso", "rendicion", "garantia", "estado"]
        writer = csv.DictWriter(f, fieldnames=campos)
        writer.writeheader()
        writer.writerows(filas_log)
    print(f"\nLog guardado en: {LOG_CSV}")

    if dry_run:
        print("\nPara ejecutar la copia real: python scripts/B_normalizar.py --ejecutar")


# ── main ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    folios = leer_libro1()
    normalizar(folios, DRY_RUN)
