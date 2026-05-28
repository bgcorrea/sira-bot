"""
Script 01: Descubrir archivos en disco
=======================================
Escanea las carpetas locales donde están los documentos respaldados y
genera un CSV con TODOS los PDFs encontrados, intentando extraer el folio
de cada uno cuando es posible.

NO hace matching folio→documento todavía. Solo inventario crudo.
Eso lo hace el script 03_armar_master.py.

Uso:
    python scripts/01_descubrir_archivos.py
"""

import os
import re
import csv
from pathlib import Path

# ====== CONFIGURACIÓN ======
BASE_DIR = "/home/bgcorrea/personal/workspace/caigg"
RAIZ_RESPALDOS = BASE_DIR + "/Archivos"
RAIZ_CONVENIOS = BASE_DIR + "/Archivos"
RAIZ_REGISTRO_ESTADO = BASE_DIR + "/Colaboradores del Estado"
OUTPUT_CSV = "logs/inventario_disco.csv"

# Solo nos interesan PDFs e imágenes
EXTENSIONES_VALIDAS = {".pdf", ".jpg", ".jpeg", ".png"}

# Patrón para extraer folios (números de 5 dígitos típicamente)
PATRON_FOLIO = re.compile(r"\b(5[0-9]{4}|6[0-9]{4}|7[0-9]{4}|8[0-9]{4}|9[0-9]{4}|1[0-9]{5})\b")

# Categorías por palabras clave en la ruta
CATEGORIAS = {
    "convenio": ["convenio", "firma de convenios", "12. convenio", "11. convenio", "convenios escaneados"],
    "acto_administrativo": ["resolución", "resolucion", "res. aprueba", "res aprueba", "rex"],
    "rendicion": ["rendición", "rendicion", "rendiciones", "cierre", "cfc", "respaldo de rendicion"],
    "garantia": ["garantía", "garantias", "garantías", "letra de cambio", "letras de cambio"],
    "voucher": ["egresos", "transferencia", "recepción de recursos", "recepcion de recursos", "certificado bancario"],
    "certificado_registro": ["registro colaboradores"],
}


def categorizar(ruta_relativa: str) -> str:
    """Determina la categoría del archivo según su ruta."""
    ruta_lower = ruta_relativa.lower()
    for categoria, keywords in CATEGORIAS.items():
        for kw in keywords:
            if kw in ruta_lower:
                return categoria
    return "otro"


def extraer_folios(texto: str) -> list[str]:
    """Devuelve todos los folios encontrados en el texto (ruta o nombre)."""
    return list(set(PATRON_FOLIO.findall(texto)))


def escanear_carpeta(raiz: str, etiqueta: str) -> list[dict]:
    """Escanea recursivamente y devuelve lista de dicts con info de cada archivo."""
    resultados = []
    raiz_path = Path(raiz)

    if not raiz_path.exists():
        print(f"  [!] No existe: {raiz}")
        return resultados

    print(f"  Escaneando: {raiz}")

    for archivo in raiz_path.rglob("*"):
        if not archivo.is_file():
            continue
        if archivo.suffix.lower() not in EXTENSIONES_VALIDAS:
            continue

        ruta_completa = str(archivo)
        ruta_relativa = str(archivo.relative_to(raiz_path))
        nombre = archivo.name

        # Intentar extraer folios desde el nombre del archivo y de la carpeta padre
        folios_nombre = extraer_folios(nombre)
        folios_carpeta = extraer_folios(ruta_relativa)
        folios_todos = list(set(folios_nombre + folios_carpeta))

        # Si hay un solo folio claro lo usamos, si hay varios los listamos
        folio_principal = folios_todos[0] if len(folios_todos) == 1 else ""
        folios_multiples = ",".join(folios_todos) if len(folios_todos) > 1 else ""

        categoria = categorizar(ruta_relativa)

        resultados.append({
            "fuente": etiqueta,
            "ruta_completa": ruta_completa,
            "ruta_relativa": ruta_relativa,
            "nombre_archivo": nombre,
            "extension": archivo.suffix.lower(),
            "tamano_kb": round(archivo.stat().st_size / 1024, 1),
            "folio_principal": folio_principal,
            "folios_multiples": folios_multiples,
            "n_folios_detectados": len(folios_todos),
            "categoria": categoria,
        })

    print(f"    -> {len(resultados)} archivos encontrados")
    return resultados


def main():
    print("=" * 70)
    print("ESCANEO DE ARCHIVOS EN DISCO")
    print("=" * 70)

    todos = []

    print("\n[1/2] Archivos (convenios, rendiciones, garantías, egresos)")
    todos.extend(escanear_carpeta(RAIZ_RESPALDOS, "archivos"))

    print("\n[2/2] Registro Colaboradores del Estado (certificados)")
    todos.extend(escanear_carpeta(RAIZ_REGISTRO_ESTADO, "registro_estado"))

    # Crear carpeta logs si no existe
    Path("logs").mkdir(exist_ok=True)

    # Escribir CSV
    if todos:
        with open(OUTPUT_CSV, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=todos[0].keys())
            writer.writeheader()
            writer.writerows(todos)

        print("\n" + "=" * 70)
        print(f"RESULTADO: {len(todos)} archivos inventariados")
        print(f"CSV guardado en: {OUTPUT_CSV}")
        print("=" * 70)

        # Resumen por categoría
        from collections import Counter
        por_cat = Counter(r["categoria"] for r in todos)
        print("\nResumen por categoría:")
        for cat, n in sorted(por_cat.items(), key=lambda x: -x[1]):
            print(f"  {cat:30s} {n:>6d}")

        por_fuente = Counter(r["fuente"] for r in todos)
        print("\nResumen por fuente:")
        for fuente, n in sorted(por_fuente.items()):
            print(f"  {fuente:30s} {n:>6d}")

        # Cuántos folios únicos encontró
        folios_unicos = set()
        for r in todos:
            if r["folio_principal"]:
                folios_unicos.add(r["folio_principal"])
            elif r["folios_multiples"]:
                folios_unicos.update(r["folios_multiples"].split(","))
        print(f"\nFolios únicos detectados: {len(folios_unicos)}")
    else:
        print("\n[!] No se encontraron archivos. Revisá las rutas de configuración.")


if __name__ == "__main__":
    main()
