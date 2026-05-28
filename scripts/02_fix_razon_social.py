"""
Script 02-FIX v3: Re-parseo final con estados normalizados
============================================================
Descubrimiento: cuando un convenio está en estado "Enviado", el badge en
el body aparece como "✓ Enviado" (con checkmark), no como "Enviado" solo.
Lo mismo podría pasar con "Cerrado" cuando esté aprobado.

Este script:
1. Normaliza el estado (quita el "✓ " del prefijo si existe).
2. Re-parsea la razón social usando la estructura confirmada.
3. Guarda el CSV final limpio.

Uso:
    python scripts\02_fix_razon_social.py
"""

import json
import csv
import re
from pathlib import Path

JSONL_INPUT = "logs/sira_raw_dump.jsonl"
CSV_INPUT = "logs/rut_por_folio.csv"
CSV_OUTPUT = "logs/rut_por_folio_fixed.csv"

# Estados conocidos, en cualquier variante
ESTADOS_BASE = {"Borrador", "Enviado", "Cerrado", "Vigente", "En revisión"}


def normalizar_estado(linea: str) -> str | None:
    """Quita el checkmark '✓ ' si existe y verifica que sea un estado conocido."""
    limpia = linea.strip()
    if limpia.startswith("✓ "):
        limpia = limpia[2:].strip()
    if limpia in ESTADOS_BASE:
        return limpia
    return None


def extraer_razon_y_estado(body: str, folio: str) -> tuple[str, str, str]:
    """
    Estructura confirmada del body:
        Convenio
        {folio}
        {estado o "✓ estado"}    <- normalizamos
        {RAZON_SOCIAL}            <- con o sin " / NOMBRE_FANTASIA"
        SUBSECRETARÍA...
    """
    lineas = [l.strip() for l in body.split("\n")]
    
    for i, l in enumerate(lineas):
        if l == "Convenio" and i + 1 < len(lineas) and lineas[i + 1] == folio:
            # i+2 podría ser el estado (con o sin checkmark)
            if i + 2 < len(lineas):
                estado = normalizar_estado(lineas[i + 2])
                if estado is not None:
                    # i+3 es la razón social
                    if i + 3 < len(lineas):
                        raw = lineas[i + 3]
                        if raw.startswith("SUBSECRETARÍA"):
                            return "", "", estado
                        if " / " in raw:
                            partes = [p.strip() for p in raw.split(" / ", 1)]
                            return partes[0], partes[1] if len(partes) > 1 else "", estado
                        return raw, "", estado
            break
    
    return "", "", ""


def main():
    print("=" * 70)
    print("RE-PARSEO FINAL: razón social + estado normalizado")
    print("=" * 70)
    
    dump = {}
    with open(JSONL_INPUT, encoding="utf-8") as f:
        for line in f:
            d = json.loads(line)
            dump[d["folio"]] = d
    print(f"\nFolios en JSONL: {len(dump)}")
    
    with open(CSV_INPUT, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    print(f"Folios en CSV:   {len(rows)}\n")
    
    cambios_razon = 0
    cambios_estado = 0
    sin_razon = []
    
    for r in rows:
        folio = r["folio"]
        if folio not in dump:
            continue
        body = dump[folio].get("body_raw", "")
        
        razon, fantasia, estado = extraer_razon_y_estado(body, folio)
        
        if not razon:
            sin_razon.append(folio)
            continue
        
        # Detectar cambios y reportar
        if r["razon_social"] != razon:
            print(f"  ✓ {folio}: razón '{r['razon_social'][:35]:35s}' → '{razon[:50]}'")
            cambios_razon += 1
        if r["estado_sira"] != estado and estado:
            print(f"    {folio}: estado '{r['estado_sira']}' → '{estado}'")
            cambios_estado += 1
        
        r["razon_social"] = razon
        r["nombre_fantasia"] = fantasia
        if estado:
            r["estado_sira"] = estado
    
    if rows:
        with open(CSV_OUTPUT, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=rows[0].keys())
            w.writeheader()
            w.writerows(rows)
    
    print(f"\n{'=' * 70}")
    print(f"Razón social: {cambios_razon} cambios")
    print(f"Estado:       {cambios_estado} cambios")
    print(f"Sin razón:    {len(sin_razon)}")
    if sin_razon:
        print(f"  Folios: {sin_razon}")
    print(f"\nOutput: {CSV_OUTPUT}")
    
    # Verificación
    rotos = [r["folio"] for r in rows if r["razon_social"] == "CONCURSO PÚBLICO"]
    print(f"\nFolios con 'CONCURSO PÚBLICO' tras fix: {len(rotos)}")
    
    # Estados finales
    from collections import Counter
    estados_finales = Counter(r["estado_sira"] for r in rows)
    print(f"\nEstados finales: {dict(estados_finales)}")


if __name__ == "__main__":
    main()
