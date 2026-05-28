import csv, openpyxl
from pathlib import Path

# Leer los 250 folios del Excel
wb = openpyxl.load_workbook(r"D:\bot_sira\data\Libro1.xlsx", data_only=True)
ws = wb.active
folios_250 = {str(row[2]).strip() for row in ws.iter_rows(min_row=2, values_only=True) if row[2]}
print(f"Folios objetivo: {len(folios_250)}")

# Leer el inventario completo y filtrar
relevantes = []
with open(r"D:\bot_sira\logs\inventario_disco.csv", encoding="utf-8-sig") as f_in:
    for r in csv.DictReader(f_in):
        folios_archivo = set()
        if r["folio_principal"]:
            folios_archivo.add(r["folio_principal"])
        if r["folios_multiples"]:
            folios_archivo.update(f.strip() for f in r["folios_multiples"].split(","))
        
        # Folios de este archivo que están en nuestros 250
        matched = folios_archivo & folios_250
        if matched:
            r["folio_match"] = ",".join(sorted(matched))
            r["n_matches"] = len(matched)
            relevantes.append(r)

print(f"Archivos que pertenecen a tus 250 folios: {len(relevantes)}")

# Guardar
out_path = r"D:\bot_sira\logs\inventario_relevante.csv"
with open(out_path, "w", encoding="utf-8-sig", newline="") as f_out:
    w = csv.DictWriter(f_out, fieldnames=relevantes[0].keys())
    w.writeheader()
    w.writerows(relevantes)
print(f"OK -> {out_path}")

# Resumen rápido
from collections import Counter
print("\nPor categoría:")
for cat, n in Counter(r["categoria"] for r in relevantes).most_common():
    print(f"  {cat:25s} {n:>5d}")

print("\nPor extensión:")
for ext, n in Counter(r["extension"] for r in relevantes).most_common():
    print(f"  {ext:25s} {n:>5d}")

# Folios de los 250 que NO aparecen en disco
folios_encontrados = set()
for r in relevantes:
    folios_encontrados.update(r["folio_match"].split(","))
folios_sin_archivos = folios_250 - folios_encontrados
print(f"\nFolios SIN archivos en disco: {len(folios_sin_archivos)}")
if folios_sin_archivos:
    print(f"  Lista: {sorted(folios_sin_archivos)[:20]}{'...' if len(folios_sin_archivos) > 20 else ''}")
