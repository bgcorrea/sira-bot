import json

# El que aún quedó con CONCURSO PÚBLICO y uno de los Enviados
folios_a_ver = ["54622", "54872"]

with open(r"logs\sira_raw_dump.jsonl", encoding="utf-8") as f:
    bodies = {}
    for line in f:
        d = json.loads(line)
        if d["folio"] in folios_a_ver:
            bodies[d["folio"]] = d["body_raw"]

for folio in folios_a_ver:
    print("=" * 70)
    print(f"FOLIO {folio} - body completo:")
    print("=" * 70)
    print(bodies.get(folio, "(no encontrado)"))
    print()

# También buscar cuál folio quedó con CONCURSO PÚBLICO en el CSV
import csv
print("=" * 70)
print("Folio con razon='CONCURSO PUBLICO' en CSV_FIXED:")
print("=" * 70)
with open(r"logs\rut_por_folio_fixed.csv", encoding="utf-8-sig") as f:
    for r in csv.DictReader(f):
        if r["razon_social"] == "CONCURSO PÚBLICO":
            print(f"  Folio: {r['folio']}, estado: {r['estado_sira']}, región: {r['region_excel']}")
