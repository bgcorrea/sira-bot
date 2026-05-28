"""
Script 09 — Reporte de folios listos para enviar a revisión en SIRA.

Un folio está listo cuando tiene los 6 archivos obligatorios en disco:
  1. Convenio
  2. Resolución (acto administrativo)
  3. Certificado de registro (colaboradores del Estado)
  4. Egreso / Transferencias (voucher)
  5. Rendición

Se cruza contra DISTRIBUCIÓN CARGA VB para procesar solo los 244 asignados.

Uso:
    python scripts/09_listos_para_enviar.py
"""

from pathlib import Path
import pandas as pd

MASTER       = Path("data/master_subida.xlsx")
DISTRIBUCION = Path("data/DISTRIBUCIÓN CARGA VB.xlsx")
LOG_OUT      = Path("logs/listos_para_enviar.csv")

CAMPOS_REQUERIDOS = {
    "convenio_pdf":     "Convenio",
    "acto_admin_pdf":   "Resolución",
    "certificado_pdf":  "Certificado",
    "voucher_pdf":      "Transferencias",
    "rendicion_pdf":    "Rendición",
}


def main() -> None:
    master = pd.read_excel(MASTER)
    dist   = pd.read_excel(DISTRIBUCION)

    folios_dist = set(dist["ID Convenio"].astype(str).str.strip())
    df = master[master["folio"].astype(str).isin(folios_dist)].copy()

    listos      = []
    incompletos = []

    for _, row in df.iterrows():
        folio  = str(row["folio"])
        region = str(row.get("region", "?"))
        razon  = str(row.get("razon_social", "?"))

        faltantes = []
        for campo, nombre in CAMPOS_REQUERIDOS.items():
            val = str(row.get(campo, "")).strip()
            if not val or val == "nan":
                faltantes.append(nombre)
            elif not Path(val).exists():
                faltantes.append(f"{nombre}(archivo_no_en_disco)")

        if not faltantes:
            listos.append({"folio": folio, "region": region, "razon_social": razon})
        else:
            incompletos.append({
                "folio": folio, "region": region, "razon_social": razon,
                "faltante": "; ".join(faltantes),
            })

    # ── Consola ───────────────────────────────────────────────────────────────
    print(f"\n{'='*65}")
    print(f"  FOLIOS LISTOS PARA ENVIAR A REVISIÓN EN SIRA")
    print(f"{'='*65}")
    print(f"\n✅  Listos para enviar:    {len(listos)}")
    print(f"⚠️   Incompletos (faltan):  {len(incompletos)}")

    # Resumen por región — listos
    print("\n── LISTOS POR REGIÓN ──")
    por_reg: dict[str, int] = {}
    for r in listos:
        por_reg[r["region"]] = por_reg.get(r["region"], 0) + 1
    for reg in sorted(por_reg):
        print(f"  {reg:<15} {por_reg[reg]:>3} folios")

    # Incompletos agrupados
    print("\n── INCOMPLETOS POR REGIÓN ──")
    inc_reg: dict[str, list] = {}
    for r in incompletos:
        inc_reg.setdefault(r["region"], []).append(r)
    for reg in sorted(inc_reg):
        print(f"\n  [{reg}] ({len(inc_reg[reg])} folios)")
        for r in inc_reg[reg]:
            print(f"    {r['folio']} — {r['razon_social'][:45]}")
            print(f"      Falta: {r['faltante']}")

    # ── CSV ───────────────────────────────────────────────────────────────────
    LOG_OUT.parent.mkdir(exist_ok=True)
    rows = []
    for r in listos:
        rows.append({**r, "estado": "LISTO", "faltante": ""})
    for r in incompletos:
        rows.append({**r, "estado": "INCOMPLETO"})

    pd.DataFrame(rows).to_csv(LOG_OUT, index=False, encoding="utf-8-sig")
    print(f"\n  CSV exportado → {LOG_OUT}")
    print()


if __name__ == "__main__":
    main()
