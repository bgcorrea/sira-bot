# Bot SIRA - Carga masiva de documentos FFOIP 2022

## Estructura del proyecto

```
bot_sira/
├── data/
│   ├── Libro1.xlsx                      # Input: folios a procesar
│   └── master_subida.xlsx               # Output etapa 2: mapping completo
├── scripts/
│   ├── 01_descubrir_archivos.py         # Escanea D:\Downloads y arma inventario
│   ├── 02_extraer_rut_sira.py           # Login manual + scraping de RUT/entidad
│   ├── 03_armar_master.py               # Cruza todo y arma master_subida.xlsx
│   └── 04_subir_documentos.py           # Bot final que sube a SIRA
├── logs/
│   ├── inventario_disco.csv             # Output etapa 1
│   ├── rut_por_folio.csv                # Output etapa 2
│   └── ejecucion_YYYYMMDD_HHMMSS.csv    # Log de cada subida
└── README.md
```

## Flujo de ejecución

### Etapa 1: Descubrir archivos en disco
```powershell
python scripts/01_descubrir_archivos.py
```
Output: `logs/inventario_disco.csv` con todos los PDFs encontrados.

### Etapa 2: Extraer RUT desde SIRA
```powershell
python scripts/02_extraer_rut_sira.py
```
Te abrirá Chrome, harás login manual, presionas Enter en la terminal, y el bot
recorre los 250 folios extrayendo RUT y razón social.
Output: `logs/rut_por_folio.csv`.

### Etapa 3: Armar master de subida
```powershell
python scripts/03_armar_master.py
```
Cruza Libro1.xlsx + inventario_disco.csv + rut_por_folio.csv.
Output: `data/master_subida.xlsx` con una fila por folio y columnas:
- folio
- rut
- razon_social
- region
- convenio_pdf
- acto_admin_pdf
- certificado_registro_pdf
- voucher_pdf
- rendicion_pdf
- garantia_pdf
- estado_subida (vacío al inicio)

### Etapa 4: Subir a SIRA
```powershell
python scripts/04_subir_documentos.py
```
Te pide hacer login manual. Después itera el master e va subiendo todo.
Idempotente: si lo cortás y reanudás, retoma desde donde quedó.
```
