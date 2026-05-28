# Criterios de estructura y nomenclatura — SIRA FFOIP

Estos criterios aplican a todos los proyectos (presente y futuro).
El objetivo es que el bot pueda leer los archivos sin lógica especial por región.

---

## Estructura de carpetas

```
Archivos/
  {NN}. {REGION} - FFOIP/
    {folio} - {RAZON_SOCIAL}/
      Convenio   - {folio}.pdf    ← OBLIGATORIO
      Resolución - {folio}.pdf    ← OBLIGATORIO (Acto Administrativo)
      Egreso     - {folio}.pdf    ← OBLIGATORIO (también acepta "Transferencia")
      Rendición  - {folio}.pdf    ← OBLIGATORIO
      Garantía   - {folio}.pdf    ← OPCIONAL (puede ser .jpg, .jpeg, .png)

Colaboradores del Estado/
  {año}/
    certificado_{RUT_SIN_PUNTOS}_{DV}.pdf   ← se cruza por RUT, no va en la carpeta del folio
```

---

## Reglas de nomenclatura

### Nombre de carpeta del folio
- Formato exacto: `{folio} - {RAZON_SOCIAL}`
- El folio es el número de 5-6 dígitos tal como aparece en SIRA.
- La razón social es la que figura en SIRA (extraída por script 02).
- Ejemplo: `59641 - COMUNIDAD INDIGENA DE PARINACOTA`

### Nombre de cada archivo
| Tipo | Nombre estándar | Extensión |
|------|----------------|-----------|
| Convenio firmado | `Convenio - {folio}.pdf` | Solo PDF |
| Resolución / Acto Administrativo | `Resolución - {folio}.pdf` | Solo PDF |
| Egreso / Transferencia / Voucher | `Egreso - {folio}.pdf` | PDF, JPG, PNG |
| Rendición / CFC | `Rendición - {folio}.pdf` | Solo PDF |
| Garantía (letra de cambio u otro) | `Garantía - {folio}.pdf` | PDF, JPG, PNG |

**Regla de oro:** el nombre debe empezar con el tipo seguido de ` - {folio}`.
Esto permite detección automática sin ambigüedad.

### Archivos compartidos (una resolución para varios folios)
Cuando una misma resolución cubre múltiples folios (ej. Arica tiene una REX que aplica a todos):
- Se copia el PDF a cada carpeta individual.
- Se renombra a `Resolución - {folio}.pdf` en cada carpeta.
- El archivo original compartido puede conservarse aparte como referencia.

### Qué NO poner en la carpeta del folio
- El certificado de Colaboradores del Estado — se busca automáticamente por RUT.
- Documentos internos de proceso (check-lists, declaraciones juradas, etc.).
- Archivos `.docx` — SIRA solo acepta PDF e imágenes.

---

## Correspondencia con secciones de SIRA

| Carpeta del folio | Sección en SIRA | Obligatorio |
|-------------------|-----------------|-------------|
| `Convenio - {folio}.pdf` | "Convenio + Acto Administrativo" (archivo 1) | Sí |
| `Resolución - {folio}.pdf` | "Convenio + Acto Administrativo" (archivo 2) | Sí |
| `Colaboradores del Estado/` → por RUT | "Certificado de registro de entidad receptora" | Sí |
| `Egreso - {folio}.pdf` | "Transferencias" | Sí |
| `Rendición - {folio}.pdf` | "Respaldo de rendición" | Sí |
| `Garantía - {folio}.pdf` | "Garantías" | No |

---

## Flujo de trabajo

```
1. Recopilar archivos físicos del proyecto
2. Crear carpeta:  {folio} - {RAZON_SOCIAL}/
3. Nombrar cada archivo según tabla de arriba
4. Ejecutar script 05 (--ejecutar) para mover lo que viene de carpetas compartidas
5. Ejecutar script 06 (validación) para confirmar completitud
6. Completar manualmente los archivos que falten
7. Repetir paso 6 hasta que todos los folios estén "COMPLETO"
8. Ejecutar script 03 para generar master_subida.xlsx
9. Ejecutar script 04 para subir a SIRA
```

---

## Criterios de detección automática

El script de validación detecta el tipo de documento por el nombre del archivo.
Acepta tanto el nombre estándar como variaciones históricas:

| Tipo | Palabras clave aceptadas en el nombre |
|------|---------------------------------------|
| convenio | `convenio` |
| resolucion | `resolución`, `resolucion`, `res.`, `rex`, `adjudicación`, `exenta` |
| egreso | `egreso`, `transferencia`, `recepción`, `recepcion`, `voucher`, `certificado bancario` |
| rendicion | `rendición`, `rendicion`, `cfc`, `fiel cumplimiento`, `memo daf` |
| garantia | `garantía`, `garantia`, `letra de cambio` |

**Regla de prioridad:** si el nombre empieza con `{Tipo} - {folio}` (nombre estándar),
se usa ese tipo directamente sin buscar palabras clave.

---

## Errores comunes a evitar

| Error | Consecuencia |
|-------|-------------|
| Archivo `.docx` en vez de PDF | SIRA lo rechaza |
| Dos archivos del mismo tipo en la misma carpeta | El script elige uno arbitrariamente |
| Carpeta sin número de folio al inicio | No se detecta |
| Certificado de Colaboradores en la carpeta del folio | Se duplica en SIRA |
| Folio con dos carpetas (nombre distinto) | El script usa solo la primera que encuentra |
