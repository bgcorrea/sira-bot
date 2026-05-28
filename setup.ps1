# Setup inicial - ejecutar UNA sola vez
# Desde PowerShell, parado en la carpeta del proyecto:

# 1. Crear entorno virtual
python -m venv .venv

# 2. Activarlo
.\.venv\Scripts\Activate.ps1

# 3. Si PowerShell te bloquea, ejecutar una vez:
# Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# 4. Instalar dependencias
pip install -r requirements.txt

# 5. Copiar Libro1.xlsx a data\
# (manualmente desde D:\Downloads\Libro1.xlsx)

# 6. Listo. Ejecutar en orden:
# python scripts\01_descubrir_archivos.py
# python scripts\02_extraer_rut_sira.py
# python scripts\03_armar_master.py        (lo armamos después)
# python scripts\04_subir_documentos.py    (lo armamos después)
