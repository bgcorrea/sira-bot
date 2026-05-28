#!/bin/bash
# Setup para Linux/WSL — ejecutar UNA sola vez desde bot_sira/

# 1. Crear entorno virtual
python3 -m venv .venv

# 2. Activar e instalar dependencias
.venv/bin/pip install -r requirements.txt

# 3. Verificar Chrome
google-chrome --version || { echo "Chrome no encontrado. Instalar con:"; echo "  sudo apt-get install google-chrome-stable"; }

# 4. Ejecutar en orden:
# cd bot_sira
# .venv/bin/python scripts/01_descubrir_archivos.py
# .venv/bin/python scripts/02_extraer_rut_sira.py
# .venv/bin/python scripts/03_armar_master.py
# .venv/bin/python scripts/04_subir_documentos.py
