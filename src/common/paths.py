# -*- coding: utf-8 -*-
"""
Created on Tue Apr 14 21:30:17 2026

@author: jesus
"""
import os

# Carpeta src
SRC_DIR = os.path.dirname(os.path.abspath(__file__))
#__file__-->ruta del archivo actual (paths.py)
#os.path.abspath(__file__)-->convierte esto en ruta absoluta
#os.path.dirname(...) → se queda con la carpeta donde está ese archivo

# Carpeta raíz del proyecto (subimos un nivel desde src/common)
PROJECT_DIR = os.path.abspath(os.path.join(SRC_DIR, "..", ".."))
#.. significa “subir una carpeta”
#os.path.join(SRC_DIR, "..", "..") → sube dos niveles:
#-de common → src
#-de src → IG_project
#abspath limpia la ruta

# Carpetas principales
ASSETS_DIR = os.path.join(PROJECT_DIR, "assets")
DATA_DIR = os.path.join(PROJECT_DIR, "data")
OUTPUT_DIR = os.path.join(PROJECT_DIR, "output")
#os.path.join une rutas correctamente

# Archivos comunes
LOGO_PATH = os.path.join(ASSETS_DIR, "SIIGROUPLOGO.jpg")
EXCEL_PATH = os.path.join(DATA_DIR, "excelplantilla.xlsx")
# Plantillas Word
OTO_WORD_TEMPLATE_PATH = os.path.join(ASSETS_DIR, "tabla_oto_melara.docx")
#aquí simplemente está apuntando a archivos especificos dentro de carpetas

# Asegurar carpeta de salida
os.makedirs(OUTPUT_DIR, exist_ok=True)
#crea la carpeta output si no existe,
#si ya existe → no da error (exist_ok=True).

##RESUMEN##
#dirname → dame la carpeta
#abspath → conviértelo en ruta completa
#join → une rutas bien
#makedirs → crea carpeta si no exist