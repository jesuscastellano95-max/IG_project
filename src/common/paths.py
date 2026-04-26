# -*- coding: utf-8 -*-
"""
Created on Tue Apr 14 21:30:17 2026

@author: jesus
"""
import os
import sys

# =====================================================
# DETECTAR SI EL PROGRAMA SE EJECUTA COMO SCRIPT NORMAL
# O COMO EJECUTABLE (.exe) GENERADO
# =====================================================

if getattr(sys, "frozen", False):
    # Si está empaquetado como .exe, tomamos como base
    # la carpeta donde está el ejecutable
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Si estamos en desarrollo normal, subimos desde:
    # src/common/paths.py -> src -> IG_project_2
    BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

# =====================================================
# CARPETAS PRINCIPALES DEL PROYECTO
# =====================================================

ASSETS_DIR = os.path.join(BASE_DIR, "assets")
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# =====================================================
# ARCHIVOS COMUNES
# =====================================================

LOGO_PATH = os.path.join(ASSETS_DIR, "SIIGROUPLOGO.jpg")
EXCEL_PATH = os.path.join(DATA_DIR, "excelplantilla.xlsx")

# =====================================================
# PLANTILLAS WORD
# =====================================================

OTO_WORD_TEMPLATE_PATH = os.path.join(ASSETS_DIR, "tabla_oto_melara.docx")

# =====================================================
# ASEGURAR QUE EXISTE LA CARPETA DE SALIDA
# =====================================================

os.makedirs(OUTPUT_DIR, exist_ok=True)