# -*- coding: utf-8 -*-
"""
Created on Tue Apr 14 21:31:23 2026

@author: jesus
"""
# Librería para dar formato al texto dentro del Word (tamaño, fuente, etc.)
from docx.shared import Pt

# Librería para trabajar con datos tipo Excel (DataFrames)
import pandas as pd


# =====================================================
# CONVERSIÓN DE TEXTO A FLOAT
# =====================================================
def convertir_a_float(texto: str) -> float:
    """
    Convierte un texto a número decimal.
    Acepta tanto coma como punto como separador decimal.
    Ejemplo:
        "3,5" -> 3.5
        "3.5" -> 3.5
    """
    return float(texto.strip().replace(",", "."))


# =====================================================
# FORMATEO DE NÚMEROS PARA MOSTRAR / ESCRIBIR EN WORD
# =====================================================
def formatear_numero(valor: float) -> str:
    """
    Formatea un número para que quede bien en el informe:
    
    - Si es entero → sin decimales (ej: 5)
    - Si es decimal → 2 decimales con coma (ej: 5,23)
    """
    try:
        v = float(valor)
    except Exception:
        # Si no se puede convertir, lo devuelve tal cual
        return str(valor)

    # Si es entero (ej: 5.0)
    if v.is_integer():
        return f"{int(v)}"

    # Si tiene decimales → 2 cifras y coma
    return f"{v:.2f}".replace(".", ",")


# =====================================================
# ESCRIBIR TEXTO EN UNA CELDA DE WORD
# =====================================================
def set_cell_text(cell, text: str):
    """
    Escribe texto en una celda de Word con formato:
    - Fuente: Times New Roman
    - Tamaño: 10
    """
    # Limpiar contenido previo
    cell.text = ""

    # Acceder al párrafo de la celda
    p = cell.paragraphs[0]

    # Añadir texto como "run" (bloque de texto con formato)
    run = p.add_run(text)

    # Configurar formato
    run.font.name = "Times New Roman"
    run.font.size = Pt(10)


# =====================================================
# LIMPIEZA DE COLUMNAS NUMÉRICAS (EXCEL → PANDAS)
# =====================================================
def serie_numerica(df: pd.DataFrame, col: str) -> pd.Series:
    """
    Convierte una columna del DataFrame a valores numéricos.
    
    Problemas que soluciona:
    - Números como texto ("3,5")
    - Mezcla de texto y números
    - Valores inválidos
    
    Resultado:
    - Serie numérica limpia
    - Sin valores NaN
    """
    s = df[col]

    # Si la columna es texto
    if s.dtype == object:
        # Cambia coma por punto
        s = s.astype(str).str.replace(",", ".", regex=False)

    # Convertir a número (lo que no se pueda → NaN)
    s = pd.to_numeric(s, errors="coerce")

    # Eliminar valores no válidos
    return s.dropna()


# =====================================================
# CONTAR FALLOS EN UNA COLUMNA
# =====================================================
def contar_fallos(df: pd.DataFrame, col: str) -> int:
    """
    Cuenta cuántas veces aparece la palabra 'Fallo' en una columna.
    
    - Ignora mayúsculas/minúsculas
    - Ignora espacios
    
    Ejemplo:
        "Fallo", "fallo ", " FALLO" → todos cuentan
    """
    # Convertir a texto, limpiar espacios y pasar a minúsculas
    s = df[col].astype(str).str.strip().str.lower()

    # Contar cuántos son exactamente "fallo"
    return int((s == "fallo").sum())
