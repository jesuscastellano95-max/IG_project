# -*- coding: utf-8 -*-
"""
Created on Tue Apr 14 21:22:56 2026

@author: jesus
"""
# Librería para comprobar si existe el Excel
import os

# Librería para leer el Excel
import pandas as pd

# Importamos la ruta del Excel común del proyecto
from common.paths import EXCEL_PATH

# Importamos funciones auxiliares reutilizables
from common.utils import serie_numerica, contar_fallos


# =====================================================
# CÁLCULO DE RESULTADOS OTO MELARA
# =====================================================
def calcular_resultados_oto(datos_entrada: dict) -> dict:
    """
    Calcula los resultados de evaluación para OTO MELARA
    a partir de un diccionario con todos los datos de entrada.

    Entrada esperada en datos_entrada:
    - nombre_prueba
    - r1_min, r1_max
    - r2_min, r2_max
    - limite_desv
    - pmed_lim
    - pmax_lim
    - col_vel
    - col_desv
    - col_pmed
    - col_pmax1
    - col_pmax2
    - col_espoleta
    - col_estopin
    """

    # =====================================================
    # EXTRAER DATOS DEL DICCIONARIO
    # =====================================================
    nombre_prueba = datos_entrada["nombre_prueba"]

    r1_min = datos_entrada["r1_min"]
    r1_max = datos_entrada["r1_max"]
    r2_min = datos_entrada["r2_min"]
    r2_max = datos_entrada["r2_max"]

    limite_desv = datos_entrada["limite_desv"]
    pmed_lim = datos_entrada["pmed_lim"]
    pmax_lim = datos_entrada["pmax_lim"]

    col_vel = datos_entrada["col_vel"]
    col_desv = datos_entrada["col_desv"]
    col_pmed = datos_entrada["col_pmed"]
    col_pmax1 = datos_entrada["col_pmax1"]
    col_pmax2 = datos_entrada["col_pmax2"]
    col_espoleta = datos_entrada["col_espoleta"]
    col_estopin = datos_entrada["col_estopin"]

    # =====================================================
    # COMPROBAR QUE EXISTE EL ARCHIVO EXCEL
    # =====================================================
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"No existe el archivo:\n{EXCEL_PATH}")

    # =====================================================
    # LEER EL EXCEL
    # =====================================================
    df = pd.read_excel(EXCEL_PATH)

    # =====================================================
    # COMPROBAR QUE EXISTEN LAS COLUMNAS NECESARIAS
    # =====================================================
    columnas_necesarias = [
        col_vel,
        col_desv,
        col_pmed,
        col_pmax1,
        col_pmax2,
        col_espoleta,
        col_estopin
    ]

    for c in columnas_necesarias:
        if not c:
            raise ValueError("Hay columnas del Excel sin rellenar.")
        if c not in df.columns:
            raise ValueError(f"La columna '{c}' no existe en el Excel.")

    # =====================================================
    # CÁLCULOS A PARTIR DEL EXCEL
    # =====================================================
    # Velocidad media
    v_med = serie_numerica(df, col_vel).mean()

    # Desviación estándar
    desv = serie_numerica(df, col_desv).std()

    # Presión media
    pmed = serie_numerica(df, col_pmed).mean()

    # Presiones máximas individuales
    pmax1 = serie_numerica(df, col_pmax1).max()
    pmax2 = serie_numerica(df, col_pmax2).max()

    # Presión máxima global (la mayor de ambas)
    pmax = max(pmax1, pmax2)

    # Fallos de espoleta y estopín
    fallos_espoleta = contar_fallos(df, col_espoleta)
    fallos_estopin = contar_fallos(df, col_estopin)

    # =====================================================
    # LÓGICA DE CALIFICACIÓN
    # =====================================================
    # Velocidad:
    # - UTIL-1 si cae dentro del rango estrecho
    # - UTIL-2 si cae fuera del estrecho pero dentro del ancho
    # - INUTIL si queda fuera de ambos
    if r1_min <= v_med <= r1_max:
        res_vel = "UTIL-1"
    elif r2_min <= v_med <= r2_max:
        res_vel = "UTIL-2"
    else:
        res_vel = "INUTIL"

    # Desviación estándar:
    # - UTIL-1 si es menor que el límite
    # - INUTIL si no lo cumple
    res_desv = "UTIL-1" if desv < limite_desv else "INUTIL"

    # Presión máxima:
    # - INUTIL si alguna presión individual supera el límite
    # - UTIL-1 si ninguna lo supera
    if (pmax1 > pmax_lim) or (pmax2 > pmax_lim):
        res_pmax = "INUTIL"
    else:
        res_pmax = "UTIL-1"

    # Presión media:
    # - UTIL-1 si está por debajo del límite
    # - INUTIL si no lo cumple
    res_pmed = "UTIL-1" if pmed < pmed_lim else "INUTIL"

    # Espoleta:
    # - INUTIL si hay 2 o más fallos
    # - UTIL-1 si hay menos de 2
    res_espoleta = "INUTIL" if fallos_espoleta >= 2 else "UTIL-1"

    # Estopín:
    # - INUTIL si hay 1 o más fallos
    # - UTIL-1 si no hay fallos
    res_estopin = "INUTIL" if fallos_estopin >= 1 else "UTIL-1"

    # =====================================================
    # DEVOLVER RESULTADOS
    # =====================================================
    return {
        "nombre_prueba": nombre_prueba,
        "r1_min": r1_min,
        "r1_max": r1_max,
        "r2_min": r2_min,
        "r2_max": r2_max,
        "limite_desv": limite_desv,
        "pmed_lim": pmed_lim,
        "pmax_lim": pmax_lim,
        "v_med": v_med,
        "desv": desv,
        "pmed": pmed,
        "pmax": pmax,
        "pmax1": pmax1,
        "pmax2": pmax2,
        "fallos_espoleta": fallos_espoleta,
        "fallos_estopin": fallos_estopin,
        "res_vel": res_vel,
        "res_desv": res_desv,
        "res_pmax": res_pmax,
        "res_pmed": res_pmed,
        "res_espoleta": res_espoleta,
        "res_estopin": res_estopin,
    }
