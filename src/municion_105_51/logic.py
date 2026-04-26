# -*- coding: utf-8 -*-
"""
Created on Wed Apr 22 20:39:38 2026

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
# FUNCIÓN AUXILIAR: CONTAR VALORES NUMÉRICOS MENORES QUE UN LÍMITE
# =====================================================
def contar_menores_que(df: pd.DataFrame, col: str, limite: float) -> int:
    """
    Cuenta cuántos valores numéricos de una columna son menores que un límite.
    Se usa para la condición de Trazador.
    """
    s = serie_numerica(df, col)
    return int((s < limite).sum())


# =====================================================
# FUNCIÓN AUXILIAR: FILTRAR EL DATAFRAME POR SERIE
# =====================================================
def normalizar_serie(valor) -> str:
    """
    Normaliza el valor de la serie para que '+21' y '21'
    se consideren equivalentes.
    """
    texto = str(valor).strip()

    if texto == "+21":
        return "21"
    return texto


def filtrar_por_serie(df: pd.DataFrame, col_serie: str, serie_objetivo: str) -> pd.DataFrame:
    """
    Filtra el DataFrame por la serie seleccionada (+21 o -35),
    normalizando los valores para evitar problemas de formato.
    """
    serie_obj = normalizar_serie(serie_objetivo)
    serie_col = df[col_serie].apply(normalizar_serie)

    return df[serie_col == serie_obj].copy()


# =====================================================
# CÁLCULO DE RESULTADOS PARA MUNICIÓN 105/51
# =====================================================
def calcular_resultados_105_51(datos_entrada: dict) -> dict:
    """
    Calcula los resultados de evaluación para munición 105/51
    a partir de un diccionario con todos los datos de entrada.

    El cálculo se realiza:
    - por serie para la mayoría de condiciones
    - sobre ambas series para Trazador
    """

    # =====================================================
    # EXTRAER DATOS DEL DICCIONARIO
    # =====================================================
    nombre_prueba = datos_entrada["nombre_prueba"]
    serie_seleccionada = datos_entrada["serie_seleccionada"]

    r1_min = datos_entrada["r1_min"]
    r1_max = datos_entrada["r1_max"]
    r2_min = datos_entrada["r2_min"]
    r2_max = datos_entrada["r2_max"]

    desv_factor = datos_entrada["desv_factor"]              # 0.7
    desv_lim = datos_entrada["desv_lim"]                    # 3
    pmed_lim = datos_entrada["pmed_lim"]                    # 420 o 470
    pmed_sigma_lim = datos_entrada["pmed_sigma_lim"]        # 505
    pmax_lim = datos_entrada["pmax_lim"]                    # 505
    trazador_lim = datos_entrada["trazador_lim"]            # 2.15
    trazador_max_fallos = datos_entrada["trazador_max_fallos"]  # 2 (si hay 3 o más => inútil)

    col_serie = datos_entrada["col_serie"]
    col_vel = datos_entrada["col_vel"]
    col_desv = datos_entrada["col_desv"]
    col_pmed = datos_entrada["col_pmed"]
    col_pmax1 = datos_entrada["col_pmax1"]
    col_pmax2 = datos_entrada["col_pmax2"]

    col_espoleta = datos_entrada["col_espoleta"]
    col_estopin = datos_entrada["col_estopin"]
    col_separacion = datos_entrada["col_separacion"]
    col_vuelo = datos_entrada["col_vuelo"]
    col_trazador = datos_entrada["col_trazador"]

    # =====================================================
    # COMPROBAR QUE EXISTE EL ARCHIVO EXCEL
    # =====================================================
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(f"No existe el archivo:\n{EXCEL_PATH}")

    # =====================================================
    # LEER EL EXCEL (HOJA 105_51)
    # =====================================================
    df = pd.read_excel(EXCEL_PATH, sheet_name="105_51")

    # =====================================================
    # COMPROBAR QUE EXISTEN LAS COLUMNAS NECESARIAS
    # =====================================================
    columnas_necesarias = [
        col_serie,
        col_vel,
        col_desv,
        col_pmed,
        col_pmax1,
        col_pmax2,
        col_espoleta,
        col_estopin,
        col_separacion,
        col_vuelo,
        col_trazador,
    ]

    for c in columnas_necesarias:
        if not c:
            raise ValueError("Hay columnas del Excel sin rellenar.")
        if c not in df.columns:
            raise ValueError(f"La columna '{c}' no existe en la hoja 105_51 del Excel.")

    # =====================================================
    # FILTRAR POR LA SERIE SELECCIONADA
    # =====================================================
    df_serie = filtrar_por_serie(df, col_serie, serie_seleccionada)

    if df_serie.empty:
        raise ValueError(
            f"No se han encontrado filas para la serie '{serie_seleccionada}' "
            f"en la columna '{col_serie}'."
        )

    # =====================================================
    # CÁLCULOS NUMÉRICOS DE LA SERIE
    # =====================================================
    # Velocidad media corregida de la serie
    v_med = serie_numerica(df_serie, col_vel).mean()

    # Desviación estándar de la serie para la condición de velocidad
    desv = serie_numerica(df_serie, col_desv).std()

    # Valor corregido de desviación: sigma * 0.7
    desv_corregida = desv * desv_factor

    # Presión media de la serie
    pmed = serie_numerica(df_serie, col_pmed).mean()

    # Desviación estándar de la presión media
    sigma_pmed = serie_numerica(df_serie, col_pmed).std()

    # Presión media corregida: Pmed + 3 * sigma(Pmed)
    pmed_sigma = pmed + 3 * sigma_pmed

    # Presión máxima individual de la serie
    pmax1 = serie_numerica(df_serie, col_pmax1).max()
    pmax2 = serie_numerica(df_serie, col_pmax2).max()
    pmax = max(pmax1, pmax2)

    # =====================================================
    # CÁLCULOS DE FALLOS DE LA SERIE
    # =====================================================
    fallos_espoleta = contar_fallos(df_serie, col_espoleta)
    fallos_estopin = contar_fallos(df_serie, col_estopin)

    # Proyectil: dos condiciones independientes
    fallos_separacion = contar_fallos(df_serie, col_separacion)
    fallos_vuelo = contar_fallos(df_serie, col_vuelo)

    # =====================================================
    # CÁLCULO GLOBAL DE TRAZADOR (AMBAS SERIES)
    # =====================================================
    # Si hay 3 o más disparos con tiempo < 2.15 s => INUTIL
    vuelos_trazador_fuera = contar_menores_que(df, col_trazador, trazador_lim)

    # =====================================================
    # LÓGICA DE CALIFICACIÓN
    # =====================================================

    # -----------------------------------------------------
    # VELOCIDAD
    # -----------------------------------------------------
    # UTIL-1: rango fijo 1161 - 1186
    # UTIL-2: Vnom ± 2%
    # INUTIL: fuera de ambos
    if r1_min <= v_med <= r1_max:
        res_vel = "UTIL-1"
    elif r2_min <= v_med <= r2_max:
        res_vel = "UTIL-2"
    else:
        res_vel = "INUTIL"

    # -----------------------------------------------------
    # DESVIACIÓN
    # -----------------------------------------------------
    # sigma * 0.7 <= 3 => UTIL-1
    # si no => INUTIL
    res_desv = "UTIL-1" if desv_corregida <= desv_lim else "INUTIL"

    # -----------------------------------------------------
    # PRESIÓN MÁXIMA INDIVIDUAL
    # -----------------------------------------------------
    res_pmax = "UTIL-1" if pmax <= pmax_lim else "INUTIL"

    # -----------------------------------------------------
    # PRESIÓN MEDIA DE LA SERIE
    # -----------------------------------------------------
    res_pmed = "UTIL-1" if pmed <= pmed_lim else "INUTIL"

    # -----------------------------------------------------
    # PMED + 3 * SIGMA(PMED)
    # -----------------------------------------------------
    res_pmed_sigma = "UTIL-1" if pmed_sigma <= pmed_sigma_lim else "INUTIL"

    # -----------------------------------------------------
    # PROYECTIL - SEPARACIÓN DE PARTES METÁLICAS
    # -----------------------------------------------------
    res_separacion = "INUTIL" if fallos_separacion >= 1 else "UTIL-1"

    # -----------------------------------------------------
    # PROYECTIL - ANOMALÍAS EN VUELO
    # -----------------------------------------------------
    res_vuelo = "INUTIL" if fallos_vuelo >= 1 else "UTIL-1"

    # -----------------------------------------------------
    # ESPOLETA (MISMA LÓGICA QUE OTO)
    # -----------------------------------------------------
    res_espoleta = "INUTIL" if fallos_espoleta >= 2 else "UTIL-1"

    # -----------------------------------------------------
    # ESTOPÍN (MISMA LÓGICA QUE OTO)
    # -----------------------------------------------------
    res_estopin = "INUTIL" if fallos_estopin >= 1 else "UTIL-1"

    # -----------------------------------------------------
    # TRAZADOR (GLOBAL SOBRE AMBAS SERIES)
    # -----------------------------------------------------
    # si hay 3 o más disparos fuera => inútil
    res_trazador = "INUTIL" if vuelos_trazador_fuera > trazador_max_fallos else "UTIL-1"

    # =====================================================
    # DEVOLVER RESULTADOS
    # =====================================================
    return {
        "nombre_prueba": nombre_prueba,
        "serie_seleccionada": serie_seleccionada,

        "r1_min": r1_min,
        "r1_max": r1_max,
        "r2_min": r2_min,
        "r2_max": r2_max,

        "desv_factor": desv_factor,
        "desv_lim": desv_lim,
        "pmed_lim": pmed_lim,
        "pmed_sigma_lim": pmed_sigma_lim,
        "pmax_lim": pmax_lim,
        "trazador_lim": trazador_lim,
        "trazador_max_fallos": trazador_max_fallos,

        "v_med": v_med,
        "desv": desv,
        "desv_corregida": desv_corregida,
        "pmed": pmed,
        "sigma_pmed": sigma_pmed,
        "pmed_sigma": pmed_sigma,
        "pmax1": pmax1,
        "pmax2": pmax2,
        "pmax": pmax,

        "fallos_espoleta": fallos_espoleta,
        "fallos_estopin": fallos_estopin,
        "fallos_separacion": fallos_separacion,
        "fallos_vuelo": fallos_vuelo,
        "vuelos_trazador_fuera": vuelos_trazador_fuera,

        "res_vel": res_vel,
        "res_desv": res_desv,
        "res_pmax": res_pmax,
        "res_pmed": res_pmed,
        "res_pmed_sigma": res_pmed_sigma,
        "res_separacion": res_separacion,
        "res_vuelo": res_vuelo,
        "res_espoleta": res_espoleta,
        "res_estopin": res_estopin,
        "res_trazador": res_trazador,
    }
