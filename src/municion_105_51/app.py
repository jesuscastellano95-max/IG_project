# -*- coding: utf-8 -*-
"""
Created on Wed Apr 22 20:38:58 2026

@author: jesus
"""
# =====================================================
# AJUSTE DE RUTA PARA EJECUCIÓN DIRECTA
# =====================================================
import os
import sys

SRC_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)

# =====================================================
# IMPORTACIONES
# =====================================================
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk

from common.paths import LOGO_PATH
from common.utils import convertir_a_float, formatear_numero, formatear_calificacion
from municion_105_51.logic import calcular_resultados_105_51
from municion_105_51.word_generator import generar_word_105_51


# =====================================================
# CONFIGURACIÓN FIJA DEL MÓDULO
# =====================================================
TITULO_MODULO = "Vigilancia de munición 105/51"
SUBTITULO_MODULO = "Aplicable a VRC y Centauro"

# Rango fijo para ÚTIL-1 en velocidad
R1_MIN_FIJO = 1161
R1_MAX_FIJO = 1186

# Desviación
DESV_FACTOR = 0.7
DESV_LIM = 3.0

# Presiones
PMED_LIM_MAS_21 = 420.0
PMED_LIM_MENOS_35 = 470.0
PMED_SIGMA_LIM = 505.0
PMAX_LIM = 505.0

# Trazador
TRAZADOR_LIM = 2.15
TRAZADOR_MAX_FALLOS = 2  # 3 o más => INUTIL


# =====================================================
# VARIABLES DE ESTADO
# =====================================================
RESULTADOS = None
VNOM_GLOBAL = None
SERIE_GLOBAL = None


# =====================================================
# PANTALLA 3 — CONFIRMACIÓN Y GENERACIÓN DE WORD
# =====================================================
def abrir_pantalla_3():
    global ventana3

    if RESULTADOS is None:
        messagebox.showerror("Error", "Primero debes calcular los resultados.")
        return

    ventana3 = tk.Toplevel(ventana2)
    ventana3.title("105/51 - Confirmación")

    # -------------------------------------------------
    # HEADER
    # -------------------------------------------------
    header = ttk.Frame(ventana3)
    header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(10, 5), padx=10)

    try:
        img = Image.open(LOGO_PATH).resize((250, 100))
        ventana3.logo_img = ImageTk.PhotoImage(img)
        ttk.Label(header, image=ventana3.logo_img).grid(row=0, column=0, rowspan=3, padx=(0, 12))
    except Exception:
        ttk.Label(header, text="(Logo no encontrado)").grid(row=0, column=0, rowspan=3, padx=(0, 12))

    ttk.Label(header, text=TITULO_MODULO, font=("Arial", 12, "bold")).grid(row=0, column=1, sticky="w")
    ttk.Label(header, text=SUBTITULO_MODULO, font=("Arial", 10)).grid(row=1, column=1, sticky="w")
    ttk.Label(
        header,
        text=f"Vnom: {formatear_numero(VNOM_GLOBAL)} m/s | Serie: {SERIE_GLOBAL} ºC",
        font=("Arial", 10)
    ).grid(row=2, column=1, sticky="w")

    # -------------------------------------------------
    # TABLA VISUAL DE RESULTADOS
    # -------------------------------------------------
    frame = ttk.LabelFrame(ventana3, text="Resultados calculados (confirmación)")
    frame.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

    filas = [
    ("Velocidad media", f"{formatear_numero(RESULTADOS['v_med'])} m/s", formatear_calificacion(RESULTADOS["res_vel"])),
    ("Desv. corregida", f"{formatear_numero(RESULTADOS['desv_corregida'])} m/s", formatear_calificacion(RESULTADOS["res_desv"])),
    ("Pmáx", f"{formatear_numero(RESULTADOS['pmax'])} MPa", formatear_calificacion(RESULTADOS["res_pmax"])),
    ("Pmed", f"{formatear_numero(RESULTADOS['pmed'])} MPa", formatear_calificacion(RESULTADOS["res_pmed"])),
    ("Pmed + 3·σPmed", f"{formatear_numero(RESULTADOS['pmed_sigma'])} MPa", formatear_calificacion(RESULTADOS["res_pmed_sigma"])),
    ("Separación partes metálicas", f"{RESULTADOS['fallos_separacion']} fallos", formatear_calificacion(RESULTADOS["res_separacion"])),
    ("Anomalías en vuelo", f"{RESULTADOS['fallos_vuelo']} fallos", formatear_calificacion(RESULTADOS["res_vuelo"])),
    ("Espoleta", f"{RESULTADOS['fallos_espoleta']} fallos", formatear_calificacion(RESULTADOS["res_espoleta"])),
    ("Estopín", f"{RESULTADOS['fallos_estopin']} fallos", formatear_calificacion(RESULTADOS["res_estopin"])),
    (
        "Trazador (global)",
        f"{RESULTADOS['vuelos_trazador_fuera']} vuelos < {formatear_numero(RESULTADOS['trazador_lim'])} s",
        formatear_calificacion(RESULTADOS["res_trazador"])
    ),
]

    ttk.Label(frame, text="Condición", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=6, pady=4, sticky="w")
    ttk.Label(frame, text="Valor", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=6, pady=4, sticky="w")
    ttk.Label(frame, text="Calificación", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=6, pady=4, sticky="w")

    for i, (nom, val, cal) in enumerate(filas, start=1):
        ttk.Label(frame, text=nom).grid(row=i, column=0, padx=6, pady=3, sticky="w")
        ttk.Label(frame, text=val).grid(row=i, column=1, padx=6, pady=3, sticky="w")
        ttk.Label(frame, text=cal, font=("Arial", 10, "bold")).grid(row=i, column=2, padx=6, pady=3, sticky="w")

    # -------------------------------------------------
    # BOTONES
    # -------------------------------------------------
    btns = ttk.Frame(ventana3)
    btns.grid(row=2, column=0, columnspan=2, pady=10)

    def volver_a_pantalla_2():
        ventana3.destroy()
        ventana2.deiconify()

    def generar_word():
        try:
            out_path = generar_word_105_51(RESULTADOS)
            messagebox.showinfo("105/51", f"Documento Word generado correctamente:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    ttk.Button(btns, text="Volver", command=volver_a_pantalla_2).grid(row=0, column=0, padx=6)
    ttk.Button(btns, text="Generar Word", command=generar_word).grid(row=0, column=1, padx=6)

    # -------------------------------------------------
    # FOOTER
    # -------------------------------------------------
    ttk.Label(
        ventana3,
        text="Software diseñado y creado por Jesús Castellano Chillas",
        font=("Arial", 9),
        foreground="gray"
    ).grid(row=3, column=0, columnspan=2, pady=(0, 10))

    ventana2.withdraw()


# =====================================================
# PANTALLA 2 — INPUTS Y CÁLCULO
# =====================================================
def abrir_pantalla_2(vnom: float):
    global ventana2, VNOM_GLOBAL, SERIE_GLOBAL
    global entry_nombre_prueba
    global combo_serie
    global entry_r1_min, entry_r1_max
    global entry_r2_min, entry_r2_max
    global entry_desv_factor, entry_desv_lim
    global entry_pmed_lim, entry_pmed_sigma_lim, entry_pmax_lim
    global entry_trazador_lim, entry_trazador_max_fallos
    global entry_col_serie, entry_col_vel, entry_col_desv, entry_col_pmed
    global entry_col_pmax1, entry_col_pmax2
    global entry_col_espoleta, entry_col_estopin, entry_col_separacion, entry_col_vuelo, entry_col_trazador

    VNOM_GLOBAL = vnom

    ventana2 = tk.Toplevel(ventana1)
    ventana2.title("105/51 - Evaluación Automática")

    # -------------------------------------------------
    # HEADER
    # -------------------------------------------------
    header = ttk.Frame(ventana2)
    header.grid(row=0, column=0, columnspan=4, sticky="ew", pady=(10, 5), padx=10)

    try:
        img = Image.open(LOGO_PATH).resize((250, 100))
        ventana2.logo_img = ImageTk.PhotoImage(img)
        ttk.Label(header, image=ventana2.logo_img).grid(row=0, column=0, rowspan=3, padx=(0, 12))
    except Exception:
        ttk.Label(header, text="(Logo no encontrado)").grid(row=0, column=0, rowspan=3, padx=(0, 12))

    ttk.Label(header, text=TITULO_MODULO, font=("Arial", 12, "bold")).grid(row=0, column=1, sticky="w")
    ttk.Label(header, text=SUBTITULO_MODULO, font=("Arial", 10)).grid(row=1, column=1, sticky="w")
    ttk.Label(header, text=f"Velocidad nominal (Vnom): {formatear_numero(vnom)} m/s", font=("Arial", 10)).grid(
        row=2, column=1, sticky="w"
    )

    # -------------------------------------------------
    # CONTENEDOR PRINCIPAL
    # -------------------------------------------------
    main = ttk.Frame(ventana2)
    main.grid(row=1, column=0, columnspan=4, sticky="nsew", padx=10, pady=5)

    lf_inputs = ttk.LabelFrame(main, text="Inputs (manuales)")
    lf_inputs.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=5)

    lf_excel = ttk.LabelFrame(main, text="Columnas (Excel)")
    lf_excel.grid(row=0, column=1, sticky="nsew", pady=5)

    # -------------------------------------------------
    # INPUTS MANUALES
    # -------------------------------------------------
    r = 0

    ttk.Label(lf_inputs, text="Nombre de la prueba").grid(row=r, column=0, sticky="w", padx=6, pady=4)
    entry_nombre_prueba = ttk.Entry(lf_inputs, width=30)
    entry_nombre_prueba.grid(row=r, column=1, columnspan=2, sticky="w", padx=6, pady=4)

    r += 1
    ttk.Label(lf_inputs, text="Serie de temperatura").grid(row=r, column=0, sticky="w", padx=6, pady=4)
    combo_serie = ttk.Combobox(lf_inputs, values=["+21", "-35"], state="readonly", width=10)
    combo_serie.grid(row=r, column=1, sticky="w", padx=6, pady=4)
    combo_serie.set("+21")

    r += 1
    ttk.Label(lf_inputs, text="Rango UTIL-1 (min / max)").grid(row=r, column=0, sticky="w", padx=6, pady=4)
    entry_r1_min = ttk.Entry(lf_inputs, width=10)
    entry_r1_min.grid(row=r, column=1, sticky="w", padx=6, pady=4)
    entry_r1_max = ttk.Entry(lf_inputs, width=10)
    entry_r1_max.grid(row=r, column=2, sticky="w", padx=6, pady=4)

    r += 1
    ttk.Label(lf_inputs, text="Rango UTIL-2 (min / max)").grid(row=r, column=0, sticky="w", padx=6, pady=4)
    entry_r2_min = ttk.Entry(lf_inputs, width=10)
    entry_r2_min.grid(row=r, column=1, sticky="w", padx=6, pady=4)
    entry_r2_max = ttk.Entry(lf_inputs, width=10)
    entry_r2_max.grid(row=r, column=2, sticky="w", padx=6, pady=4)

    r += 1
    ttk.Label(lf_inputs, text="Factor desviación").grid(row=r, column=0, sticky="w", padx=6, pady=4)
    entry_desv_factor = ttk.Entry(lf_inputs, width=10)
    entry_desv_factor.grid(row=r, column=1, sticky="w", padx=6, pady=4)

    r += 1
    ttk.Label(lf_inputs, text="Límite desviación").grid(row=r, column=0, sticky="w", padx=6, pady=4)
    entry_desv_lim = ttk.Entry(lf_inputs, width=10)
    entry_desv_lim.grid(row=r, column=1, sticky="w", padx=6, pady=4)

    r += 1
    ttk.Label(lf_inputs, text="Pmed límite serie").grid(row=r, column=0, sticky="w", padx=6, pady=4)
    entry_pmed_lim = ttk.Entry(lf_inputs, width=12)
    entry_pmed_lim.grid(row=r, column=1, sticky="w", padx=6, pady=4)

    r += 1
    ttk.Label(lf_inputs, text="Límite Pmed + 3·σPmed").grid(row=r, column=0, sticky="w", padx=6, pady=4)
    entry_pmed_sigma_lim = ttk.Entry(lf_inputs, width=12)
    entry_pmed_sigma_lim.grid(row=r, column=1, sticky="w", padx=6, pady=4)

    r += 1
    ttk.Label(lf_inputs, text="Pmáx límite").grid(row=r, column=0, sticky="w", padx=6, pady=4)
    entry_pmax_lim = ttk.Entry(lf_inputs, width=12)
    entry_pmax_lim.grid(row=r, column=1, sticky="w", padx=6, pady=4)

    r += 1
    ttk.Label(lf_inputs, text="Límite trazador (s)").grid(row=r, column=0, sticky="w", padx=6, pady=4)
    entry_trazador_lim = ttk.Entry(lf_inputs, width=12)
    entry_trazador_lim.grid(row=r, column=1, sticky="w", padx=6, pady=4)

    r += 1
    ttk.Label(lf_inputs, text="Máx. vuelos fuera trazador").grid(row=r, column=0, sticky="w", padx=6, pady=4)
    entry_trazador_max_fallos = ttk.Entry(lf_inputs, width=12)
    entry_trazador_max_fallos.grid(row=r, column=1, sticky="w", padx=6, pady=4)

    # -------------------------------------------------
    # COLUMNAS DEL EXCEL
    # -------------------------------------------------
    e = 0

    ttk.Label(lf_excel, text="Columna Serie").grid(row=e, column=0, sticky="w", padx=6, pady=4)
    entry_col_serie = ttk.Entry(lf_excel, width=28)
    entry_col_serie.grid(row=e, column=1, sticky="w", padx=6, pady=4)

    e += 1
    ttk.Label(lf_excel, text="Columna velocidad media").grid(row=e, column=0, sticky="w", padx=6, pady=4)
    entry_col_vel = ttk.Entry(lf_excel, width=28)
    entry_col_vel.grid(row=e, column=1, sticky="w", padx=6, pady=4)

    e += 1
    ttk.Label(lf_excel, text="Columna desviación estándar").grid(row=e, column=0, sticky="w", padx=6, pady=4)
    entry_col_desv = ttk.Entry(lf_excel, width=28)
    entry_col_desv.grid(row=e, column=1, sticky="w", padx=6, pady=4)

    e += 1
    ttk.Label(lf_excel, text="Columna presión media").grid(row=e, column=0, sticky="w", padx=6, pady=4)
    entry_col_pmed = ttk.Entry(lf_excel, width=28)
    entry_col_pmed.grid(row=e, column=1, sticky="w", padx=6, pady=4)

    e += 1
    ttk.Label(lf_excel, text="Columnas Pmáx (2 columnas)").grid(row=e, column=0, sticky="w", padx=6, pady=4)

    frame_pmax = ttk.Frame(lf_excel)
    frame_pmax.grid(row=e, column=1, sticky="w", padx=6, pady=4)

    entry_col_pmax1 = ttk.Entry(frame_pmax, width=13)
    entry_col_pmax1.grid(row=0, column=0, padx=(0, 6))

    entry_col_pmax2 = ttk.Entry(frame_pmax, width=13)
    entry_col_pmax2.grid(row=0, column=1)

    e += 1
    ttk.Label(lf_excel, text="Columna Espoleta").grid(row=e, column=0, sticky="w", padx=6, pady=4)
    entry_col_espoleta = ttk.Entry(lf_excel, width=28)
    entry_col_espoleta.grid(row=e, column=1, sticky="w", padx=6, pady=4)

    e += 1
    ttk.Label(lf_excel, text="Columna Estopín").grid(row=e, column=0, sticky="w", padx=6, pady=4)
    entry_col_estopin = ttk.Entry(lf_excel, width=28)
    entry_col_estopin.grid(row=e, column=1, sticky="w", padx=6, pady=4)

    e += 1
    ttk.Label(lf_excel, text="Columna separación partes metálicas").grid(row=e, column=0, sticky="w", padx=6, pady=4)
    entry_col_separacion = ttk.Entry(lf_excel, width=28)
    entry_col_separacion.grid(row=e, column=1, sticky="w", padx=6, pady=4)

    e += 1
    ttk.Label(lf_excel, text="Columna anomalías en vuelo").grid(row=e, column=0, sticky="w", padx=6, pady=4)
    entry_col_vuelo = ttk.Entry(lf_excel, width=28)
    entry_col_vuelo.grid(row=e, column=1, sticky="w", padx=6, pady=4)

    e += 1
    ttk.Label(lf_excel, text="Columna trazador").grid(row=e, column=0, sticky="w", padx=6, pady=4)
    entry_col_trazador = ttk.Entry(lf_excel, width=28)
    entry_col_trazador.grid(row=e, column=1, sticky="w", padx=6, pady=4)

    # -------------------------------------------------
    # LÓGICA DE SERIE
    # -------------------------------------------------
    def actualizar_limite_pmed(*args):
        serie = combo_serie.get().strip()
        entry_pmed_lim.delete(0, tk.END)

        if serie == "+21":
            entry_pmed_lim.insert(0, formatear_numero(PMED_LIM_MAS_21))
        elif serie == "-35":
            entry_pmed_lim.insert(0, formatear_numero(PMED_LIM_MENOS_35))

    combo_serie.bind("<<ComboboxSelected>>", actualizar_limite_pmed)

    # -------------------------------------------------
    # BOTONES
    # -------------------------------------------------
    btn_frame = ttk.Frame(ventana2)
    btn_frame.grid(row=2, column=0, columnspan=4, pady=12)

    def accion_calcular():
        global RESULTADOS, SERIE_GLOBAL

        try:
            SERIE_GLOBAL = combo_serie.get().strip()

            datos_entrada = {
                "nombre_prueba": entry_nombre_prueba.get().strip(),
                "serie_seleccionada": SERIE_GLOBAL,

                "r1_min": convertir_a_float(entry_r1_min.get()),
                "r1_max": convertir_a_float(entry_r1_max.get()),
                "r2_min": convertir_a_float(entry_r2_min.get()),
                "r2_max": convertir_a_float(entry_r2_max.get()),

                "desv_factor": convertir_a_float(entry_desv_factor.get()),
                "desv_lim": convertir_a_float(entry_desv_lim.get()),
                "pmed_lim": convertir_a_float(entry_pmed_lim.get()),
                "pmed_sigma_lim": convertir_a_float(entry_pmed_sigma_lim.get()),
                "pmax_lim": convertir_a_float(entry_pmax_lim.get()),
                "trazador_lim": convertir_a_float(entry_trazador_lim.get()),
                "trazador_max_fallos": int(convertir_a_float(entry_trazador_max_fallos.get())),

                "col_serie": entry_col_serie.get().strip(),
                "col_vel": entry_col_vel.get().strip(),
                "col_desv": entry_col_desv.get().strip(),
                "col_pmed": entry_col_pmed.get().strip(),
                "col_pmax1": entry_col_pmax1.get().strip(),
                "col_pmax2": entry_col_pmax2.get().strip(),
                "col_espoleta": entry_col_espoleta.get().strip(),
                "col_estopin": entry_col_estopin.get().strip(),
                "col_separacion": entry_col_separacion.get().strip(),
                "col_vuelo": entry_col_vuelo.get().strip(),
                "col_trazador": entry_col_trazador.get().strip(),
            }

            if not datos_entrada["nombre_prueba"]:
                raise ValueError("Introduce un nombre de prueba (será el nombre del Word).")

            if not datos_entrada["serie_seleccionada"]:
                raise ValueError("Selecciona una serie de temperatura válida.")

            RESULTADOS = calcular_resultados_105_51(datos_entrada)

        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        abrir_pantalla_3()

    def volver_a_pantalla_1():
        ventana2.destroy()
        ventana1.deiconify()

    ttk.Button(btn_frame, text="Calcular", command=accion_calcular).grid(row=0, column=0, padx=6)
    ttk.Button(btn_frame, text="Volver", command=volver_a_pantalla_1).grid(row=0, column=1, padx=6)

    # -------------------------------------------------
    # FOOTER
    # -------------------------------------------------
    ttk.Label(
        ventana2,
        text="Software diseñado y creado por Jesús Castellano Chillas",
        font=("Arial", 9),
        foreground="gray"
    ).grid(row=3, column=0, columnspan=4, pady=(0, 10))

    # -------------------------------------------------
    # AUTORRELLENAR CAMPOS
    # -------------------------------------------------
    # UTIL-1 fijo
    entry_r1_min.insert(0, str(R1_MIN_FIJO))
    entry_r1_max.insert(0, str(R1_MAX_FIJO))

    # UTIL-2 = Vnom ± 2%
    r2_min = int(round(vnom - 0.02 * vnom))
    r2_max = int(round(vnom + 0.02 * vnom))
    entry_r2_min.insert(0, str(r2_min))
    entry_r2_max.insert(0, str(r2_max))

    # Desviación
    entry_desv_factor.insert(0, formatear_numero(DESV_FACTOR))
    entry_desv_lim.insert(0, formatear_numero(DESV_LIM))

    # Presiones
    actualizar_limite_pmed()
    entry_pmed_sigma_lim.insert(0, formatear_numero(PMED_SIGMA_LIM))
    entry_pmax_lim.insert(0, formatear_numero(PMAX_LIM))

    # Trazador
    entry_trazador_lim.insert(0, formatear_numero(TRAZADOR_LIM))
    entry_trazador_max_fallos.insert(0, str(TRAZADOR_MAX_FALLOS))

    # Columnas sugeridas por defecto
    entry_col_serie.insert(0, "Serie")
    entry_col_vel.insert(0, "velocidad_media")
    entry_col_desv.insert(0, "velocidad_media")
    entry_col_pmed.insert(0, "presion_media")
    entry_col_pmax1.insert(0, "P1")
    entry_col_pmax2.insert(0, "P2")
    entry_col_espoleta.insert(0, "Espoleta")
    entry_col_estopin.insert(0, "Estopin")
    entry_col_separacion.insert(0, "separacion_partes_metalicas")
    entry_col_vuelo.insert(0, "Vuelo")
    entry_col_trazador.insert(0, "Trazador")

    ventana1.withdraw()


# =====================================================
# PANTALLA 1 — INICIO
# =====================================================
def continuar_a_pantalla_2():
    try:
        vnom = convertir_a_float(entry_vnom.get())
    except Exception:
        messagebox.showerror("Error", "Introduce una velocidad nominal válida.")
        return

    abrir_pantalla_2(vnom)


# =====================================================
# FUNCIÓN PRINCIPAL DE LA APP
# =====================================================
def lanzar_app_105_51():
    global ventana1, entry_vnom

    ventana1 = tk.Tk()
    ventana1.title("105/51 - Inicio")
    ventana1.geometry("760x320")
    ventana1.resizable(False, False)

    # -------------------------------------------------
    # LOGO
    # -------------------------------------------------
    try:
        img = Image.open(LOGO_PATH).resize((250, 100))
        ventana1.logo_img = ImageTk.PhotoImage(img)
        ttk.Label(ventana1, image=ventana1.logo_img).grid(row=0, column=0, columnspan=2, pady=10, padx=10)
    except Exception:
        ttk.Label(ventana1, text="(Logo no encontrado)").grid(row=0, column=0, columnspan=2, pady=10)

    # -------------------------------------------------
    # TÍTULOS
    # -------------------------------------------------
    ttk.Label(ventana1, text=TITULO_MODULO, font=("Arial", 12, "bold")).grid(
        row=1, column=0, columnspan=2, pady=(0, 4)
    )
    ttk.Label(ventana1, text=SUBTITULO_MODULO, font=("Arial", 10)).grid(
        row=2, column=0, columnspan=2, pady=(0, 10)
    )

    # -------------------------------------------------
    # CONTROLES DE ENTRADA
    # -------------------------------------------------
    ttk.Label(ventana1, text="Velocidad nominal (m/s)").grid(row=3, column=0, sticky="w", padx=10, pady=6)
    entry_vnom = ttk.Entry(ventana1, width=15)
    entry_vnom.grid(row=3, column=1, sticky="w", padx=10, pady=6)

    # -------------------------------------------------
    # BOTÓN CONTINUAR
    # -------------------------------------------------
    ttk.Button(ventana1, text="Continuar", command=continuar_a_pantalla_2).grid(
        row=4, column=0, columnspan=2, pady=12
    )

    # -------------------------------------------------
    # FOOTER
    # -------------------------------------------------
    ttk.Label(
        ventana1,
        text="Software diseñado y creado por Jesús Castellano Chillas",
        font=("Arial", 9),
        foreground="gray"
    ).grid(row=5, column=0, columnspan=2, pady=(0, 10))

    ventana1.mainloop()


# =====================================================
# EJECUCIÓN DIRECTA DEL MÓDULO
# =====================================================
if __name__ == "__main__":
    lanzar_app_105_51()
