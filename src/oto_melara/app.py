# -*- coding: utf-8 -*-
"""
Created on Tue Apr 14 21:22:31 2026

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

# Librerías estándar
import tkinter as tk
from tkinter import ttk, messagebox

# Librería para trabajar con imágenes
from PIL import Image, ImageTk

# Importamos rutas comunes
from common.paths import LOGO_PATH

# Importamos funciones auxiliares comunes
from common.utils import convertir_a_float, formatear_numero

# Importamos la lógica de cálculo de OTO MELARA
from oto_melara.logic import calcular_resultados_oto

# Importamos la generación del Word de OTO MELARA
from oto_melara.word_generator import generar_word_oto


# =====================================================
# CONFIGURACIÓN FIJA DE OTO MELARA
# =====================================================

ARMA_NOMBRE = "OTO MELARA 76/62"

DESV_LIMITE = 5.0
PMED_LIMITE = 3285.23
PMAX_LIMITE = 3522.55


# =====================================================
# VARIABLES DE ESTADO
# =====================================================

RESULTADOS = None
VNOM_GLOBAL = None


# =====================================================
# PANTALLA 3 — CONFIRMACIÓN Y GENERACIÓN DE WORD
# =====================================================

def abrir_pantalla_3():
    global ventana3

    if RESULTADOS is None:
        messagebox.showerror("Error", "Primero debes calcular los resultados.")
        return

    ventana3 = tk.Toplevel(ventana2)
    ventana3.title("OTO MELARA - Confirmación")

    # -------------------------------------------------
    # HEADER
    # -------------------------------------------------
    header = ttk.Frame(ventana3)
    header.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(10, 5), padx=10)

    try:
        img = Image.open(LOGO_PATH).resize((250, 100))
        ventana3.logo_img = ImageTk.PhotoImage(img)
        ttk.Label(header, image=ventana3.logo_img).grid(row=0, column=0, rowspan=2, padx=(0, 12))
    except Exception:
        ttk.Label(header, text="(Logo no encontrado)").grid(row=0, column=0, rowspan=2, padx=(0, 12))

    ttk.Label(header, text=f"Arma: {ARMA_NOMBRE}", font=("Arial", 12, "bold")).grid(row=0, column=1, sticky="w")
    ttk.Label(header, text=f"Velocidad nominal (Vnom): {int(round(VNOM_GLOBAL))} m/s", font=("Arial", 11)).grid(row=1, column=1, sticky="w")

    # -------------------------------------------------
    # TABLA VISUAL DE RESULTADOS
    # -------------------------------------------------
    frame = ttk.LabelFrame(ventana3, text="Resultados calculados (confirmación)")
    frame.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

    filas = [
        ("Velocidad media", f"{formatear_numero(RESULTADOS['v_med'])} m/s", RESULTADOS["res_vel"]),
        ("Desv. estándar", f"{formatear_numero(RESULTADOS['desv'])} m/s", RESULTADOS["res_desv"]),
        ("Pmáx", f"{formatear_numero(RESULTADOS['pmax'])} bar", RESULTADOS["res_pmax"]),
        ("Pmed", f"{formatear_numero(RESULTADOS['pmed'])} bar", RESULTADOS["res_pmed"]),
        ("Espoleta", f"{RESULTADOS['fallos_espoleta']} fallos", RESULTADOS["res_espoleta"]),
        ("Estopín", f"{RESULTADOS['fallos_estopin']} fallos", RESULTADOS["res_estopin"]),
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
            out_path = generar_word_oto(RESULTADOS)
            messagebox.showinfo("OTO MELARA", f"Documento Word generado correctamente:\n{out_path}")
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

    # Ocultar pantalla 2 mientras se confirma
    ventana2.withdraw()


# =====================================================
# PANTALLA 2 — INPUTS Y CÁLCULO
# =====================================================

def abrir_pantalla_2(vnom: float):
    global ventana2, VNOM_GLOBAL
    global entry_nombre_prueba
    global entry_r1_min, entry_r1_max
    global entry_r2_min, entry_r2_max
    global entry_desv_lim, entry_pmed_lim, entry_pmax_lim
    global entry_col_vel, entry_col_desv, entry_col_pmed
    global entry_col_pmax1, entry_col_pmax2
    global entry_col_espoleta, entry_col_estopin

    VNOM_GLOBAL = vnom

    ventana2 = tk.Toplevel(ventana1)
    ventana2.title("OTO MELARA - Evaluación Automática")

    # -------------------------------------------------
    # HEADER
    # -------------------------------------------------
    header = ttk.Frame(ventana2)
    header.grid(row=0, column=0, columnspan=4, sticky="ew", pady=(10, 5), padx=10)

    try:
        img = Image.open(LOGO_PATH).resize((250, 100))
        ventana2.logo_img = ImageTk.PhotoImage(img)
        ttk.Label(header, image=ventana2.logo_img).grid(row=0, column=0, rowspan=2, padx=(0, 12))
    except Exception:
        ttk.Label(header, text="(Logo no encontrado)").grid(row=0, column=0, rowspan=2, padx=(0, 12))

    ttk.Label(header, text=f"Arma: {ARMA_NOMBRE}", font=("Arial", 12, "bold")).grid(row=0, column=1, sticky="w")
    ttk.Label(header, text=f"Velocidad nominal (Vnom): {int(round(vnom))} m/s", font=("Arial", 11)).grid(row=1, column=1, sticky="w")

    # -------------------------------------------------
    # CONTENEDOR PRINCIPAL
    # -------------------------------------------------
    main = ttk.Frame(ventana2)
    main.grid(row=1, column=0, columnspan=4, sticky="nsew", padx=10, pady=5)

    # Dos bloques visuales
    lf_inputs = ttk.LabelFrame(main, text="Inputs (manuales)")
    lf_inputs.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=5)

    lf_excel = ttk.LabelFrame(main, text="Columnas (Excel)")
    lf_excel.grid(row=0, column=1, sticky="nsew", pady=5)

    # -------------------------------------------------
    # INPUTS MANUALES
    # -------------------------------------------------
    r = 0

    ttk.Label(lf_inputs, text="Nombre de la prueba").grid(row=r, column=0, sticky="w", padx=6, pady=4)
    entry_nombre_prueba = ttk.Entry(lf_inputs, width=28)
    entry_nombre_prueba.grid(row=r, column=1, columnspan=2, sticky="w", padx=6, pady=4)

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
    ttk.Label(lf_inputs, text="Desv. estándar límite").grid(row=r, column=0, sticky="w", padx=6, pady=4)
    entry_desv_lim = ttk.Entry(lf_inputs, width=12)
    entry_desv_lim.grid(row=r, column=1, sticky="w", padx=6, pady=4)

    r += 1
    ttk.Label(lf_inputs, text="Presión media límite").grid(row=r, column=0, sticky="w", padx=6, pady=4)
    entry_pmed_lim = ttk.Entry(lf_inputs, width=12)
    entry_pmed_lim.grid(row=r, column=1, sticky="w", padx=6, pady=4)

    r += 1
    ttk.Label(lf_inputs, text="Presión individual máx. permitida").grid(row=r, column=0, sticky="w", padx=6, pady=4)
    entry_pmax_lim = ttk.Entry(lf_inputs, width=12)
    entry_pmax_lim.grid(row=r, column=1, sticky="w", padx=6, pady=4)

    # -------------------------------------------------
    # COLUMNAS DEL EXCEL
    # -------------------------------------------------
    e = 0

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
    ttk.Label(lf_excel, text="Columna Espoleta  (<2 fallos)").grid(row=e, column=0, sticky="w", padx=6, pady=4)
    entry_col_espoleta = ttk.Entry(lf_excel, width=28)
    entry_col_espoleta.grid(row=e, column=1, sticky="w", padx=6, pady=4)

    e += 1
    ttk.Label(lf_excel, text="Columna Estopín  (<1 fallos)").grid(row=e, column=0, sticky="w", padx=6, pady=4)
    entry_col_estopin = ttk.Entry(lf_excel, width=28)
    entry_col_estopin.grid(row=e, column=1, sticky="w", padx=6, pady=4)

    # -------------------------------------------------
    # BOTONES
    # -------------------------------------------------
    btn_frame = ttk.Frame(ventana2)
    btn_frame.grid(row=2, column=0, columnspan=4, pady=12)

    def accion_calcular():
        global RESULTADOS

        try:
            datos_entrada = {
                "nombre_prueba": entry_nombre_prueba.get().strip(),
                "r1_min": convertir_a_float(entry_r1_min.get()),
                "r1_max": convertir_a_float(entry_r1_max.get()),
                "r2_min": convertir_a_float(entry_r2_min.get()),
                "r2_max": convertir_a_float(entry_r2_max.get()),
                "limite_desv": convertir_a_float(entry_desv_lim.get()),
                "pmed_lim": convertir_a_float(entry_pmed_lim.get()),
                "pmax_lim": convertir_a_float(entry_pmax_lim.get()),
                "col_vel": entry_col_vel.get().strip(),
                "col_desv": entry_col_desv.get().strip(),
                "col_pmed": entry_col_pmed.get().strip(),
                "col_pmax1": entry_col_pmax1.get().strip(),
                "col_pmax2": entry_col_pmax2.get().strip(),
                "col_espoleta": entry_col_espoleta.get().strip(),
                "col_estopin": entry_col_estopin.get().strip(),
            }

            if not datos_entrada["nombre_prueba"]:
                raise ValueError("Introduce un nombre de prueba (será el nombre del Word).")

            RESULTADOS = calcular_resultados_oto(datos_entrada)

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
    footer = ttk.Label(
        ventana2,
        text="Software diseñado y creado por Jesús Castellano Chillas",
        font=("Arial", 9),
        foreground="gray"
    )
    footer.grid(row=3, column=0, columnspan=4, pady=(0, 10))

    # -------------------------------------------------
    # AUTORRELLENAR CAMPOS DE OTO
    # -------------------------------------------------
    r1_min = int(round(vnom - 7))
    r1_max = int(round(vnom + 7))
    r2_min = int(round(vnom - 0.02 * vnom))
    r2_max = int(round(vnom + 0.02 * vnom))

    entry_r1_min.insert(0, str(r1_min))
    entry_r1_max.insert(0, str(r1_max))
    entry_r2_min.insert(0, str(r2_min))
    entry_r2_max.insert(0, str(r2_max))

    entry_desv_lim.insert(0, formatear_numero(DESV_LIMITE))
    entry_pmed_lim.insert(0, formatear_numero(PMED_LIMITE))
    entry_pmax_lim.insert(0, formatear_numero(PMAX_LIMITE))

    # Ocultar pantalla 1
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

def lanzar_app_oto():
    global ventana1, entry_vnom

    ventana1 = tk.Tk()
    ventana1.title("OTO MELARA - Inicio")
    ventana1.geometry("700x300")
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
    # CONTROLES DE ENTRADA
    # -------------------------------------------------
    ttk.Label(ventana1, text="Arma").grid(row=1, column=0, sticky="w", padx=10, pady=6)
    ttk.Label(ventana1, text=ARMA_NOMBRE).grid(row=1, column=1, sticky="w", padx=10, pady=6)

    ttk.Label(ventana1, text="Velocidad nominal (m/s)").grid(row=2, column=0, sticky="w", padx=10, pady=6)
    entry_vnom = ttk.Entry(ventana1, width=15)
    entry_vnom.grid(row=2, column=1, sticky="w", padx=10, pady=6)

    # -------------------------------------------------
    # BOTÓN CONTINUAR
    # -------------------------------------------------
    ttk.Button(ventana1, text="Continuar", command=continuar_a_pantalla_2).grid(
        row=3, column=0, columnspan=2, pady=12
    )

    # -------------------------------------------------
    # FOOTER
    # -------------------------------------------------
    ttk.Label(
        ventana1,
        text="Software diseñado y creado por Jesús Castellano Chillas",
        font=("Arial", 9),
        foreground="gray"
    ).grid(row=4, column=0, columnspan=2, pady=(0, 10))

    ventana1.mainloop()
    
# =====================================================
# EJECUCIÓN DIRECTA DEL MÓDULO
# =====================================================
if __name__ == "__main__":
    lanzar_app_oto()
