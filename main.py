import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import AutoMinorLocator
from datetime import datetime, time
import os
import gc
import shutil
import tempfile
import math

# =========================================================
# FUNCIONES MATEMÁTICAS
# =========================================================
def excel_a_minutos(valor):
    if pd.isna(valor): return None
    if isinstance(valor, (int, float)): return float(valor)
    if isinstance(valor, (datetime, time)):
        return valor.hour * 60 + valor.minute + valor.second / 60
    return None

def generar_pico_hplc_simetria(t, tR, sigma, H, simetria):
    if simetria <= 0.001: simetria = 1.0
    if sigma <= 0.00001: sigma = 0.01
    sigma_L = 2 * sigma / (1 + simetria)
    sigma_R = simetria * sigma_L
    y = np.zeros_like(t)
    y[t <= tR] = H * np.exp(-0.5 * ((t[t <= tR] - tR) / sigma_L) ** 2)
    y[t > tR] = H * np.exp(-0.5 * ((t[t > tR] - tR) / sigma_R) ** 2)
    return y

def calcular_limite_y_escalado(max_data):
    target_max = max(max_data * 1.1, 5)
    ideal_step = target_max / 4.5
    pow10 = 10**math.floor(math.log10(ideal_step))
    candidatos = sorted(set([1, 2, 5, 10]))
    paso_y = min([c * pow10 for c in candidatos if c * pow10 >= ideal_step])
    limite = math.ceil(target_max / paso_y) * paso_y
    return limite, paso_y

# =========================================================
# PROCESAMIENTO
# =========================================================
def procesar_archivo_local(local_filepath, t_final, hoja_leida):
    df = pd.read_excel(local_filepath, sheet_name=hoja_leida, header=None)
    t = np.linspace(0, t_final, 15000)
    y_total = np.zeros_like(t) + 0.2

    fila_inicio = 61
    altura_maxima_detectada = 0
    picos_encontrados = 0

    for i in range(50):
        fila = fila_inicio + i
        tR = excel_a_minutos(df.iloc[fila, 1])
        if tR is None: break

        H = float(df.iloc[fila, 9]) if pd.notna(df.iloc[fila, 9]) else 0
        Sym = float(df.iloc[fila, 14]) if pd.notna(df.iloc[fila, 14]) else 1
        W = float(df.iloc[fila, 17]) if pd.notna(df.iloc[fila, 17]) else 0

        if H > 0:
            sigma = W / 2.355 if W > 0 else t_final / 200
            y_total += generar_pico_hplc_simetria(t, tR, sigma, H, Sym)
            altura_maxima_detectada = max(altura_maxima_detectada, H)
            picos_encontrados += 1

    # ===== RUIDO HPLC REALISTA =====
    ruido_blanco = np.random.normal(0, 0.35, len(t))
    deriva_lenta = 0.35 * np.sin(t * 0.25)
    oscilacion_bomba = 0.18 * np.sin(t * 6.0)
    micro_ruido = np.random.normal(0, 0.05, len(t))

    y_total += ruido_blanco + deriva_lenta + oscilacion_bomba + micro_ruido

    plt.rcParams.update({"font.family": "Arial", "font.size": 8})
    fig, ax = plt.subplots(figsize=(14.72, 6.93), dpi=100)

    ax.plot(t, y_total, color="#205ea6", linewidth=0.8)

    # ===== MARCADO DE BASE DEL PICO =====
    idx_max = np.argmax(y_total)
    y_base = np.percentile(y_total, 5)

    i_ini = idx_max
    while i_ini > 0 and y_total[i_ini] > y_base:
        i_ini -= 1

    i_fin = idx_max
    while i_fin < len(y_total)-1 and y_total[i_fin] > y_base:
        i_fin += 1

    ax.hlines(y_base, t[i_ini], t[i_fin], colors="black", linewidth=0.8)
    ax.vlines([t[i_ini], t[i_fin]], y_base, y_total[idx_max]*0.05, colors="black", linewidth=0.8)

    # ===== EJES =====
    limite_y, paso_y = calcular_limite_y_escalado(np.max(y_total))
    ax.set_ylim(-limite_y*0.01, limite_y)
    ax.set_yticks(np.arange(0, limite_y+paso_y, paso_y))
    ax.set_yticklabels([("mAU" if i == len(ax.get_yticks())-1 else int(v)) for i, v in enumerate(ax.get_yticks())])

    ax.set_xlim(0, t_final)
    ticks_x = np.arange(0, math.ceil(t_final)+1, 1)
    ax.set_xticks(ticks_x)
    ax.set_xticklabels([("min" if i == len(ticks_x)-1 else int(x)) for i, x in enumerate(ticks_x)])

    ax.xaxis.set_minor_locator(AutoMinorLocator(5))
    ax.yaxis.set_minor_locator(AutoMinorLocator(5))
    ax.tick_params(axis='both', which='major', labelsize=9)
    ax.tick_params(axis='both', which='minor', length=2)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    plt.tight_layout()
    return fig, picos_encontrados, altura_maxima_detectada, limite_y

# =========================================================
# INTERFAZ
# =========================================================
def seleccionar_archivo():
    archivo = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xlsm")])
    if not archivo: return

    temp_dir = tempfile.mkdtemp()
    try:
        local = os.path.join(temp_dir, os.path.basename(archivo))
        shutil.copy2(archivo, local)

        try:
            df = pd.read_excel(local, sheet_name="STD VALORACIÓN Y UD", header=None)
            hoja = "STD VALORACIÓN Y UD"
        except:
            df = pd.read_excel(local, header=None)
            hoja = 0

        t_final = excel_a_minutos(df.iloc[2, 46]) or 10
        fig, picos, _, escala = procesar_archivo_local(local, t_final, hoja)

        salida = os.path.splitext(archivo)[0] + "_cromatograma.png"
        fig.savefig(salida, dpi=100)
        plt.close(fig)

        messagebox.showinfo("Listo", f"Cromatograma generado\nPicos: {picos}\nEscala: {escala} mAU")
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        gc.collect()

# =========================================================
# MAIN
# =========================================================
root = tk.Tk()
root.title("HPLC Visualizer v4.2")
root.geometry("420x350")

tk.Label(root, text="Generador de Cromatogramas", font=("Arial", 14, "bold")).pack(pady=20)
tk.Button(root, text="Cargar Excel", command=seleccionar_archivo,
          bg="#205ea6", fg="white", font=("Arial", 11, "bold"),
          padx=25, pady=10).pack()

root.mainloop()
