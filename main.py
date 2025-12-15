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

# =========================================================
# LÓGICA MATEMÁTICA
# =========================================================
def excel_a_minutos(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    if isinstance(valor, (datetime, time)):
        return valor.hour * 60 + valor.minute + valor.second / 60
    return None

def generar_pico_hplc_simetria(t, tR, sigma, H, simetria):
    if simetria <= 0.001:
        simetria = 1.0
    if sigma <= 0.00001:
        sigma = 0.01

    sigma_L = 2 * sigma / (1 + simetria)
    sigma_R = simetria * sigma_L

    y = np.zeros_like(t)
    y[t <= tR] = H * np.exp(-0.5 * ((t[t <= tR] - tR) / sigma_L) ** 2)
    y[t > tR] = H * np.exp(-0.5 * ((t[t > tR] - tR) / sigma_R) ** 2)

    return y

# =========================================================
# PROCESAMIENTO
# =========================================================
def procesar_archivo_local(local_filepath, t_final, hoja_leida):
    df = pd.read_excel(local_filepath, sheet_name=hoja_leida, engine="openpyxl", header=None)

    t = np.linspace(0, t_final, 15000)
    y_total = np.zeros_like(t) + 0.5

    fila_inicio = 61
    picos = 0
    altura_max = 0.0

    for i in range(50):
        fila = fila_inicio + i
        tR = excel_a_minutos(df.iloc[fila, 1])
        if tR is None:
            break

        H = float(df.iloc[fila, 9]) if pd.notna(df.iloc[fila, 9]) else 0
        W = float(df.iloc[fila, 17]) if pd.notna(df.iloc[fila, 17]) else 0
        Sym = float(df.iloc[fila, 14]) if pd.notna(df.iloc[fila, 14]) else 1

        if H > 0:
            sigma = W / 2.355 if W > 0 else t_final / 200
            y_total += generar_pico_hplc_simetria(t, tR, sigma, H, Sym)
            picos += 1
            altura_max = max(altura_max, H)

    # ===== Ruido + deriva realista =====
    ruido = np.random.normal(0, 0.15, len(t))
    amplitud_deriva = 0.6
    tau = t_final / 3
    deriva_lenta = amplitud_deriva * (1 - np.exp(-t / tau))
    rw = np.cumsum(np.random.normal(0, 0.002, len(t)))
    rw = np.convolve(rw, np.ones(300)/300, mode="same")

    y_total += ruido + deriva_lenta + rw

    # ===== Gráfica =====
    plt.rcParams.update({
        "font.family": "sans-serif",
        "font.sans-serif": ["Arial"],
        "font.size": 8
    })

    fig, ax = plt.subplots(figsize=(10, 4))
    ax.plot(t, y_total, linewidth=0.8)

    ax.set_xlim(0, t_final)
    ax.set_ylim(0, max(100, np.max(y_total) * 1.1))

    # Escala X coherente
    if t_final <= 10:
        paso = 1
    elif t_final <= 30:
        paso = 5
    elif t_final <= 60:
        paso = 10
    else:
        paso = 20

    ticks = np.arange(0, t_final + paso, paso)
    ax.set_xticks(ticks)

    labels = [str(int(x)) if i < len(ticks)-1 else "min"
              for i, x in enumerate(ticks)]
    ax.set_xticklabels(labels)

    ax.set_ylabel("mAU", loc="top", rotation=0, labelpad=-20)
    ax.xaxis.set_minor_locator(AutoMinorLocator(5))
    ax.yaxis.set_minor_locator(AutoMinorLocator(5))
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    plt.tight_layout()
    return fig, picos, altura_max

# =========================================================
# INTERFAZ
# =========================================================
def seleccionar_archivo():
    archivo = filedialog.askopenfilename(
        title="Selecciona Excel HPLC",
        filetypes=[("Excel Files", "*.xlsx *.xlsm")]
    )
    if not archivo:
        return

    btn_cargar.config(text="Procesando...", state="disabled")
    root.update()

    temp_dir = tempfile.mkdtemp()

    try:
        local = os.path.join(temp_dir, os.path.basename(archivo))
        shutil.copy2(archivo, local)

        HOJA = "STD VALORACIÓN Y UD"
        try:
            df = pd.read_excel(local, sheet_name=HOJA, header=None)
            hoja = HOJA
        except:
            df = pd.read_excel(local, header=None)
            hoja = "Primera hoja"

        t_final = excel_a_minutos(df.iloc[2, 46]) or 10
        fig, picos, alt = procesar_archivo_local(local, t_final, hoja)

        png_temp = os.path.join(temp_dir, "crom.png")
        fig.savefig(png_temp, dpi=300, bbox_inches="tight")
        plt.close(fig)

        destino = os.path.splitext(archivo)[0] + "_cromatograma.png"
        shutil.move(png_temp, destino)

        messagebox.showinfo(
            "Proceso finalizado",
            f"✔ Cromatograma generado\n\nHoja: {hoja}\n"
            f"Picos: {picos}\nAltura máx: {alt:.1f} mAU"
        )

    except Exception as e:
        messagebox.showerror("Error", str(e))
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        gc.collect()
        btn_cargar.config(text="Cargar Excel", state="normal")

# =========================================================
# MAIN
# =========================================================
if __name__ == "__main__":
    root = tk.Tk()
    root.title("HPLC Gen v2.9 – Deriva Realista")
    root.geometry("420x320")

    tk.Label(root, text="Generador de Cromatogramas HPLC",
             font=("Arial", 12, "bold"), pady=10).pack()
    tk.Label(root, text="Deriva instrumental realista",
             font=("Arial", 9), fg="darkgreen").pack()

    btn_cargar = tk.Button(root, text="Cargar Excel",
                           command=seleccionar_archivo,
                           bg="#205ea6", fg="white",
                           font=("Arial", 11, "bold"),
                           padx=20, pady=10)
    btn_cargar.pack(pady=20)

    tk.Label(root, text="TR B62 | Altura J62 | Simetría O62 | Ancho R62",
             font=("Arial", 8), fg="gray").pack()

    root.mainloop()
