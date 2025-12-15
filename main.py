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
    mask_L = t <= tR
    mask_R = t > tR

    y[mask_L] = H * np.exp(-0.5 * ((t[mask_L] - tR) / sigma_L) ** 2)
    y[mask_R] = H * np.exp(-0.5 * ((t[mask_R] - tR) / sigma_R) ** 2)

    return y

# =========================================================
# PROCESAMIENTO PRINCIPAL
# =========================================================
def procesar_archivo_local(local_filepath, t_final, hoja_leida):
    df = pd.read_excel(local_filepath, sheet_name=hoja_leida, engine="openpyxl", header=None)

    t = np.linspace(0, t_final, 15000)
    y_total = np.zeros_like(t) + 0.5

    fila_inicio = 61
    picos_encontrados = 0
    altura_maxima_detectada = 0.0

    for i in range(50):
        fila_actual = fila_inicio + i

        dato_tR = df.iloc[fila_actual, 1]
        tR = excel_a_minutos(dato_tR)
        if tR is None:
            break

        raw_H = df.iloc[fila_actual, 9]
        raw_Sym = df.iloc[fila_actual, 14]
        raw_W = df.iloc[fila_actual, 17]

        H = float(raw_H) if pd.notna(raw_H) else 0.0
        W = float(raw_W) if pd.notna(raw_W) else 0.0
        Sym = float(raw_Sym) if pd.notna(raw_Sym) else 1.0

        if H > 0:
            sigma = W / 2.355 if W > 0 else t_final / 200
            y_total += generar_pico_hplc_simetria(t, tR, sigma, H, Sym)
            picos_encontrados += 1
            altura_maxima_detectada = max(altura_maxima_detectada, H)

    # Ruido y deriva
    y_total += np.random.normal(0, 0.18, len(t))
    y_total += 0.3 * np.sin(t * 0.8)

    # Gráfica
    plt.rcParams.update({
        "font.family": "sans-serif",
        "font.sans-serif": ["Arial"],
        "font.size": 8
    })

    fig, ax = plt.subplots(figsize=(10, 4))
    ax.plot(t, y_total, linewidth=0.8)

    ax.set_xlim(0, t_final)
    ax.set_ylim(0, max(100, np.max(y_total) * 1.1))

    ticks = np.linspace(0, t_final, 7)
    ax.set_xticks(ticks)
    labels = [f"{int(x)}" if x.is_integer() else f"{x:.1f}" for x in ticks]
    labels[-1] = "min"
    ax.set_xticklabels(labels)

    ax.set_ylabel("mAU", loc="top", rotation=0, labelpad=-20)
    ax.xaxis.set_minor_locator(AutoMinorLocator(5))
    ax.yaxis.set_minor_locator(AutoMinorLocator(5))

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    plt.tight_layout()
    return fig, picos_encontrados, altura_maxima_detectada

# =========================================================
# INTERFAZ
# =========================================================
def seleccionar_archivo():
    archivo_origen = filedialog.askopenfilename(
        title="Selecciona el archivo Excel HPLC",
        filetypes=[("Excel Files", "*.xlsx *.xlsm")]
    )
    if not archivo_origen:
        return

    btn_cargar.config(text="Procesando...", state="disabled")
    root.update()

    temp_dir = tempfile.mkdtemp()

    try:
        archivo_local = os.path.join(temp_dir, os.path.basename(archivo_origen))
        shutil.copy2(archivo_origen, archivo_local)

        HOJA = "STD VALORACIÓN Y UD"
        try:
            df_temp = pd.read_excel(archivo_local, sheet_name=HOJA, header=None)
            hoja_leida = HOJA
        except:
            df_temp = pd.read_excel(archivo_local, header=None)
            hoja_leida = "Primera hoja"

        raw_t_final = df_temp.iloc[2, 46]
        t_final = excel_a_minutos(raw_t_final) or 10.0

        fig, picos, alt_max = procesar_archivo_local(archivo_local, t_final, hoja_leida)

        # 🔒 Guardar primero en LOCAL
        ruta_png_temp = os.path.join(temp_dir, "cromatograma_temp.png")
        fig.savefig(ruta_png_temp, dpi=300, bbox_inches="tight")
        plt.close(fig)

        # 🔥 Mover al destino final (dispara refresco)
        ruta_destino = os.path.splitext(archivo_origen)[0] + "_cromatograma.png"
        shutil.move(ruta_png_temp, ruta_destino)

        messagebox.showinfo(
            "Proceso finalizado",
            f"✔ Cromatograma generado correctamente\n\n"
            f"Hoja: {hoja_leida}\n"
            f"Picos detectados: {picos}\n"
            f"Altura máxima: {alt_max:.1f} mAU\n\n"
            f"Ubicación:\n{ruta_destino}"
        )

    except Exception as e:
        messagebox.showerror("Error crítico", str(e))

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        gc.collect()
        btn_cargar.config(text="Cargar Excel", state="normal")

# =========================================================
# MAIN
# =========================================================
if __name__ == "__main__":
    root = tk.Tk()
    root.title("HPLC Gen v2.8 – Escritura Segura")
    root.geometry("420x320")

    tk.Label(root, text="Generador de Cromatogramas HPLC",
             font=("Arial", 12, "bold"), pady=10).pack()

    tk.Label(root, text="Guardado seguro para carpetas de red / OneDrive",
             font=("Arial", 9), fg="darkgreen").pack()

    btn_cargar = tk.Button(
        root,
        text="Cargar Excel",
        command=seleccionar_archivo,
        bg="#205ea6",
        fg="white",
        font=("Arial", 11, "bold"),
        padx=20,
        pady=10
    )
    btn_cargar.pack(pady=20)

    tk.Label(root,
             text="TR: B62 | Altura: J62 | Simetría: O62 | Ancho: R62",
             font=("Arial", 8), fg="gray").pack()

    root.mainloop()
