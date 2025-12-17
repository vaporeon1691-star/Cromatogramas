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
import time as time_module 
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
    if max_data < 0.5: max_data = 0.5 
    target_max = max_data * 1.1 
    ideal_step = target_max / 4.5 

    try:
        pow10 = 10**math.floor(math.log10(ideal_step)) if ideal_step > 0 else 1
    except ValueError:
        pow10 = 1 
    
    candidatos_raw = [1 * pow10, 2 * pow10, 5 * pow10]
    if ideal_step < 1: candidatos_raw.extend([0.1, 0.2, 0.5])
    if pow10 >= 10: candidatos_raw.extend([10 * pow10, 20 * pow10])

    candidatos_raw = sorted(list(set([round(c, 3) for c in candidatos_raw if c > 0])))
    candidatos_validos = [c for c in candidatos_raw if c >= ideal_step]
    paso_y = min(candidatos_validos) if candidatos_validos else 100

    limite_superior_y = math.ceil(target_max / paso_y) * paso_y
    if limite_superior_y < 5.0:
        limite_superior_y = 5.0
        paso_y = 1.0

    return limite_superior_y, paso_y

# =========================================================
# PROCESAMIENTO
# =========================================================
def procesar_archivo_local(local_filepath, t_final, hoja_leida):
    df = pd.read_excel(local_filepath, sheet_name=hoja_leida, engine="openpyxl", header=None)
    t = np.linspace(0, t_final, 15000)
    y_total = np.zeros_like(t) + 0.2 

    fila_inicio = 61
    picos_encontrados = 0
    altura_maxima_detectada = 0.0
    
    for i in range(50):
        fila_actual = fila_inicio + i
        tR = excel_a_minutos(df.iloc[fila_actual, 1])
        if tR is None: break 
        
        H = float(df.iloc[fila_actual, 9]) if pd.notna(df.iloc[fila_actual, 9]) else 0.0
        Sym = float(df.iloc[fila_actual, 14]) if pd.notna(df.iloc[fila_actual, 14]) else 1.0
        W = float(df.iloc[fila_actual, 17]) if pd.notna(df.iloc[fila_actual, 17]) else 0.0

        if H > 0:
            sigma = W / 2.355 if W > 0 else t_final / 200
            y_total += generar_pico_hplc_simetria(t, tR, sigma, H, Sym)
            picos_encontrados += 1
            if H > altura_maxima_detectada: altura_maxima_detectada = H

    ruido = np.random.normal(0, 0.15, len(t)) + (0.25 * np.sin(t * 1.5) + 0.15 * np.sin(t * 12.0))
    y_total += ruido

    plt.rcParams.update({"font.family": "sans-serif", "font.sans-serif": ["Arial"], "font.size": 8})
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.plot(t, y_total, color="#205ea6", linewidth=0.6) 

    # AJUSTE EJE Y: 5% POR DEBAJO DE CERO
    max_y_total = np.max(y_total)
    limite_superior_y, paso_y = calcular_limite_y_escalado(max_y_total)
    ax.set_ylim(-(limite_superior_y * 0.05), limite_superior_y)
    
    ticks_y = np.arange(0, limite_superior_y + paso_y, paso_y)
    ticks_y = [t for t in ticks_y if t <= limite_superior_y * 1.01]
    ax.set_yticks(ticks_y)
    
    etiquetas_y = [("mAU" if i == len(ticks_y)-1 else ("0" if v==0 else (str(int(v)) if v>=10 and float(v).is_integer() else f"{v:.1f}"))) for i, v in enumerate(ticks_y)]
    ax.set_yticklabels(etiquetas_y)

    # EJE X
    ax.set_xlim(0, t_final) 
    paso_x = 1 if t_final <= 10 else (5 if t_final <= 30 else (10 if t_final <= 60 else 20))
    ticks_x = np.arange(0, (math.ceil(t_final / paso_x) * paso_x) + 0.001, paso_x)
    ticks_filtrados = [x for x in ticks_x if x <= t_final * 1.05]
    if not ticks_filtrados or ticks_filtrados[-1] < t_final * 0.9: ticks_filtrados.append(t_final)
    ax.set_xticks(ticks_filtrados)
    ax.set_xticklabels([("min" if i == len(ticks_filtrados)-1 else (str(int(x)) if float(x).is_integer() else f"{x:.1f}")) for i, x in enumerate(ticks_filtrados)])
    
    ax.xaxis.set_minor_locator(AutoMinorLocator(5))
    ax.yaxis.set_minor_locator(AutoMinorLocator(5))
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    
    plt.tight_layout()
    return fig, picos_encontrados, altura_maxima_detectada, limite_superior_y

# =========================================================
# INTERFAZ Y LÓGICA DE ARCHIVOS
# =========================================================
def seleccionar_archivo():
    archivo_red_original = filedialog.askopenfilename(title="Selecciona Excel HPLC", filetypes=[("Excel Files", "*.xlsx *.xlsm")])
    if not archivo_red_original: return

    btn_cargar.config(text="Procesando...", state="disabled")
    root.update()
    temp_dir = tempfile.mkdtemp()
    
    try:
        local_filepath = os.path.join(temp_dir, os.path.basename(archivo_red_original))
        shutil.copy2(archivo_red_original, local_filepath)
        
        try:
            df_temp = pd.read_excel(local_filepath, sheet_name="STD VALORACIÓN Y UD", header=None)
            hoja_leida = "STD VALORACIÓN Y UD"
        except:
            df_temp = pd.read_excel(local_filepath, header=None)
            hoja_leida = df_temp.index.name or "Hoja 1"
            
        t_final = excel_a_minutos(df_temp.iloc[2, 46]) or 10
        fig, picos, alt_max, limite_y = procesar_archivo_local(local_filepath, t_final, hoja_leida)
        
        ruta_destino_png = os.path.splitext(archivo_red_original)[0] + "_cromatograma.png"
        fig.savefig(os.path.join(temp_dir, "crom.png"), dpi=300, bbox_inches='tight')
        plt.close(fig) 
        shutil.copy2(os.path.join(temp_dir, "crom.png"), ruta_destino_png)
        
        messagebox.showinfo("Éxito", f"Cromatograma generado.\n\nPicos: {picos}\nAltura Máx: {alt_max:.1f} mAU\nEscala: {limite_y} mAU")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo procesar:\n{str(e)}")
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        gc.collect()
        btn_cargar.config(text="Cargar Excel", state="normal")

# =========================================================
# MAIN - INTERFAZ GRÁFICA RENOVADA
# =========================================================
if __name__ == "__main__":
    root = tk.Tk()
    root.title("HPLC Visualizer v4.0")
    root.geometry("450x450")
    root.configure(bg="#f5f5f5")

    # Título Principal
    tk.Label(root, text="Generador de Cromatogramas", font=("Arial", 14, "bold"), bg="#f5f5f5", fg="#333").pack(pady=(20, 5))
    tk.Label(root, text="Sistema de Visualización de Datos HPLC", font=("Arial", 9), bg="#f5f5f5", fg="#666").pack()

    # Bloque de Instrucciones (Sustituye a las letras verdes)
    frame_inst = tk.LabelFrame(root, text=" Instrucciones y Consideraciones ", font=("Arial", 9, "bold"), bg="#f5f5f5", padx=15, pady=10)
    frame_inst.pack(padx=20, pady=20, fill="both")

    instrucciones = (
        "• El archivo debe ser .xlsx o .xlsm.\n"
        "• Busca la hoja: 'STD VALORACIÓN Y UD'.\n"
        "• Tiempo Final: Celda AU3 (Col 46).\n"
        "• Datos de Picos: Desde la Fila 62.\n"
        "  - Tiempo Retención: Columna B.\n"
        "  - Altura (mAU): Columna J.\n"
        "  - Simetría: Columna O.\n"
        "  - Ancho (W): Columna R.\n"
        "• El eje Y inicia un 5% debajo de cero para mejor visibilidad."
    )
    tk.Label(frame_inst, text=instrucciones, font=("Arial", 8), bg="#f5f5f5", justify="left", fg="#444").pack()

    # Botón de Carga
    btn_cargar = tk.Button(root, text="Cargar Excel", command=seleccionar_archivo, 
                           bg="#205ea6", fg="white", font=("Arial", 11, "bold"), 
                           padx=30, pady=12, cursor="hand2", relief="flat")
    btn_cargar.pack(pady=10)

    # Pie de página
    tk.Label(root, text="La imagen se guardará automáticamente junto al archivo original.", 
             font=("Arial", 7, "italic"), bg="#f5f5f5", fg="#888").pack(side="bottom", pady=15)

    root.mainloop()
