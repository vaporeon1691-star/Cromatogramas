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
import math # Para la función ceil

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

# =========================================================
# PROCESAMIENTO
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
        if tR is None: break 
        
        raw_H = df.iloc[fila_actual, 9]
        raw_Sym = df.iloc[fila_actual, 14]
        raw_W = df.iloc[fila_actual, 17]

        H = float(raw_H) if pd.notna(raw_H) else 0.0
        W = float(raw_W) if pd.notna(raw_W) else 0.0
        Sym = float(raw_Sym) if pd.notna(raw_Sym) else 1.0

        if H > 0:
            if W > 0: sigma = W / 2.355
            else: sigma = t_final / 200
            y_pico = generar_pico_hplc_simetria(t, tR, sigma, H, Sym)
            y_total += y_pico
            picos_encontrados += 1
            if H > altura_maxima_detectada: altura_maxima_detectada = H

    # ===== RUIDO + DERIVA ESTABLE (Línea Base Correcta) =====
    ruido_estatico = np.random.normal(0, 0.15, len(t))
    vibracion_y_deriva = 0.25 * np.sin(t * 1.5) + 0.15 * np.sin(t * 12.0)
    y_total += ruido_estatico + vibracion_y_deriva

    # GRAFICADO
    plt.rcParams.update({"font.family": "sans-serif", "font.sans-serif": ["Arial"], "font.size": 8})
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.plot(t, y_total, color="#205ea6", linewidth=0.8)

    # --- CORRECCIÓN CLAVE 1: Límite del Eje X ---
    # Forzar el límite a ser exactamente el tiempo final de AU3
    max_y = np.max(y_total)
    ax.set_xlim(0, t_final) # <-- Establece el límite estricto de la gráfica
    ax.set_ylim(0, max(100, max_y * 1.1))

    # --- CORRECCIÓN CLAVE 2: Cálculo de Ticks para Etiquetado ---
    # Calcular el final de los ticks redondeado hacia arriba para el etiquetado "min"
    
    if t_final <= 10: paso = 1
    elif t_final <= 30: paso = 5
    elif t_final <= 60: paso = 10
    else: paso = 20
    
    # Redondeamos el límite superior para asegurar que el último tick (etiqueta "min") se dibuje
    # Esto soluciona que 5.03 se extienda a 6.0
    limite_superior_ticks = math.ceil(t_final / paso) * paso 
    ticks = np.arange(0, limite_superior_ticks + 0.001, paso)

    # Eliminar el último tick si queda muy fuera del límite AU3
    ticks_filtrados = [t for t in ticks if t <= t_final * 1.05]
    if not ticks_filtrados or ticks_filtrados[-1] < t_final * 0.9:
        ticks_filtrados.append(t_final) # Asegura que el tiempo final esté marcado

    ax.set_xticks(ticks_filtrados)

    # Formateo de etiquetas (última como "min")
    labels = []
    for i, x in enumerate(ticks_filtrados):
        if i == len(ticks_filtrados) - 1:
            labels.append("min")
        elif float(x).is_integer():
            labels.append(str(int(x)))
        else:
            labels.append(f"{x:.1f}")

    ax.set_xticklabels(labels)

    ax.set_ylabel("mAU", loc="top", rotation=0, labelpad=-20)
    ax.xaxis.set_minor_locator(AutoMinorLocator(5))
    ax.yaxis.set_minor_locator(AutoMinorLocator(5))
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    
    plt.tight_layout()
    return fig, picos_encontrados, altura_maxima_detectada

# =========================================================
# INTERFAZ (Sincronización de Red)
# =========================================================
def seleccionar_archivo():
    archivo_red_original = filedialog.askopenfilename(title="Selecciona Excel HPLC", filetypes=[("Excel Files", "*.xlsx *.xlsm")])
    if not archivo_red_original: return

    btn_cargar.config(text="Procesando...", state="disabled")
    root.update()
    
    temp_dir = tempfile.mkdtemp()
    
    try:
        # --- LECTURA INICIAL Y COPIA LOCAL ---
        local_filepath = os.path.join(temp_dir, os.path.basename(archivo_red_original))
        shutil.copy2(archivo_red_original, local_filepath)
        
        HOJA_DATOS = "STD VALORACIÓN Y UD"
        try:
            df_temp = pd.read_excel(local_filepath, sheet_name=HOJA_DATOS, header=None)
            hoja_leida = HOJA_DATOS
        except:
            df_temp = pd.read_excel(local_filepath, header=None)
            hoja_leida = "Primera Hoja (Default)"
            
        raw_t_final = df_temp.iloc[2, 46]
        t_final = excel_a_minutos(raw_t_final) or 10
        
        # --- PROCESAMIENTO ---
        fig, picos, alt_max = procesar_archivo_local(local_filepath, t_final, hoja_leida)
        
        # --- GUARDADO EN RED Y SINCRONIZACIÓN FORZADA ---
        ruta_destino_png = os.path.splitext(archivo_red_original)[0] + "_cromatograma.png"
        
        fig.savefig(ruta_destino_png, dpi=300, bbox_inches='tight')
        plt.close(fig) 

        # Forzar la escritura completa al servidor
        with open(ruta_destino_png, 'ab') as f:
             os.fsync(f.fileno())
        
        # PAUSA OBLIGATORIA
        time_module.sleep(1) 

        # --- INFORME FINAL ---
        mensaje = (f"✅ ¡PROCESO FINALIZADO!\n\n"
                   f"Límite de tiempo: {t_final:.2f} min\n"
                   f"Línea Base: Corregida.\n"
                   f"Picos detectados: {picos}\n"
                   f"Busca la imagen en la misma carpeta que el Excel.")
        
        messagebox.showinfo("Cromatograma Generado", mensaje)
        
    except Exception as e:
        messagebox.showerror("Error Crítico", f"Fallo en el procesamiento:\n{str(e)}")
    
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        gc.collect()
        btn_cargar.config(text="Cargar Excel", state="normal")

# =========================================================
# MAIN
# =========================================================
if __name__ == "__main__":
    root = tk.Tk()
    root.title("HPLC Gen v3.1 (Ajuste Final de Eje X)")
    root.geometry("400x320")
    
    tk.Label(root, text="Generador de Cromatogramas (Modo Estable)", font=("Arial", 12, "bold"), pady=10).pack()
    tk.Label(root, text="Línea base y límites de tiempo (AU3) corregidos.", font=("Arial", 9), fg="darkgreen").pack()
    
    btn_cargar = tk.Button(root, text="Cargar Excel", command=seleccionar_archivo, padx=20, pady=10, bg="#205ea6", fg="white", font=("Arial", 11, "bold"))
    btn_cargar.pack(pady=20)

    tk.Label(root, text="Tiempo de Retención: B62 (Col 1)\nAltura Máxima: J62 (Col 9)", font=("Arial", 8), fg="gray").pack()
    
    root.mainloop()
