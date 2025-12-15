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
    """Calcula el límite superior y el paso de forma inteligente (5-7 divisiones) con blindaje."""
    # Blindaje: Asegurar un valor mínimo para la escala Y
    if max_data < 0.5: 
        max_data = 0.5 

    target_max = max_data * 1.1 # Margen del 10%
    ideal_step = target_max / 5.0 # Objetivo: 5 divisiones

    # Cálculo de potencia de 10
    if ideal_step <= 0: pow10 = 1
    else: pow10 = 10**math.floor(math.log10(ideal_step))
    
    # Candidatos limpios: 1x, 2x, 5x. Incluye 0.5, 0.2, 0.1 para escalas bajas
    candidatos_raw = [1 * pow10, 2 * pow10, 5 * pow10]
    
    if ideal_step < 1:
        candidatos_raw.extend([0.1, 0.2, 0.5, 1.0])
        candidatos_raw = [c for c in candidatos_raw if c > 0.001]

    # Filtro de Candidatos: Elige el paso más pequeño que es >= al ideal
    candidatos_validos = [c for c in candidatos_raw if c >= ideal_step]
    
    # --- BLINDAJE CRÍTICO (min() arg is empty fix) ---
    if not candidatos_validos:
        # Si la lista está vacía (ej. ideal_step es 60 y los candidatos son [10, 20, 50]),
        # usamos el más pequeño de los candidatos mayores que ideal_step * 2 para asegurar un paso limpio.
        paso_y = min([c for c in [10, 20, 50, 100, 200, 500, 1000] if c >= ideal_step * 2])
    else:
        paso_y = min(candidatos_validos)

    # Cálculo final
    limite_superior_y = math.ceil(target_max / paso_y) * paso_y
    
    # Ajuste final para asegurar que se muestre algo si es muy bajo
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
    y_total = np.zeros_like(t) + 0.5 

    fila_inicio = 61
    picos_encontrados = 0
    altura_maxima_detectada = 0.0
    
    # --- LECTURA DE MÚLTIPLES PICOS (RESTAURADO) ---
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

    # ===== RUIDO + DERIVA ESTABLE (Línea Base Corregida) =====
    ruido_estatico = np.random.normal(0, 0.15, len(t))
    vibracion_y_deriva = 0.25 * np.sin(t * 1.5) + 0.15 * np.sin(t * 12.0)
    y_total += ruido_estatico + vibracion_y_deriva

    # GRAFICADO
    plt.rcParams.update({"font.family": "sans-serif", "font.sans-serif": ["Arial"], "font.size": 8})
    fig, ax = plt.subplots(figsize=(10, 4))
    
    # --- GROSOR DE LÍNEA (Restaurado a Fino) ---
    ax.plot(t, y_total, color="#205ea6", linewidth=0.6) 

    # --- ESCALA Y (Blindada y Estable) ---
    max_y_total = np.max(y_total)
    limite_superior_y, paso_y = calcular_limite_y_escalado(max_y_total)

    ax.set_ylim(0, limite_superior_y)
    
    ticks_y = np.arange(0, limite_superior_y + paso_y, paso_y)
    ticks_y = [t for t in ticks_y if t <= limite_superior_y * 1.01]
    ax.set_yticks(ticks_y)
    
    # Formatear etiquetas Y
    etiquetas_y = []
    for t_val in ticks_y:
        # 1 decimal para valores muy bajos, entero para valores >= 10
        if t_val >= 10 and float(t_val).is_integer():
            etiquetas_y.append(str(int(t_val)))
        else:
            etiquetas_y.append(f"{t_val:.1f}")
    ax.set_yticklabels(etiquetas_y)

    # --- ESCALA X (Estable) ---
    ax.set_xlim(0, t_final) 
    
    if t_final <= 10: paso_x = 1
    elif t_final <= 30: paso_x = 5
    elif t_final <= 60: paso_x = 10
    else: paso_x = 20
    
    limite_superior_x_ticks = math.ceil(t_final / paso_x) * paso_x 
    ticks_x = np.arange(0, limite_superior_x_ticks + 0.001, paso_x)

    ticks_filtrados = [t for t in ticks_x if t <= t_final * 1.05]
    if not ticks_filtrados or ticks_filtrados[-1] < t_final * 0.9:
        ticks_filtrados.append(t_final)

    ax.set_xticks(ticks_filtrados)

    labels_x = []
    for i, x in enumerate(ticks_filtrados):
        if i == len(ticks_filtrados) - 1:
            labels_x.append("min")
        elif float(x).is_integer():
            labels_x.append(str(int(x)))
        else:
            labels_x.append(f"{x:.1f}")

    ax.set_xticklabels(labels_x)
    
    # --- POSICIÓN ETIQUETA "mAU" (Ajuste Final Anti-Solapamiento) ---
    ax.set_ylabel("mAU", loc="top", rotation=0, labelpad=-10) 
    
    # --- SUBDIVISIONES (Mantenidas) ---
    ax.xaxis.set_minor_locator(AutoMinorLocator(5))
    ax.yaxis.set_minor_locator(AutoMinorLocator(5))
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    
    plt.tight_layout()
    return fig, picos_encontrados, altura_maxima_detectada, limite_superior_y

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
        fig, picos, alt_max, limite_y = procesar_archivo_local(local_filepath, t_final, hoja_leida)
        
        # --- GUARDADO EN RED Y SINCRONIZACIÓN FORZADA ---
        ruta_destino_png = os.path.splitext(archivo_red_original)[0] + "_cromatograma.png"
        
        fig.savefig(ruta_destino_png, dpi=300, bbox_inches='tight')
        plt.close(fig) 

        with open(ruta_destino_png, 'ab') as f:
             os.fsync(f.fileno())
        
        time_module.sleep(1) 

        # --- INFORME FINAL ---
        mensaje = (f"✅ ¡PROCESO FINALIZADO!\n\n"
                   f"Límite de tiempo: {t_final:.2f} min\n"
                   f"Picos detectados: {picos}\n"
                   f"Altura Máx detectada: {alt_max:.1f} mAU\n"
                   f"Escala Y (Límite): {limite_y} mAU\n\n"
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
    root.title("HPLC Gen v3.5 (Estética y Multi-Pico estable)")
    root.geometry("400x320")
    
    tk.Label(root, text="Generador de Cromatogramas (Modo Estable)", font=("Arial", 12, "bold"), pady=10).pack()
    tk.Label(root, text="Blindaje contra errores de picos pequeños. Línea base más fina (0.6).", font=("Arial", 9), fg="darkgreen").pack()
    
    btn_cargar = tk.Button(root, text="Cargar Excel", command=seleccionar_archivo, padx=20, pady=10, bg="#205ea6", fg="white", font=("Arial", 11, "bold"))
    btn_cargar.pack(pady=20)

    tk.Label(root, text="Tiempo de Retención: B62 (Col 1)\nAltura Máxima: J62 (Col 9)", font=("Arial", 8), fg="gray").pack()
    
    root.mainloop()
