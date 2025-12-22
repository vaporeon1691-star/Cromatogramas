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
    # Aseguramos un mínimo para evitar errores con picos vacíos
    if max_data < 0.5: max_data = 0.5 
    target_max = max_data * 1.05 # Margen del 5%

    # --- LÓGICA DE 6 NIVELES ---
    # Queremos dividir el rango en aprox 6 partes.
    ideal_step = target_max / 6.0

    # Calculamos la magnitud del paso (potencia de 10)
    exponent = math.floor(math.log10(ideal_step)) if ideal_step > 0 else 0
    fraction = ideal_step / (10**exponent) # Normalizamos a un valor entre 1 y 10

    # Seleccionamos el paso estándar más cercano (1, 2, 5)
    # Usamos medias geométricas para decidir el corte óptimo
    if fraction < 1.42:  # sqrt(2) aprox
        step_base = 1
    elif fraction < 3.16: # sqrt(10) aprox
        step_base = 2
    elif fraction < 7.07: # sqrt(50) aprox
        step_base = 5
    else:
        step_base = 10
    
    paso_y = step_base * (10**exponent)
    
    # Calculamos el límite superior basado en este paso
    limite_superior_y = math.ceil(target_max / paso_y) * paso_y
    
    # Corrección de seguridad para evitar gráficas planas
    if limite_superior_y <= 0: limite_superior_y = 1.0 * paso_y

    return limite_superior_y, paso_y

# =========================================================
# PROCESAMIENTO
# =========================================================
def procesar_archivo_local(local_filepath, t_final, hoja_leida):
    df = pd.read_excel(local_filepath, sheet_name=hoja_leida, engine="openpyxl", header=None)
    
    t = np.linspace(0, t_final, 15000)
    y_total = np.zeros_like(t) + 0.8 

    fila_inicio = 61
    picos_encontrados = 0
    altura_maxima_detectada = 0.0
    
    lista_picos = [] 
    
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
            
            # Integración a línea base
            ancho_ref = W if W > 0 else 0.5
            inicio_pico = tR - (ancho_ref * 1.7)
            fin_pico = tR + (ancho_ref * 1.7)

            lista_picos.append({
                'tR': tR,
                'inicio': inicio_pico,
                'fin': fin_pico
            })

    # Ruido
    ruido_base = np.random.normal(0, 0.55, len(t)) 
    ruido_vibracion = (np.random.rand(len(t)) - 0.5) * 0.8
    ruido_ondas = (0.5 * np.sin(t * 2.0) + 0.2 * np.sin(t * 20.0))
    y_total += (ruido_base + ruido_vibracion + ruido_ondas)

    # Configuración Gráfica
    plt.rcParams.update({
        "font.family": "sans-serif", 
        "font.sans-serif": ["Arial"], 
        "font.size": 12,      
        "axes.linewidth": 0.9
    })
    
    fig, ax = plt.subplots(figsize=(14.72, 6.93), dpi=100)
    
    ax.plot(t, y_total, color="#205ea6", linewidth=0.7, zorder=2) 

    # Marcas e integración
    for pico in lista_picos:
        try:
            idx_ini = (np.abs(t - pico['inicio'])).argmin()
            idx_fin = (np.abs(t - pico['fin'])).argmin()
            val_y_ini = y_total[idx_ini]
            val_y_fin = y_total[idx_fin]
            
            ax.plot([t[idx_ini], t[idx_fin]], [val_y_ini, val_y_fin], 
                    color="red", linewidth=1.1, linestyle="-", zorder=3)
            
            tick_size = altura_maxima_detectada * 0.03 
            if tick_size < 1.5: tick_size = 1.5
            if tick_size > 50: tick_size = 50 # Aumentado el tope para picos gigantes

            ax.plot([t[idx_ini], t[idx_ini]], [val_y_ini, val_y_ini + tick_size], 
                    color="black", linewidth=1.3, zorder=4)
            ax.plot([t[idx_fin], t[idx_fin]], [val_y_fin, val_y_fin - tick_size], 
                    color="black", linewidth=1.3, zorder=4)
                    
        except Exception:
            pass

    # --- EJE Y: DINÁMICO (Target ~6 divisiones) ---
    max_y_total = np.max(y_total)
    limite_superior_y, paso_y = calcular_limite_y_escalado(max_y_total)
    ax.set_ylim(-(limite_superior_y * 0.02), limite_superior_y) 
    
    ticks_y = np.arange(0, limite_superior_y + paso_y, paso_y)
    ticks_y = [t for t in ticks_y if t <= limite_superior_y * 1.001] 
    ax.set_yticks(ticks_y)
    
    # Formato inteligente de etiquetas (sin decimales si es entero)
    etiquetas_y = [("mAU" if i == len(ticks_y)-1 else ("0" if v==0 else (str(int(v)) if v>=10 and float(v).is_integer() else f"{v:.1f}"))) for i, v in enumerate(ticks_y)]
    ax.set_yticklabels(etiquetas_y)

    # --- EJE X: DINÁMICO (Target ~10 divisiones) ---
    ax.set_xlim(0, t_final) 
    
    # Calculamos paso ideal
    ideal_step_x = t_final / 10.0
    
    # Buscamos el paso estándar más cercano (1, 2, 5, 10...)
    exponent_x = math.floor(math.log10(ideal_step_x)) if ideal_step_x > 0 else 0
    fraction_x = ideal_step_x / (10**exponent_x)
    
    if fraction_x < 1.42: paso_base_x = 1
    elif fraction_x < 3.16: paso_base_x = 2
    elif fraction_x < 7.07: paso_base_x = 5
    else: paso_base_x = 10
    
    paso_x = paso_base_x * (10**exponent_x)
    
    # Generamos ticks
    ticks_x = np.arange(0, math.ceil(t_final / paso_x) * paso_x + 0.001, paso_x)
    ticks_filtrados = [x for x in ticks_x if x <= t_final * 1.02]
    
    # Aseguramos tick final si falta
    if ticks_filtrados and (t_final - ticks_filtrados[-1]) > (paso_x * 0.6):
         ticks_filtrados.append(int(t_final) if int(t_final)==t_final else t_final)

    ax.set_xticks(ticks_filtrados)
    
    # Etiquetas eje X
    etiquetas_x = [("min" if i == len(ticks_filtrados)-1 else (str(int(x)) if float(x).is_integer() else f"{x:.1f}")) for i, x in enumerate(ticks_filtrados)]
    ax.set_xticklabels(etiquetas_x)
    
    # Ajustes visuales finales
    ax.tick_params(axis='both', which='major', width=1.0, length=5, labelsize=12)
    ax.tick_params(axis='both', which='minor', width=0.7, length=3)

    ax.xaxis.set_minor_locator(AutoMinorLocator(5))
    ax.yaxis.set_minor_locator(AutoMinorLocator(5)) 
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color('black')
    ax.spines["bottom"].set_color('black')
    
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
            hoja_leida = "Hoja 1"
            
        t_final = excel_a_minutos(df_temp.iloc[2, 46]) or 10
        fig, picos, alt_max, limite_y = procesar_archivo_local(local_filepath, t_final, hoja_leida)
        
        ruta_destino_png = os.path.splitext(archivo_red_original)[0] + "_cromatograma.png"
        
        fig.savefig(os.path.join(temp_dir, "crom.png"), dpi=100)
        plt.close(fig) 
        shutil.copy2(os.path.join(temp_dir, "crom.png"), ruta_destino_png)
        
        messagebox.showinfo("Éxito", f"Cromatograma generado.\n\nDimensiones: 1472x693\nPicos: {picos}\nEscala: {limite_y} mAU")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo procesar:\n{str(e)}")
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
        gc.collect()
        btn_cargar.config(text="Cargar Excel", state="normal")

# =========================================================
# MAIN - INTERFAZ GRÁFICA
# =========================================================
if __name__ == "__main__":
    root = tk.Tk()
    root.title("HPLC Visualizer v4.7 (Escalado Inteligente)")
    root.geometry("450x450")
    root.configure(bg="#f5f5f5")

    tk.Label(root, text="Generador de Cromatogramas", font=("Arial", 14, "bold"), bg="#f5f5f5", fg="#333").pack(pady=(20, 5))
    tk.Label(root, text="Sistema de Visualización de Datos HPLC", font=("Arial", 9), bg="#f5f5f5", fg="#666").pack()

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
        "• Salida: 1472 x 693 px.\n"
        "• Eje Y: Aprox 6 niveles. Eje X: Aprox 10 divisiones."
    )
    tk.Label(frame_inst, text=instrucciones, font=("Arial", 8), bg="#f5f5f5", justify="left", fg="#444").pack()

    btn_cargar = tk.Button(root, text="Cargar Excel", command=seleccionar_archivo, 
                           bg="#205ea6", fg="white", font=("Arial", 11, "bold"), 
                           padx=30, pady=12, cursor="hand2", relief="flat")
    btn_cargar.pack(pady=10)

    tk.Label(root, text="La imagen se guardará automáticamente junto al archivo original.", 
             font=("Arial", 7, "italic"), bg="#f5f5f5", fg="#888").pack(side="bottom", pady=15)

    root.mainloop()
