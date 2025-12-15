import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import AutoMinorLocator
from datetime import datetime, time
import os
import gc
import shutil # Nuevo: Para copiar/mover archivos
import tempfile # Nuevo: Para usar carpetas temporales de Windows

# =========================================================
# LÓGICA MATEMÁTICA
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
    mask_L = t <= tR
    y[mask_L] = H * np.exp(-0.5 * ((t[mask_L] - tR) / sigma_L)**2)
    mask_R = t > tR
    y[mask_R] = H * np.exp(-0.5 * ((t[mask_R] - tR) / sigma_R)**2)
    return y

def procesar_archivo_local(local_filepath, t_final, hoja_leida):
    """
    Función que realiza todo el proceso de cálculo y graficado
    usando exclusivamente el archivo en la ruta local.
    """
    df = pd.read_excel(local_filepath, sheet_name=hoja_leida, engine="openpyxl", header=None)
    
    # Eje X
    t = np.linspace(0, t_final, 15000)
    y_total = np.zeros_like(t) + 0.5 

    # BARRIDO DE PICOS
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

    # RUIDO
    ruido = np.random.normal(0, 0.18, len(t))
    deriva = 0.3 * np.sin(t * 0.8)
    y_total = y_total + ruido + deriva

    # GRAFICADO
    plt.rcParams['font.family'] = 'sans-serif'
    plt.rcParams['font.sans-serif'] = ['Arial']
    plt.rcParams['font.size'] = 8

    fig = plt.figure(figsize=(10, 4))
    ax = fig.add_subplot(111)
    
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')
    ax.plot(t, y_total, color="#205ea6", linewidth=0.8)

    max_y = np.max(y_total)
    if max_y < 10: max_y = 100
    ax.set_xlim(0, t_final)
    ax.set_ylim(0, max_y * 1.1)

    mis_ticks = np.linspace(0, t_final, 7)
    ax.set_xticks(mis_ticks)
    labels = [f"{int(x)}" if float(x).is_integer() else f"{x:.1f}" for x in mis_ticks]
    labels[-1] = "min"
    ax.set_xticklabels(labels)

    ax.set_ylabel("mAU", loc='top', rotation=0, fontsize=8, labelpad=-20)
    ax.xaxis.set_minor_locator(AutoMinorLocator(5))
    ax.yaxis.set_minor_locator(AutoMinorLocator(5))
    ax.tick_params(which='major', direction='out', length=4, width=0.6, colors='black')
    ax.tick_params(which='minor', direction='out', length=2, width=0.5, colors='black')
    for spine in ax.spines.values():
        spine.set_linewidth(0.6)
        spine.set_color('black')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    
    plt.tight_layout()
    
    # Devolvemos la figura y los datos para guardar en la ruta final
    return fig, picos_encontrados, altura_maxima_detectada

def seleccionar_archivo():
    archivo_red_original = filedialog.askopenfilename(
        title="Selecciona el archivo Excel HPLC",
        filetypes=[("Excel Files", "*.xlsm *.xlsx")]
    )
    if not archivo_red_original: return

    btn_cargar.config(text="Procesando...", state="disabled")
    root.update()
    
    # Variables de control
    ruta_temporal_completa = None
    ruta_destino_png = None

    try:
        # --- PARTE 1: COPIAR A RUTA TEMPORAL (VELOCIDAD) ---
        
        # 1. Crear carpeta temporal segura de Windows
        temp_dir = tempfile.mkdtemp()
        nombre_archivo = os.path.basename(archivo_red_original)
        ruta_temporal_completa = os.path.join(temp_dir, nombre_archivo)
        
        # 2. Copiar el archivo de la red a la PC local (¡Rápido!)
        shutil.copy2(archivo_red_original, ruta_temporal_completa)
        
        # 3. Leer el tiempo final y determinar la hoja
        HOJA_DATOS = "STD VALORACIÓN Y UD"
        try:
            df_temp = pd.read_excel(ruta_temporal_completa, sheet_name=HOJA_DATOS, header=None)
            hoja_leida = HOJA_DATOS
        except:
            df_temp = pd.read_excel(ruta_temporal_completa, header=None)
            hoja_leida = "Primera Hoja (Default)"
            
        raw_t_final = df_temp.iloc[2, 46]
        t_final = excel_a_minutos(raw_t_final)
        if not t_final or t_final <= 0.1: t_final = 10.0
        
        # --- PARTE 2: PROCESAR EN LOCAL (MÁXIMO RENDIMIENTO) ---
        fig, picos, alt_max = procesar_archivo_local(ruta_temporal_completa, t_final, hoja_leida)
        
        # 4. Determinar la ruta final del PNG
        ruta_destino_png = os.path.splitext(archivo_red_original)[0] + "_cromatograma.png"
        
        # 5. Guardar la figura en la RUTA FINAL (Red)
        fig.savefig(ruta_destino_png, dpi=300, bbox_inches='tight')
        plt.close(fig) # Cierre definitivo

        # --- PARTE 3: INFORMAR Y LIMPIAR ---
        
        mensaje = (f"✅ ¡PROCESO FINALIZADO EN SEGUNDOS!\n\n"
                   f"Tiempo de proceso (estimado): < 5 segundos\n"
                   f"Hoja leída: {hoja_leida}\n"
                   f"Picos detectados: {picos}\n"
                   f"Altura Máx: {alt_max:.1f} mAU\n\n"
                   f"Imagen guardada en:\n{ruta_destino_png}")
        
        messagebox.showinfo("Cromatograma Generado", mensaje)
        
    except Exception as e:
        messagebox.showerror("Error Crítico", f"Fallo en el procesamiento o la ruta de red:\n{str(e)}")
    
    finally:
        # 6. Limpieza de la carpeta temporal (obligatorio)
        if ruta_temporal_completa and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)
        gc.collect()
        btn_cargar.config(text="Cargar Excel y Generar", state="normal")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("HPLC Gen v2.6 (Aislamiento de Red)")
    root.geometry("400x320")
    
    tk.Label(root, text="Generador de Cromatogramas (Modo Rápido)", font=("Arial", 12, "bold"), pady=10).pack()
    tk.Label(root, text="IMPORTANTE: El archivo será procesado localmente para máxima velocidad.", font=("Arial", 9), fg="darkgreen").pack()
    
    btn_cargar = tk.Button(root, text="Cargar Excel y Generar", command=seleccionar_archivo, padx=20, pady=10, bg="#205ea6", fg="white", font=("Arial", 11, "bold"))
    btn_cargar.pack(pady=20)

    tk.Label(root, text="Tiempo de Retención: B62 (Col 1)\nAltura Máxima: J62 (Col 9)", font=("Arial", 8), fg="gray").pack()
    
    root.mainloop()
