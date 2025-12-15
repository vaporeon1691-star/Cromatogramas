import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import AutoMinorLocator
from datetime import datetime, time
import os
import sys
import gc  # Garbage Collector para limpieza de memoria RAM

# =========================================================
# LÓGICA MATEMÁTICA (BLINDADA)
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

def procesar_archivo(filepath):
    # 1. LIMPIEZA PREVENTIVA (Elimina fantasmas anteriores)
    plt.close('all') 
    
    HOJA_DATOS = "STD VALORACIÓN Y UD"
    nombre_hoja_leida = ""
    
    try:
        # Intentar leer la hoja específica
        try:
            df = pd.read_excel(filepath, sheet_name=HOJA_DATOS, engine="openpyxl", header=None)
            nombre_hoja_leida = HOJA_DATOS
        except:
            # Si falla, leer la primera
            df = pd.read_excel(filepath, engine="openpyxl", header=None)
            nombre_hoja_leida = "Primera Hoja (Default)"

        # --- LECTURA DE DATOS ---
        raw_t_final = df.iloc[2, 46] # AU3
        t_final = excel_a_minutos(raw_t_final)
        if not t_final or t_final <= 0.1: t_final = 10.0

        # Eje X
        t = np.linspace(0, t_final, 15000)
        y_total = np.zeros_like(t) + 0.5 

        # BARRIDO DE PICOS
        fila_inicio = 61
        picos_encontrados = 0
        
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

        # RUIDO
        ruido_estatico = np.random.normal(0, 0.18, len(t))
        vibracion = 0.15 * np.sin(t * 12.0)
        deriva = 0.3 * np.sin(t * 0.8)
        y_total = y_total + ruido_estatico + vibracion + deriva

        # --- GRAFICADO ---
        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['font.sans-serif'] = ['Arial']
        plt.rcParams['font.size'] = 8

        fig, ax = plt.subplots(figsize=(10, 4))
        fig.patch.set_facecolor('white')
        ax.set_facecolor('white')
        ax.plot(t, y_total, color="#205ea6", linewidth=0.8)

        max_y = np.max(y_total)
        if max_y < 10: max_y = 100
        ax.set_xlim(0, t_final)
        ax.set_ylim(0, max_y * 1.1)

        mis_ticks = np.linspace(0, t_final, 7)
        ax.set_xticks(mis_ticks)
        labels_limpios = []
        for x in mis_ticks:
            if float(x).is_integer(): labels_limpios.append(f"{int(x)}")
            else: labels_limpios.append(f"{x:.1f}")
        labels_limpios[-1] = "min"
        ax.set_xticklabels(labels_limpios)

        ax.set_ylabel("mAU", loc='top', rotation=0, fontsize=8, labelpad=-20)
        ax.xaxis.set_minor_locator(AutoMinorLocator(5))
        ax.yaxis.set_minor_locator(AutoMinorLocator(5))
        ax.tick_params(which='major', direction='out', length=4, width=0.6, colors='black', top=False, right=False)
        ax.tick_params(which='minor', direction='out', length=2, width=0.5, colors='black', top=False, right=False)
        for spine in ax.spines.values():
            spine.set_linewidth(0.6)
            spine.set_color('black')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        plt.tight_layout()

        # Guardado
        output_path = os.path.splitext(filepath)[0] + "_cromatograma.png"
        plt.savefig(output_path, dpi=300, bbox_inches='tight')
        
        return output_path, nombre_hoja_leida, picos_encontrados

    except Exception as e:
        raise e
    finally:
        # LIMPIEZA TOTAL OBLIGATORIA
        plt.close('all')
        gc.collect() # Fuerza a Windows a liberar la RAM

# =========================================================
# INTERFAZ GRÁFICA (GUI)
# =========================================================
def seleccionar_archivo():
    archivo = filedialog.askopenfilename(
        title="Selecciona el archivo Excel HPLC",
        filetypes=[("Excel Files", "*.xlsm *.xlsx")]
    )
    if archivo:
        btn_cargar.config(text="Procesando...", state="disabled")
        root.update()
        try:
            ruta, hoja, picos = procesar_archivo(archivo)
            
            # Mensaje detallado para control de calidad
            mensaje = (f"✅ ¡Éxito!\n\n"
                       f"📂 Archivo: {os.path.basename(archivo)}\n"
                       f"📄 Hoja leída: {hoja}\n"
                       f"📊 Picos detectados: {picos}\n\n"
                       f"Guardado en:\n{ruta}")
            
            messagebox.showinfo("Cromatograma Generado", mensaje)
            
        except Exception as e:
            messagebox.showerror("Error Crítico", f"No se pudo procesar:\n{str(e)}")
        finally:
            btn_cargar.config(text="Cargar Excel y Generar", state="normal")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Generador HPLC v2.4 (Blindado)")
    root.geometry("400x250")
    
    lbl_instruccion = tk.Label(root, text="Selecciona tu archivo Excel", font=("Arial", 12), pady=20)
    lbl_instruccion.pack()

    btn_cargar = tk.Button(root, text="Cargar Excel y Generar", command=seleccionar_archivo, padx=20, pady=10, bg="#205ea6", fg="white", font=("Arial", 10, "bold"))
    btn_cargar.pack()

    lbl_info = tk.Label(root, text="Modo Técnico Activado\nLimpieza de memoria automática", font=("Arial", 8), fg="gray", pady=10)
    lbl_info.pack(side="bottom")

    root.mainloop()
