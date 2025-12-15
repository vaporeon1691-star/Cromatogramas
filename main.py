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

def calcular_pasos_y(altura_max):
    """Calcula los pasos para el eje Y similar al eje X"""
    if altura_max <= 10: 
        paso = 1
    elif altura_max <= 50: 
        paso = 5
    elif altura_max <= 100: 
        paso = 10
    elif altura_max <= 200: 
        paso = 20
    elif altura_max <= 500: 
        paso = 50
    elif altura_max <= 1000: 
        paso = 100
    elif altura_max <= 2000: 
        paso = 200
    elif altura_max <= 5000: 
        paso = 500
    elif altura_max <= 10000: 
        paso = 1000
    elif altura_max <= 20000: 
        paso = 2000
    elif altura_max <= 50000: 
        paso = 5000
    else: 
        paso = 10000  # Para alturas muy grandes
    
    return paso

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
    ancho_promedio = 0.0
    simetria_promedio = 0.0
    picos_con_datos = 0
    
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
            
            altura_maxima_detectada = max(altura_maxima_detectada, H)
            if W > 0:
                ancho_promedio += W
                picos_con_datos += 1
            if Sym > 0:
                simetria_promedio += Sym
    
    if picos_con_datos > 0:
        ancho_promedio = ancho_promedio / picos_con_datos
        simetria_promedio = simetria_promedio / picos_con_datos

    # ===== RUIDO + DERIVA ESTABLE =====
    ruido_estatico = np.random.normal(0, 0.15, len(t))
    vibracion_y_deriva = 0.25 * np.sin(t * 1.5) + 0.15 * np.sin(t * 12.0)
    y_total += ruido_estatico + vibracion_y_deriva

    # GRAFICADO CON ESCALA Y MEJORADA
    plt.rcParams.update({"font.family": "sans-serif", "font.sans-serif": ["Arial"], "font.size": 8})
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.plot(t, y_total, color="#205ea6", linewidth=0.8)

    max_y = np.max(y_total)
    
    # --- ESCALA Y MEJORADA (similar al eje X) ---
    # Calcular paso apropiado para Y
    paso_y = calcular_pasos_y(max_y)
    
    # Calcular límite superior redondeado hacia arriba
    limite_superior_y = math.ceil(max_y * 1.1 / paso_y) * paso_y
    
    # Asegurar un mínimo de 100 si es muy bajo
    if limite_superior_y < 100 and max_y > 10:
        limite_superior_y = 100
    elif max_y <= 10:
        limite_superior_y = max(20, math.ceil(max_y * 1.5))
    
    ax.set_xlim(0, t_final)
    ax.set_ylim(0, limite_superior_y)
    
    # Establecer ticks principales para Y
    ticks_y = np.arange(0, limite_superior_y + paso_y, paso_y)
    # Filtrar ticks que estén dentro del límite
    ticks_y = [t for t in ticks_y if t <= limite_superior_y * 1.01]
    ax.set_yticks(ticks_y)
    
    # Formatear etiquetas Y (enteros sin decimales)
    ax.set_yticklabels([str(int(t)) if t >= 10 else f"{t:.0f}" for t in ticks_y])
    
    # --- EJE X (MANTENIENDO FORMATO ORIGINAL) ---
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

    ax.set_ylabel("mAU", loc="top", rotation=0, labelpad=-20)
    
    # SUBDIVISIONES (4 líneas intermedias como ya se viene manejando)
    ax.xaxis.set_minor_locator(AutoMinorLocator(5))  # 4 líneas menores entre ticks principales
    ax.yaxis.set_minor_locator(AutoMinorLocator(5))  # 4 líneas menores entre ticks principales
    
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    
    plt.tight_layout()
    return fig, picos_encontrados, altura_maxima_detectada, ancho_promedio, simetria_promedio, limite_superior_y

# =========================================================
# INTERFAZ
# =========================================================
def seleccionar_archivo():
    archivo_red_original = filedialog.askopenfilename(
        title="Selecciona Excel HPLC", 
        filetypes=[("Excel Files", "*.xlsx *.xlsm")]
    )
    if not archivo_red_original: 
        return

    btn_cargar.config(text="Procesando...", state="disabled")
    root.update()
    
    temp_dir = tempfile.mkdtemp()
    
    try:
        local_filepath = os.path.join(temp_dir, os.path.basename(archivo_red_original))
        shutil.copy2(archivo_red_original, local_filepath)
        
        HOJA_DATOS = "STD VALORACIÓN Y UD"
        try:
            df_temp = pd.read_excel(local_filepath, sheet_name=HOJA_DATOS, header=None)
            hoja_leida = HOJA_DATOS
        except:
            try:
                df_temp = pd.read_excel(local_filepath, header=None)
                hoja_leida = "Primera Hoja (Default)"
            except Exception as e:
                raise Exception(f"No se puede leer el archivo Excel: {str(e)}")
            
        raw_t_final = df_temp.iloc[2, 46]
        t_final = excel_a_minutos(raw_t_final) or 10
        
        fig, picos, alt_max, ancho_prom, simetria_prom, limite_y = procesar_archivo_local(
            local_filepath, t_final, hoja_leida
        )
        
        ruta_destino_png = os.path.splitext(archivo_red_original)[0] + "_cromatograma.png"
        
        fig.savefig(ruta_destino_png, dpi=300, bbox_inches='tight', metadata={'CreationDate': None})
        plt.close(fig)
        
        # --- VERIFICACIÓN ---
        max_intentos = 15
        intento = 0
        archivo_creado = False
        
        while intento < max_intentos and not archivo_creado:
            try:
                if os.path.exists(ruta_destino_png):
                    tamano = os.path.getsize(ruta_destino_png)
                    if tamano > 1000:
                        with open(ruta_destino_png, 'rb') as f:
                            if f.read(8).startswith(b'\x89PNG\r\n\x1a\n'):
                                archivo_creado = True
                                break
            except:
                pass
            
            time_module.sleep(0.3 * (intento + 1))
            intento += 1
            root.update()
        
        # --- INFORME MEJORADO ---
        mensaje = (f"✅ ¡PROCESO FINALIZADO!\n\n"
                  f"📊 DATOS UTILIZADOS:\n"
                  f"-------------------\n"
                  f"• Tiempo total: {t_final:.2f} min\n"
                  f"• Picos detectados: {picos}\n"
                  f"• Altura máxima: {alt_max:.1f} mAU\n")
        
        if ancho_prom > 0:
            mensaje += f"• Ancho promedio: {ancho_prom:.3f} min\n"
        
        if simetria_prom > 0:
            if simetria_prom < 0.9: clasif = "Fronting"
            elif simetria_prom > 1.1: clasif = "Tailing"
            else: clasif = "Normal"
            mensaje += f"• Simetría promedio: {simetria_prom:.3f} ({clasif})\n"
        
        mensaje += f"\n📈 ESCALAS APLICADAS:\n"
        mensaje += f"• Eje X: 0 a {t_final:.1f} min\n"
        mensaje += f"• Eje Y: 0 a {limite_y} mAU (autoajustado)\n"
        mensaje += f"• Subdivisiones: 4 líneas entre cada tick\n"
        
        mensaje += f"\n📁 Imagen guardada en:\n{os.path.dirname(ruta_destino_png)}"
        
        if os.path.exists(ruta_destino_png):
            messagebox.showinfo("Cromatograma Generado", mensaje)
        else:
            messagebox.showerror("Error", "No se pudo crear la imagen")
        
    except Exception as e:
        messagebox.showerror("Error Crítico", f"Fallo en el procesamiento:\n{str(e)}")
    
    finally:
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except:
            pass
        gc.collect()
        btn_cargar.config(text="Cargar Excel", state="normal")

# =========================================================
# MAIN
# =========================================================
if __name__ == "__main__":
    root = tk.Tk()
    root.title("HPLC Gen v3.5 (Escalas Mejoradas)")
    root.geometry("500x450")
    
    tk.Label(root, text="Generador de Cromatogramas HPLC", 
             font=("Arial", 14, "bold"), pady=10).pack()
    
    tk.Label(root, text="Versión 3.5 - Escalas X/Y coherentes con autoajuste", 
             font=("Arial", 9), fg="darkgreen").pack()
    
    # Información de escalas
    escalas_frame = tk.LabelFrame(root, text=" 📈 ESCALAS AUTO-AJUSTABLES", 
                                 font=("Arial", 10, "bold"), padx=10, pady=10)
    escalas_frame.pack(pady=10, padx=20, fill=tk.X)
    
    tk.Label(escalas_frame, text="• Eje X: escala de tiempo con ticks inteligentes", 
             font=("Arial", 9)).pack(anchor=tk.W, pady=2)
    tk.Label(escalas_frame, text="• Eje Y: escala de mAU autoajustable por altura", 
             font=("Arial", 9)).pack(anchor=tk.W, pady=2)
    tk.Label(escalas_frame, text="• 4 subdivisiones entre cada tick principal", 
             font=("Arial", 9)).pack(anchor=tk.W, pady=2)
    tk.Label(escalas_frame, text="• Límite Y: 10-20% sobre altura máxima detectada", 
             font=("Arial", 8), fg="blue").pack(anchor=tk.W, pady=5)
    
    btn_cargar = tk.Button(root, text="📂 Cargar Excel HPLC", 
                          command=seleccionar_archivo, 
                          padx=25, pady=15, 
                          bg="#205ea6", fg="white", 
                          font=("Arial", 12, "bold"))
    btn_cargar.pack(pady=20)
    
    estado_label = tk.Label(root, text="Estado: Listo", fg="green", font=("Arial", 9))
    estado_label.pack(pady=10)
    
    root.mainloop()
