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
    max_y = np.max(y_total)
    ax.set_xlim(0, t_final)
    ax.set_ylim(0, max(100, max_y * 1.1))

    # --- CORRECCIÓN CLAVE 2: Cálculo de Ticks para Etiquetado ---
    if t_final <= 10: paso = 1
    elif t_final <= 30: paso = 5
    elif t_final <= 60: paso = 10
    else: paso = 20
    
    limite_superior_ticks = math.ceil(t_final / paso) * paso 
    ticks = np.arange(0, limite_superior_ticks + 0.001, paso)

    ticks_filtrados = [t for t in ticks if t <= t_final * 1.05]
    if not ticks_filtrados or ticks_filtrados[-1] < t_final * 0.9:
        ticks_filtrados.append(t_final)

    ax.set_xticks(ticks_filtrados)

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
        # --- LECTURA INICIAL Y COPIA LOCAL ---
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
        
        # --- PROCESAMIENTO ---
        fig, picos, alt_max = procesar_archivo_local(local_filepath, t_final, hoja_leida)
        
        # --- GUARDADO EN RED ---
        ruta_destino_png = os.path.splitext(archivo_red_original)[0] + "_cromatograma.png"
        
        # Guardar la figura
        fig.savefig(ruta_destino_png, dpi=300, bbox_inches='tight', metadata={'CreationDate': None})
        plt.close(fig)
        
        # --- VERIFICACIÓN ACTIVA DE CREACIÓN DEL ARCHIVO ---
        max_intentos = 15  # Aumentado para redes lentas
        intento = 0
        archivo_creado = False
        tamano_archivo = 0
        
        while intento < max_intentos and not archivo_creado:
            # Intentar abrir y leer el archivo
            try:
                if os.path.exists(ruta_destino_png):
                    tamano_actual = os.path.getsize(ruta_destino_png)
                    # Verificar que el archivo tenga contenido válido (más de 1KB)
                    if tamano_actual > 1000:
                        # Verificar si el tamaño se estabilizó (no está en proceso de escritura)
                        if tamano_archivo == tamano_actual:
                            # Intentar leer el archivo para asegurar que esté disponible
                            with open(ruta_destino_png, 'rb') as f:
                                header = f.read(8)  # Leer cabecera PNG
                                if header.startswith(b'\x89PNG\r\n\x1a\n'):
                                    archivo_creado = True
                                    break
                        tamano_archivo = tamano_actual
            except (IOError, OSError, PermissionError):
                pass  # El archivo aún no está listo o está bloqueado
            
            # Esperar progresivamente más tiempo
            tiempo_espera = 0.3 * (intento + 1)
            time_module.sleep(tiempo_espera)
            intento += 1
            root.update()  # Mantener la interfaz responsiva
        
        # --- VERIFICACIÓN FINAL ---
        archivo_existe = os.path.exists(ruta_destino_png)
        archivo_valido = False
        
        if archivo_existe:
            try:
                tamano = os.path.getsize(ruta_destino_png)
                if tamano > 1000:
                    with open(ruta_destino_png, 'rb') as f:
                        if f.read(8).startswith(b'\x89PNG\r\n\x1a\n'):
                            archivo_valido = True
            except:
                pass
        
        # --- INFORME FINAL ---
        mensaje_base = (f"✅ ¡PROCESO FINALIZADO!\n\n"
                       f"Límite de tiempo: {t_final:.2f} min\n"
                       f"Línea Base: Corregida.\n"
                       f"Picos detectados: {picos}\n"
                       f"Altura máxima: {alt_max:.1f} mAU\n\n")
        
        if archivo_valido:
            mensaje = mensaje_base + f"📍 Imagen guardada en:\n{os.path.dirname(ruta_destino_png)}\n\n📄 Archivo: {os.path.basename(ruta_destino_png)}"
            messagebox.showinfo("Cromatograma Generado", mensaje)
        elif archivo_existe:
            mensaje = mensaje_base + f"⚠️ La imagen se guardó pero puede estar incompleta.\nRuta: {ruta_destino_png}"
            messagebox.showwarning("Advertencia", mensaje)
        else:
            mensaje = mensaje_base + f"❌ No se pudo crear la imagen.\nVerifique permisos de escritura en:\n{os.path.dirname(ruta_destino_png)}"
            messagebox.showerror("Error", mensaje)
        
    except Exception as e:
        messagebox.showerror("Error Crítico", f"Fallo en el procesamiento:\n{str(e)}")
    
    finally:
        # Limpieza
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
    root.title("HPLC Gen v3.2 (Corrección de Creación de Imagen)")
    root.geometry("450x350")
    
    # Estilo mejorado
    tk.Label(root, text="Generador de Cromatogramas", 
             font=("Arial", 14, "bold"), pady=10).pack()
    
    tk.Label(root, text="Versión 3.2 - Creación instantánea de imagen", 
             font=("Arial", 9), fg="darkgreen").pack()
    
    tk.Label(root, text="Soluciona el problema de retraso en la generación", 
             font=("Arial", 8), fg="gray").pack(pady=5)
    
    btn_cargar = tk.Button(root, text="📂 Cargar Excel HPLC", 
                          command=seleccionar_archivo, 
                          padx=25, pady=12, 
                          bg="#205ea6", fg="white", 
                          font=("Arial", 11, "bold"),
                          activebackground="#1a4d8c",
                          cursor="hand2")
    btn_cargar.pack(pady=20)

    # Panel de información
    info_frame = tk.Frame(root, relief=tk.GROOVE, borderwidth=1)
    info_frame.pack(pady=10, padx=20, fill=tk.X)
    
    tk.Label(info_frame, text="📊 Datos utilizados:", 
             font=("Arial", 9, "bold")).pack(anchor=tk.W, padx=10, pady=5)
    
    tk.Label(info_frame, text="• Tiempo de Retención: Fila 62, Columna B", 
             font=("Arial", 8)).pack(anchor=tk.W, padx=20)
    tk.Label(info_frame, text="• Altura Máxima: Fila 62, Columna J", 
             font=("Arial", 8)).pack(anchor=tk.W, padx=20)
    tk.Label(info_frame, text="• Tiempo Final: Fila 3, Columna AU", 
             font=("Arial", 8)).pack(anchor=tk.W, padx=20)
    
    # Estado
    estado_label = tk.Label(root, text="Listo", fg="green", font=("Arial", 8))
    estado_label.pack(pady=5)
    
    # Actualizar estado del botón
    def actualizar_estado():
        if btn_cargar['state'] == 'disabled':
            estado_label.config(text="Procesando...", fg="orange")
        else:
            estado_label.config(text="Listo", fg="green")
        root.after(100, actualizar_estado)
    
    actualizar_estado()
    
    root.mainloop()
