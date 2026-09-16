"""
EXTRACTOR DE CONSULTAS POR SERVICIO 
-----------------------------------------------------------
Lee un archivo TXT y genera CSV con:
DNI | NOMBRES | CANTIDAD DE CONSULTA | DESCRIPCION
"""
# Librerias necesarias para el funcionamiento adecuado del programa 
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import time
from datetime import datetime

# ============================================================================
# FUNCIONES
# ============================================================================

def limpiar_numero(texto):
    """Convierte texto con puntos de miles a entero"""
    if not texto:
        return 0
    num = ''.join(c for c in texto if c.isdigit())
    return int(num) if num else 0


def procesar_archivo(ruta):
    """
    Procesa archivo extrayendo registros:
    - DNI: primer token alfanumérico (ej: ADMIN, 123456)
    - Cantidad: último número (con puntos de miles) al final de la línea
    - Nombres: todo lo que está entre código y cantidad
    - Descripción: detectado por línea "SERVICIO: ..."
    - Ignora líneas de encabezado y basura (DEL, RUC:, NO ESPECIFICADO)
    """
    
    regex_servicio = re.compile(r'^SERVICIO:\s*(.+)$', re.IGNORECASE)
    
    registros = []
    descripcion_actual = "NO ESPECIFICADO"
    
    # Intentar con varias codificaciones
    codificaciones = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
    for enc in codificaciones:
        try:
            with open(ruta, "r", encoding=enc) as f:
                lineas = f.readlines()
            break
        except UnicodeDecodeError:
            continue
    else:
        with open(ruta, "r", encoding='latin-1') as f:
            lineas = f.readlines()
    
    for linea in lineas:
        linea = linea.strip()
        if not linea:
            continue
        
        # ============================================================
        # 1. DETECTAR DESCRIPCIÓN (SERVICIO)
        # ============================================================
        if linea.startswith('SERVICIO:'):
            match = regex_servicio.match(linea)
            if match:
                descripcion_actual = match.group(1).strip().upper()
            continue  # No procesar esta línea como dato
        
        # ============================================================
        # 2. IGNORAR LÍNEAS DE ENCABEZADO O BASURA
        # ============================================================
        if linea.startswith(('DEL ', 'RUC:', 'PAGINA')):
            continue
        
        if "NO ESPECIFICADO" in linea:
            continue
        
        # ============================================================
        # 3. DETECTAR SEPARADOR Y DIVIDIR
        # ============================================================
        if ';' in linea:
            partes = linea.split(';')
        elif ',' in linea:
            partes = linea.split(',')
        else:
            partes = linea.split()
        
        partes = [p.strip() for p in partes if p.strip() != '']
        if len(partes) < 2:
            continue
        
        # ============================================================
        # 4. EXTRAER CANTIDAD (ÚLTIMO NÚMERO CON PUNTOS DE MILES)
        # ============================================================
        cantidad = None
        idx_cantidad = -1
        for i in range(len(partes)-1, -1, -1):
            token = partes[i]
            if token.replace('.', '').isdigit():
                cantidad = limpiar_numero(token)
                idx_cantidad = i
                break
        
        if cantidad is None or cantidad == 0:
            continue
        
        # ============================================================
        # 5. EXTRAER DNI (PRIMER TOKEN)
        # ============================================================
        dni = partes[0]
        if not dni:
            continue
        
        if dni.upper() in {"CODIGO", "USUARIO", "NOMBRE", "CANTIDAD", "TOTAL"}:
            continue
        
        # ============================================================
        # 6. EXTRAER NOMBRES (ENTRE DNI Y CANTIDAD)
        # ============================================================
        if idx_cantidad <= 1:
            nombres = ""
        else:
            nombres = ' '.join(partes[1:idx_cantidad]).strip()
        
        nombres = nombres.strip().strip('"').strip("'")
        
        # ============================================================
        # 7. AGREGAR REGISTRO (DNI, NOMBRES, CANTIDAD, DESCRIPCION)
        # ============================================================
        registros.append([dni, nombres, cantidad, descripcion_actual])
    
    return registros


def guardar_csv(registros, ruta_salida):
    """Guarda CSV delimitado por punto y coma (;) y sin comillas,
       con encabezados: DNI, NOMBRES, CANTIDAD DE CONSULTA, DESCRIPCION"""
    
    if not registros:
        return False, "No hay datos para guardar"
    
    df = pd.DataFrame(registros, columns=[
        "DNI", "NOMBRES", "CANTIDAD DE CONSULTA", "DESCRIPCION"
    ])
    
    df["DNI"] = df["DNI"].astype(str)
    
    try:
        # utf-8-sig para que Excel muestre correctamente tildes/ñ
        # sep=';' -> Delimitador punto y coma
        # quoting=0 -> Sin comillas (solo las usaría si el texto contuviera el delimitador)
        df.to_csv(
            ruta_salida,
            index=False,
            sep=';',
            encoding='utf-8-sig',
            quoting=0
        )
        
        total_consultas = df['CANTIDAD DE CONSULTA'].sum()
        return True, f"✅ {len(df)} registros | {total_consultas:,} consultas"
    
    except PermissionError:
        return False, "❌ Permiso denegado. Cierra el archivo CSV si está abierto."
    except Exception as e:
        return False, f"❌ Error: {str(e)}"


# ============================================================================
# INTERFAZ SIMPLIFICADA
# ============================================================================

class AppExtractor:
    def __init__(self):
        self.ventana = tk.Tk()
        self.ventana.title("Reporte de consultas por servicio")
        self.ventana.geometry("600x500")
        self.ventana.resizable(False,False)
        
        self.ventana.configure(bg='#2c3e50')
        self.ruta_entrada = tk.StringVar()
        self.carpeta_destino = tk.StringVar()
        self.crear_interfaz()
        
    def crear_interfaz(self):
        titulo = tk.Label(
            self.ventana,
            text="Reporte de Consultas por Servicio",
            font=('Arial', 16, 'bold'),
            bg='#2c3e50',
            fg='white'
        )
        titulo.pack(pady=20)
        
        frame = tk.Frame(self.ventana, bg='#ecf0f1', padx=20, pady=20)
        frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        tk.Label(frame, text="📄 Archivo TXT:", font=('Arial', 10, 'bold'), 
                bg='#ecf0f1').grid(row=0, column=0, sticky='w', pady=5)
        tk.Entry(frame, textvariable=self.ruta_entrada, width=45, 
                font=('Arial', 9)).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(frame, text="📂", command=self.seleccionar_entrada, width=3, 
                 bg='#3498db', fg='white', cursor='hand2').grid(row=0, column=2, pady=5)
        
        tk.Label(frame, text="📁 Carpeta destino:", font=('Arial', 10, 'bold'), 
                bg='#ecf0f1').grid(row=1, column=0, sticky='w', pady=5)
        self.lbl_carpeta = tk.Label(frame, text="Misma carpeta del archivo", font=('Arial', 9), 
                                    bg='#ecf0f1', fg='#27ae60')
        self.lbl_carpeta.grid(row=1, column=1, sticky='w', pady=5)
        tk.Button(frame, text="📁 Cambiar", command=self.seleccionar_carpeta, width=8, 
                 bg='#27ae60', fg='white', cursor='hand2').grid(row=1, column=2, pady=5)
        
        self.btn_procesar = tk.Button(
            frame,
            text="🚀 PROCESAR",
            command=self.procesar,
            bg='#e74c3c',
            fg='white',
            font=('Arial', 12, 'bold'),
            cursor='hand2',
            height=2
        )
        self.btn_procesar.grid(row=2, column=0, columnspan=3, pady=20, sticky='ew')
        
        self.lbl_estado = tk.Label(frame, text="✅ Listo para procesar", font=('Arial', 9), 
                                   bg='#ecf0f1', fg='#27ae60')
        self.lbl_estado.grid(row=3, column=0, columnspan=3, pady=5)
        
        tk.Label(frame, text="📊 Resultados:", font=('Arial', 10, 'bold'), 
                bg='#ecf0f1').grid(row=4, column=0, sticky='w', pady=(10,0))
        
        self.txt_resultados = tk.Text(frame, height=10, width=60, font=('Consolas', 8))
        self.txt_resultados.grid(row=5, column=0, columnspan=3, pady=5, sticky='nsew')
        
        scrollbar = tk.Scrollbar(self.txt_resultados)
        scrollbar.pack(side='right', fill='y')
        self.txt_resultados.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.txt_resultados.yview)
        
        frame.grid_rowconfigure(5, weight=1)
        frame.grid_columnconfigure(1, weight=1)
    
    def seleccionar_entrada(self):
        archivo = filedialog.askopenfilename(filetypes=[("Archivos TXT", "*.txt")])
        if archivo:
            self.ruta_entrada.set(archivo)
            if not self.carpeta_destino.get():
                self.carpeta_origen = Path(archivo).parent
                self.lbl_carpeta.config(text=str(self.carpeta_origen), fg='#27ae60')
    
    def seleccionar_carpeta(self):
        if not self.ruta_entrada.get():
            messagebox.showwarning("Advertencia", "Primero seleccione el archivo TXT")
            return
        carpeta = filedialog.askdirectory(title="Seleccionar carpeta para guardar el CSV")
        if carpeta:
            self.carpeta_destino.set(carpeta)
            self.lbl_carpeta.config(text=carpeta, fg='#e74c3c')
    
    def mostrar_resultado(self, texto):
        self.txt_resultados.delete(1.0, tk.END)
        self.txt_resultados.insert(tk.END, texto)
    
    def procesar(self):
        if not self.ruta_entrada.get():
            messagebox.showerror("Error", "Seleccione un archivo TXT")
            return
        
        if self.carpeta_destino.get():
            carpeta = self.carpeta_destino.get()
        else:
            carpeta = str(Path(self.ruta_entrada.get()).parent)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        nombre_archivo = f"PERSONA_{timestamp}.csv"
        ruta_salida = Path(carpeta) / nombre_archivo
        
        self.btn_procesar.config(state='disabled', text='⏳ PROCESANDO...')
        self.lbl_estado.config(text="⏳ Procesando archivo...", fg='#f39c12')
        self.ventana.update()
        
        inicio = time.time()
        
        try:
            datos = procesar_archivo(self.ruta_entrada.get())
            
            if not datos:
                self.mostrar_resultado("❌ No se encontraron datos en el archivo")
                self.lbl_estado.config(text="❌ No se encontraron datos", fg='#e74c3c')
                return
            
            df = pd.DataFrame(datos, columns=["DNI", "NOMBRES", "CANTIDAD DE CONSULTA", "DESCRIPCION"])
            
            resumen = f"✅ PROCESADO EN {time.time() - inicio:.1f} SEGUNDOS\n\n"
            resumen += f"📊 TOTAL REGISTROS: {len(df):,}\n"
            resumen += f"📊 TOTAL CONSULTAS: {df['CANTIDAD DE CONSULTA'].sum():,}\n"
            resumen += f"📊 PROMEDIO: {df['CANTIDAD DE CONSULTA'].mean():.1f}\n\n"
            resumen += "📋 POR DESCRIPCION (SERVICIO):\n"
            
            for desc, grupo in df.groupby('DESCRIPCION'):
                resumen += f"   • {desc[:40]}: {len(grupo)} registros | {grupo['CANTIDAD DE CONSULTA'].sum():,} consultas\n"
            
            resumen += f"\n📋 PRIMEROS 5 REGISTROS:\n"
            resumen += df.head(5).to_string(index=False)
            
            resumen += f"\n\n💾 Archivo: {nombre_archivo}"
            resumen += f"\n📁 Carpeta: {carpeta}"
            
            self.mostrar_resultado(resumen)
            
            self.lbl_estado.config(text="💾 Guardando CSV...", fg='#3498db')
            self.ventana.update()
            
            exito, mensaje = guardar_csv(datos, str(ruta_salida))
            
            if exito:
                self.lbl_estado.config(text=f"✅ {mensaje}", fg='#27ae60')
                messagebox.showinfo("Éxito", f"Archivo guardado:\n{ruta_salida}")
            else:
                self.lbl_estado.config(text=f"❌ {mensaje}", fg='#e74c3c')
                messagebox.showerror("Error", mensaje)
                
        except Exception as e:
            self.mostrar_resultado(f"❌ ERROR: {str(e)}")
            self.lbl_estado.config(text="❌ Error en el procesamiento", fg='#e74c3c')
            messagebox.showerror("Error", str(e))
        
        finally:
            self.btn_procesar.config(state='normal', text='🚀 PROCESAR')
    
    def ejecutar(self):
        self.ventana.mainloop()


if __name__ == "__main__":
    app = AppExtractor()
    app.ejecutar()