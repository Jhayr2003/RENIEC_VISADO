import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading

# ==========================================
# INTERFAZ GRÁFICA
# ==========================================
class AppConsolidacion:
    def __init__(self, root):
        self.root = root
        self.root.title("Consolidación RRHH + Hermes/Prosegur")
        self.root.geometry("720x520")
        self.root.resizable(False, False)

        self.archivo_rrhh = tk.StringVar()
        self.archivo_hermes = tk.StringVar()
        self.carpeta_salida = tk.StringVar()

        # ---- Archivo RRHH ----
        tk.Label(root, text="Archivo RRHH (Planta de Personal y Funcionarios):",
                 anchor="w", font=("Segoe UI", 9, "bold")).pack(fill="x", padx=15, pady=(15, 2))
        frame1 = tk.Frame(root)
        frame1.pack(fill="x", padx=15)
        tk.Entry(frame1, textvariable=self.archivo_rrhh, width=70).pack(side="left", fill="x", expand=True)
        tk.Button(frame1, text="Examinar...", command=self.seleccionar_rrhh).pack(side="left", padx=(5, 0))

        # ---- Archivo Hermes ----
        tk.Label(root, text="Archivo Hermes y Prosegur:",
                 anchor="w", font=("Segoe UI", 9, "bold")).pack(fill="x", padx=15, pady=(12, 2))
        frame2 = tk.Frame(root)
        frame2.pack(fill="x", padx=15)
        tk.Entry(frame2, textvariable=self.archivo_hermes, width=70).pack(side="left", fill="x", expand=True)
        tk.Button(frame2, text="Examinar...", command=self.seleccionar_hermes).pack(side="left", padx=(5, 0))

        # ---- Carpeta de salida ----
        tk.Label(root, text="Carpeta donde se guardará el resultado:",
                 anchor="w", font=("Segoe UI", 9, "bold")).pack(fill="x", padx=15, pady=(12, 2))
        frame3 = tk.Frame(root)
        frame3.pack(fill="x", padx=15)
        tk.Entry(frame3, textvariable=self.carpeta_salida, width=70).pack(side="left", fill="x", expand=True)
        tk.Button(frame3, text="Examinar...", command=self.seleccionar_carpeta).pack(side="left", padx=(5, 0))

        # ---- Botón principal ----
        self.btn_ejecutar = tk.Button(root, text="▶  EJECUTAR CONSOLIDACIÓN",
                                      bg="#2E7D32", fg="white",
                                      font=("Segoe UI", 10, "bold"),
                                      command=self.ejecutar_en_hilo)
        self.btn_ejecutar.pack(pady=15, ipadx=20, ipady=6)

        # ---- Área de log ----
        tk.Label(root, text="Registro de proceso:",
                 anchor="w", font=("Segoe UI", 9, "bold")).pack(fill="x", padx=15)
        self.log = scrolledtext.ScrolledText(root, height=12, font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, padx=15, pady=(2, 15))

    # ---------- Selectores ----------
    def seleccionar_rrhh(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona el archivo de RRHH",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        if ruta:
            self.archivo_rrhh.set(ruta)

    def seleccionar_hermes(self):
        ruta = filedialog.askopenfilename(
            title="Selecciona el archivo de Hermes y Prosegur",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        if ruta:
            self.archivo_hermes.set(ruta)

    def seleccionar_carpeta(self):
        ruta = filedialog.askdirectory(title="Selecciona la carpeta de salida")
        if ruta:
            self.carpeta_salida.set(ruta)

    # ---------- Log ----------
    def escribir(self, texto):
        self.log.insert(tk.END, texto + "\n")
        self.log.see(tk.END)
        self.root.update_idletasks()

    # ---------- Ejecutar en hilo ----------
    def ejecutar_en_hilo(self):
        threading.Thread(target=self.ejecutar, daemon=True).start()

    # ==========================================
    # LÓGICA ORIGINAL (SIN ALTERACIONES)
    # ==========================================
    def ejecutar(self):
        archivo_rrhh = self.archivo_rrhh.get().strip()
        archivo_hermes = self.archivo_hermes.get().strip()
        carpeta = self.carpeta_salida.get().strip()

        if not archivo_rrhh or not os.path.isfile(archivo_rrhh):
            messagebox.showerror("Error", "Debes seleccionar el archivo de RRHH.")
            return
        if not archivo_hermes or not os.path.isfile(archivo_hermes):
            messagebox.showerror("Error", "Debes seleccionar el archivo de Hermes y Prosegur.")
            return
        if not carpeta or not os.path.isdir(carpeta):
            messagebox.showerror("Error", "Debes seleccionar una carpeta de salida válida.")
            return

        archivo_salida = os.path.join(carpeta, "UNION_HERMES_RRHH.xlsx")

        self.btn_ejecutar.config(state="disabled")
        self.log.delete("1.0", tk.END)

        def limpiar_dataframe(df, columnas_indices, nombres_columnas):
            df = df.iloc[:, columnas_indices].copy()
            df.columns = nombres_columnas
            df[nombres_columnas[0]] = df[nombres_columnas[0]].astype(str)

            # FILTRO 1: quitar filas con "TOTAL"
            df = df[~df[nombres_columnas[0]].str.upper().str.contains('TOTAL', na=False)]

            # FILTRO 2: quitar encabezados repetidos tipo "N° DNI"
            df = df[~df[nombres_columnas[0]].str.upper().str.contains('N° DNI', na=False)]

            # FILTRO 3: dejar solo DNIs numéricos
            df = df[df[nombres_columnas[0]].str.replace(r'\.0$', '', regex=True).str.isdigit()]

            return df

        try:
            self.escribir("Iniciando proceso de consolidación...")

            # ==========================================
            # 2. PROCESAR ARCHIVO RRHH
            #    Columnas: A, B, C, D, E
            # ==========================================
            self.escribir("Leyendo y limpiando archivo de RRHH...")
            df_rrhh_raw = pd.read_excel(archivo_rrhh, sheet_name=0)

            columnas_rrhh = [0, 1, 2, 3, 4]  # A, B, C, D, E
            nombres_rrhh = ['N° DNI', 'APELLIDOS Y NOMBRES', 'REPARTICIÓN', 'COD OFIC', 'OFICINA']

            df_rrhh = limpiar_dataframe(df_rrhh_raw, columnas_rrhh, nombres_rrhh)
            df_rrhh['EMPRESA'] = ''  # RRHH no tiene empresa

            # ==========================================
            # 3. PROCESAR ARCHIVO HERMES/PROSEGUR
            #    Columnas: A, B, F, E, H
            #    OFICINA queda vacío (solo viene de RRHH)
            # ==========================================
            self.escribir("Leyendo y limpiando archivo de Hermes y Prosegur...")
            df_hermes_raw = pd.read_excel(archivo_hermes, sheet_name=0)

            # A=0, B=1, F=5, E=4, H=7
            columnas_hermes = [0, 1, 5, 4, 7]
            nombres_hermes_temp = ['N° DNI', 'APELLIDOS Y NOMBRES', 'REPARTICIÓN', 'COD OFIC', 'EMPRESA']

            df_hermes = limpiar_dataframe(df_hermes_raw, columnas_hermes, nombres_hermes_temp)

            # Insertar OFICINA vacío en la posición 4 (entre COD OFIC y EMPRESA)
            df_hermes.insert(4, 'OFICINA', '')

            # Reordenar para que coincida con el orden final
            df_hermes = df_hermes[['N° DNI', 'APELLIDOS Y NOMBRES', 'REPARTICIÓN',
                                   'COD OFIC', 'OFICINA', 'EMPRESA']]

            # ==========================================
            # 4. UNIR AMBOS DATAFRAMES (RRHH primero, Hermes después)
            # ==========================================
            self.escribir("Uniendo datos (RRHH primero, luego Hermes)...")
            df_final = pd.concat([df_rrhh, df_hermes], ignore_index=True)

            # ==========================================
            # 5. LIMPIEZA FINAL DE FORMATO
            # ==========================================
            df_final.rename(columns={df_final.columns[0]: 'N° DNI'}, inplace=True)

            # DNI -> 8 dígitos, COD OFIC -> 4 dígitos
            df_final['N° DNI'] = df_final['N° DNI'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.zfill(8)
            df_final['COD OFIC'] = df_final['COD OFIC'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.zfill(4)

            # Rellenar celdas vacías
            df_final = df_final.fillna('')

            # ==========================================
            # 6. GUARDAR RESULTADO (hoja llamada "Planta")
            # ==========================================
            self.escribir(f"Guardando archivo en: {archivo_salida}")
            df_final.to_excel(archivo_salida, index=False, sheet_name='Planta')

            self.escribir("=========================================")
            self.escribir("¡PROCESO COMPLETADO CON ÉXITO!")
            self.escribir(f"Registros RRHH:           {len(df_rrhh)}")
            self.escribir(f"Registros Hermes/Prosegur:{len(df_hermes)}")
            self.escribir(f"Total registros:          {len(df_final)}")
            self.escribir(f"Archivo guardado como:    {os.path.basename(archivo_salida)}")
            self.escribir("Hoja:                     Planta")
            self.escribir("=========================================")

            # Vista previa
            self.escribir("\n--- Encabezados ---")
            self.escribir(str(list(df_final.columns)))
            self.escribir("\n--- Primeros 5 registros ---")
            self.escribir(df_final.head(5).to_string(index=False))
            self.escribir("\n--- Últimos 5 registros ---")
            self.escribir(df_final.tail(5).to_string(index=False))

            messagebox.showinfo("Éxito", f"Proceso completado.\n\nArchivo generado:\n{archivo_salida}")

        except FileNotFoundError as e:
            self.escribir(f"\n❌ ERROR: No se encontró uno de los archivos.")
            self.escribir(f"Verifica que los nombres y la ruta sean exactos: {e}")
            messagebox.showerror("Error", f"No se encontró uno de los archivos:\n{e}")
        except Exception as e:
            self.escribir(f"\n❌ OCURRIÓ UN ERROR INESPERADO: {e}")
            messagebox.showerror("Error inesperado", str(e))
        finally:
            self.btn_ejecutar.config(state="normal")


if __name__ == "__main__":
    root = tk.Tk()
    app = AppConsolidacion(root)
    root.mainloop()