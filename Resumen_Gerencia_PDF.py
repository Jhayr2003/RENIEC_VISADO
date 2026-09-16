import os
import pandas as pd
import unicodedata
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from datetime import datetime


class ResumenGerenciaGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Resumen por Gerencia")
        self.root.geometry("700x450")
        self.root.resizable(False, False)
        self.root.lift()
        self.root.focus_force()

        # Variables para rutas
        self.gerencia_var = tk.StringVar()
        self.personas_var = tk.StringVar()
        self.salida_var = tk.StringVar()

        # Variables de periodo
        self.periodo_var = tk.StringVar(value="JUNIO")
        self.anio_var = tk.StringVar(value="2026")

        # Opción de timestamp (activado por defecto)
        self.timestamp_var = tk.BooleanVar(value=True)

        self.crear_interfaz()

    def crear_interfaz(self):
        titulo = tk.Label(self.root, text="GENERADOR DE RESUMEN POR GERENCIA",
                         font=("Arial", 14, "bold"), fg="darkblue")
        titulo.pack(pady=10)

        main_frame = tk.Frame(self.root, padx=20, pady=10)
        main_frame.pack(fill="both", expand=True)

        # Archivo GERENCIA.csv
        tk.Label(main_frame, text="1. Archivo GERENCIA.csv:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        tk.Entry(main_frame, textvariable=self.gerencia_var, width=50, state="readonly").grid(row=0, column=1, padx=5, pady=5)
        tk.Button(main_frame, text="Examinar...", command=lambda: self.seleccionar_archivo("gerencia"),
                 bg="#e3f2fd", width=12).grid(row=0, column=2, padx=5, pady=5)

        # Archivo PERSONAS.csv
        tk.Label(main_frame, text="2. Archivo PERSONAS.csv:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w", pady=5)
        tk.Entry(main_frame, textvariable=self.personas_var, width=50, state="readonly").grid(row=1, column=1, padx=5, pady=5)
        tk.Button(main_frame, text="Examinar...", command=lambda: self.seleccionar_archivo("personas"),
                 bg="#e3f2fd", width=12).grid(row=1, column=2, padx=5, pady=5)

        ttk.Separator(main_frame, orient='horizontal').grid(row=2, column=0, columnspan=3, sticky="ew", pady=10)

        # Periodo y Año
        frame_periodo = tk.Frame(main_frame)
        frame_periodo.grid(row=3, column=0, columnspan=3, pady=5)
        tk.Label(frame_periodo, text="Periodo:", font=("Arial", 10)).pack(side="left", padx=5)
        periodo_entry = tk.Entry(frame_periodo, textvariable=self.periodo_var, width=15, font=("Arial", 10))
        periodo_entry.pack(side="left", padx=5)
        tk.Label(frame_periodo, text="Año:", font=("Arial", 10)).pack(side="left", padx=5)
        anio_entry = tk.Entry(frame_periodo, textvariable=self.anio_var, width=10, font=("Arial", 10))
        anio_entry.pack(side="left", padx=5)

        # Checkbox timestamp (marcado por defecto)
        tk.Checkbutton(main_frame, text="Agregar timestamp al nombre del PDF",
                      variable=self.timestamp_var, font=("Arial", 9)).grid(row=4, column=0, columnspan=3, pady=5)

        ttk.Separator(main_frame, orient='horizontal').grid(row=5, column=0, columnspan=3, sticky="ew", pady=10)

        # Carpeta de salida
        tk.Label(main_frame, text="3. Carpeta de salida:", font=("Arial", 10, "bold")).grid(row=6, column=0, sticky="w", pady=5)
        tk.Entry(main_frame, textvariable=self.salida_var, width=50, state="readonly").grid(row=6, column=1, padx=5, pady=5)
        tk.Button(main_frame, text="Seleccionar...", command=self.seleccionar_carpeta_salida,
                 bg="#e3f2fd", width=12).grid(row=6, column=2, padx=5, pady=5)

        ttk.Separator(main_frame, orient='horizontal').grid(row=7, column=0, columnspan=3, sticky="ew", pady=10)

        # Botones
        frame_botones = tk.Frame(main_frame)
        frame_botones.grid(row=8, column=0, columnspan=3, pady=10)
        tk.Button(frame_botones, text="GENERAR PDF", command=self.generar_pdf,
                 bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), width=15, height=2).pack(side="left", padx=10)
        tk.Button(frame_botones, text="LIMPIAR", command=self.limpiar_campos,
                 bg="#FF9800", fg="white", font=("Arial", 10), width=10, height=2).pack(side="left", padx=10)
        tk.Button(frame_botones, text="SALIR", command=self.root.quit,
                 bg="#f44336", fg="white", font=("Arial", 10), width=10, height=2).pack(side="left", padx=10)

        # Barra de estado
        self.status_var = tk.StringVar(value="✅ Listo para procesar...")
        status_bar = tk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w",
                            font=("Arial", 9), bg="#f0f0f0")
        status_bar.pack(side="bottom", fill="x", padx=5, pady=5)

    def seleccionar_archivo(self, tipo):
        titulo = f"Selecciona el archivo {tipo.upper()}"
        archivo = filedialog.askopenfilename(
            title=titulo,
            filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")]
        )
        if archivo:
            if tipo == "gerencia":
                self.gerencia_var.set(archivo)
            elif tipo == "personas":
                self.personas_var.set(archivo)
            self.actualizar_salida_predeterminada()

    def seleccionar_carpeta_salida(self):
        carpeta = filedialog.askdirectory(
            title="Selecciona la carpeta donde se guardará el PDF",
            initialdir=os.path.dirname(self.salida_var.get()) if self.salida_var.get() else None
        )
        if carpeta:
            self.salida_var.set(carpeta)

    def actualizar_salida_predeterminada(self):
        if self.gerencia_var.get():
            base = os.path.dirname(self.gerencia_var.get())
        elif self.personas_var.get():
            base = os.path.dirname(self.personas_var.get())
        else:
            base = os.path.expanduser("~")
        self.salida_var.set(base)

    def limpiar_campos(self):
        self.gerencia_var.set("")
        self.personas_var.set("")
        self.salida_var.set("")
        self.status_var.set("✅ Campos limpiados")

    def leer_csv_utf8(self, ruta):
        """Intenta leer CSV con UTF-8 (con BOM), fallback a latin1 con conversión a UTF-8"""
        try:
            df = pd.read_csv(ruta, sep=";", encoding="utf-8-sig", dtype=str)
        except UnicodeDecodeError:
            df = pd.read_csv(ruta, sep=";", encoding="latin1", dtype=str)
            for col in df.select_dtypes(include=['object']).columns:
                df[col] = df[col].apply(
                    lambda x: x.encode('latin1').decode('utf-8', errors='ignore') if isinstance(x, str) else x
                )
        return df

    def normalizar_columnas_personas(self, df):
        """Renombra columnas alternativas (DNI, NOMBRES) a las que espera el script."""
        renombres = {}
        if "DNI" in df.columns and "CODIGO" not in df.columns:
            renombres["DNI"] = "CODIGO"
        if "NOMBRES" in df.columns and "NOMBRE" not in df.columns:
            renombres["NOMBRES"] = "NOMBRE"

        # Tolerancia a espacios / mayúsculas / minúsculas
        for col in list(df.columns):
            col_limpio = col.strip().upper()
            if col_limpio == "DNI" and col != "CODIGO":
                renombres[col] = "CODIGO"
            elif col_limpio == "NOMBRES":
                renombres[col] = "NOMBRE"
            elif col_limpio == "NOMBRE":
                renombres[col] = "NOMBRE"
            elif col_limpio == "DESCRIPCION":
                renombres[col] = "DESCRIPCION"
            elif col_limpio in ("CANTIDAD DE CONSULTA", "CANTIDAD_CONSULTA", "CANTIDAD"):
                renombres[col] = "CANTIDAD DE CONSULTA"

        if renombres:
            df = df.rename(columns=renombres)
        return df

    def generar_pdf(self):
        if not self.gerencia_var.get():
            messagebox.showerror("Error", "Selecciona el archivo GERENCIA.csv")
            return
        if not self.personas_var.get():
            messagebox.showerror("Error", "Selecciona el archivo PERSONAS.csv")
            return
        if not self.salida_var.get():
            messagebox.showerror("Error", "Selecciona la carpeta de salida")
            return

        try:
            self.status_var.set("⏳ Procesando archivos...")
            self.root.update()

            periodo = self.periodo_var.get().strip().upper()
            anio = self.anio_var.get().strip()

            df_gerencia = self.leer_csv_utf8(self.gerencia_var.get())
            df_personas = self.leer_csv_utf8(self.personas_var.get())

            # Limpieza de nombres de columnas (BOM y espacios)
            df_gerencia.columns = [col.replace('ï»¿', '').replace('\ufeff', '').strip() for col in df_gerencia.columns]
            df_personas.columns = [col.replace('ï»¿', '').replace('\ufeff', '').strip() for col in df_personas.columns]

            # Normalizar PERSONAS (DNI -> CODIGO, NOMBRES -> NOMBRE)
            df_personas = self.normalizar_columnas_personas(df_personas)

            # Diagnóstico por consola
            print("Columnas PERSONAS:", list(df_personas.columns))
            print("Columnas GERENCIA:", list(df_gerencia.columns))

            # =====================================================
            # Manejo de columnas CODIGO en GERENCIA
            # =====================================================
            columnas = list(df_gerencia.columns)
            indices_codigo = [i for i, col in enumerate(columnas) if 'CODIGO' in col.upper()]

            if len(indices_codigo) >= 2:
                columnas[indices_codigo[0]] = "CODIGO_DNI"
                columnas[indices_codigo[1]] = "CODIGO_UND"
            elif len(indices_codigo) == 1:
                columnas[indices_codigo[0]] = "CODIGO_DNI"
                df_gerencia["CODIGO_UND"] = "0000"
                columnas = list(df_gerencia.columns)
            else:
                df_gerencia["CODIGO_DNI"] = "00000000"
                df_gerencia["CODIGO_UND"] = "0000"
                columnas = list(df_gerencia.columns)

            df_gerencia.columns = columnas

            if "CODIGO_DNI" not in df_gerencia.columns:
                df_gerencia["CODIGO_DNI"] = "00000000"
            if "CODIGO_UND" not in df_gerencia.columns:
                df_gerencia["CODIGO_UND"] = "0000"

            df_gerencia["CODIGO_DNI"] = (
                df_gerencia["CODIGO_DNI"].astype(str).str.strip().str.zfill(8)
            )
            df_gerencia["CODIGO_UND"] = (
                df_gerencia["CODIGO_UND"].astype(str).str.strip().str.zfill(4)
            )

            # =====================================================
            # PERSONAS
            # =====================================================
            if "CODIGO" in df_personas.columns:
                df_personas["CODIGO"] = (
                    df_personas["CODIGO"].astype(str).str.strip().str.zfill(8)
                )
            else:
                df_personas["CODIGO"] = "00000000"

            if "NOMBRE" in df_personas.columns:
                df_personas["NOMBRE"] = df_personas["NOMBRE"].astype(str)
                df_personas["NOMBRE"] = df_personas["NOMBRE"].str.replace("■", "", regex=False).str.strip()
            else:
                df_personas["NOMBRE"] = "SIN NOMBRE"

            if "DESCRIPCION" not in df_personas.columns:
                df_personas["DESCRIPCION"] = "SIN DESCRIPCION"

            if "CANTIDAD DE CONSULTA" not in df_personas.columns:
                df_personas["CANTIDAD DE CONSULTA"] = 0
            df_personas["CANTIDAD DE CONSULTA"] = pd.to_numeric(
                df_personas["CANTIDAD DE CONSULTA"], errors="coerce"
            ).fillna(0).astype(int)

            if "GERENCIA" not in df_gerencia.columns:
                df_gerencia["GERENCIA"] = "OTROS"

            # =====================================================
            # MERGE
            # =====================================================
            df_final = pd.merge(
                df_personas,
                df_gerencia[["CODIGO_DNI", "GERENCIA", "CODIGO_UND"]],
                left_on="CODIGO",
                right_on="CODIGO_DNI",
                how="left"
            )
            df_final["GERENCIA"] = df_final["GERENCIA"].fillna("OTROS")
            df_final["CODIGO_UND"] = df_final["CODIGO_UND"].fillna("0000").astype(str).str.zfill(4)

            sin_match = (df_final["GERENCIA"] == "OTROS").sum()
            if sin_match > 0:
                print(f"⚠️  Filas sin coincidencia de gerencia: {sin_match}")

            df_resumen = df_final.groupby(
                ["GERENCIA", "DESCRIPCION"], as_index=False
            )["CANTIDAD DE CONSULTA"].sum()

            # Ordenar alfabéticamente sin acentos
            df_resumen["GERENCIA_normalizada"] = df_resumen["GERENCIA"].apply(
                lambda x: unicodedata.normalize("NFKD", str(x)).encode("ASCII", "ignore").decode("utf-8")
            )
            df_resumen = (
                df_resumen.sort_values("GERENCIA_normalizada")
                .reset_index(drop=True)
                .drop(columns=["GERENCIA_normalizada"])
            )

            # =====================================================
            # Reglas especiales de nombre de gerencia
            # Clave: (GERENCIA exacta, DESCRIPCION exacta)
            # Valor: nombre final a mostrar en el PDF
            # =====================================================
            EXCEPCIONES_NOMBRE = {
                ("RED DE AGENCIAS", "CONSULTAS EN LINEA VIA LINEA DEDICADA"): "RED DE AGENCIAS",
                # Agrega aquí más excepciones si lo necesitas:
                # ("OTRA GERENCIA", "OTRA DESCRIPCION"): "NOMBRE CORTO",
            }

            conteo = df_resumen.groupby("GERENCIA")["DESCRIPCION"].nunique()

            def construir_nombre(row):
                clave = (row["GERENCIA"], row["DESCRIPCION"])
                # 1) Regla especial
                if clave in EXCEPCIONES_NOMBRE:
                    return EXCEPCIONES_NOMBRE[clave]
                # 2) Varias descripciones -> agregar sufijo
                if conteo[row["GERENCIA"]] > 1:
                    return f"{row['GERENCIA']} - {row['DESCRIPCION']}"
                # 3) Caso normal
                return row["GERENCIA"]

            df_resumen["GERENCIA"] = df_resumen.apply(construir_nombre, axis=1)

            # Consolidar filas que hayan quedado con el mismo nombre tras el renombrado
            df_resumen = (
                df_resumen.groupby("GERENCIA", as_index=False)["CANTIDAD DE CONSULTA"]
                .sum()
                .sort_values("GERENCIA")
                .reset_index(drop=True)
            )

            total = df_resumen["CANTIDAD DE CONSULTA"].sum()

            # =====================================================
            # GENERAR PDF
            # =====================================================
            nombre_base = f"RESUMEN GERENCIA {periodo} {anio}.pdf"
            if self.timestamp_var.get():
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                nombre_base = f"RESUMEN GERENCIA {periodo} {anio}_{timestamp}.pdf"
            ruta_pdf = os.path.join(self.salida_var.get(), nombre_base)

            doc = SimpleDocTemplate(ruta_pdf, pagesize=A4)
            elementos = []
            styles = getSampleStyleSheet()

            style_title = ParagraphStyle(
                name="Title",
                parent=styles["Title"],
                fontName="Helvetica-Bold",
                fontSize=20,
                alignment=1,
                spaceAfter=20,
                wordWrap='CJK'
            )
            style_normal = ParagraphStyle(
                name="Normal",
                parent=styles["Normal"],
                fontName="Helvetica",
                fontSize=10
            )
            estilo_derecha = ParagraphStyle(
                name="Derecha",
                parent=styles["Normal"],
                fontName="Helvetica",
                fontSize=10,
                alignment=2
            )

            titulo = Paragraph(f"RESUMEN DE CONSULTAS POR GERENCIA - {periodo} {anio}", style_title)
            elementos.append(titulo)
            elementos.append(Spacer(1, 20))

            data = [["#", "GERENCIA", "CANTIDAD DE CONSULTA"]]
            for idx, row in df_resumen.iterrows():
                numero = idx + 1
                gerencia_paragraph = Paragraph(str(row["GERENCIA"]), style_normal)
                cantidad_paragraph = Paragraph(f'{row["CANTIDAD DE CONSULTA"]:,}', estilo_derecha)
                data.append([numero, gerencia_paragraph, cantidad_paragraph])

            ancho_hoja = A4[0] - 4 * cm
            col_widths = [2 * cm, ancho_hoja * 0.6, ancho_hoja * 0.4]

            tabla = Table(data, colWidths=col_widths, hAlign="CENTER")
            tabla.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.gray),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, 0), 10),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 12),
                ("GRID", (0, 0), (-1, -1), 1, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 1), (0, -1), "CENTER"),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 1), (-1, -1), 9),
            ]))

            elementos.append(tabla)
            elementos.append(Spacer(1, 12))

            total_paragraph = Paragraph(f"<b>TOTAL GENERAL: {total:,}</b>", estilo_derecha)
            elementos.append(total_paragraph)

            doc.build(elementos)

            self.status_var.set(f"✅ PDF generado: {nombre_base}")
            messagebox.showinfo("Éxito", f"PDF generado correctamente en:\n{ruta_pdf}")

        except Exception as e:
            self.status_var.set(f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error:\n{str(e)}")


if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = ResumenGerenciaGUI(root)
        root.mainloop()
    except Exception as e:
        print(f"Error al iniciar la interfaz: {e}")
        input("Presiona Enter para salir...")