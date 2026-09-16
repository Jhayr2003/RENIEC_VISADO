import os
import pandas as pd
import unicodedata
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (BaseDocTemplate, Frame, PageTemplate, Table,
                                TableStyle, Paragraph, Spacer)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
from reportlab.lib import colors
from datetime import datetime


class PDFGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de PDFs por Gerencia")
        self.root.geometry("700x450")
        self.root.resizable(False, False)

        self.gerencia_var = tk.StringVar()
        self.personas_var = tk.StringVar()
        self.salida_var = tk.StringVar()

        self.periodo_var = tk.StringVar(value="MAYO")
        self.anio_var = tk.StringVar(value="2026")

        # >>> Timestamp siempre activo (no configurable por el usuario)
        self.timestamp_var = tk.BooleanVar(value=True)

        self.crear_interfaz()

    def crear_interfaz(self):
        titulo = tk.Label(self.root, text="GENERADOR DE REPORTES PDF POR GERENCIA",
                          font=("Arial", 14, "bold"), fg="darkblue")
        titulo.pack(pady=10)

        main_frame = tk.Frame(self.root, padx=20, pady=10)
        main_frame.pack(fill="both", expand=True)

        tk.Label(main_frame, text="1. Archivo GERENCIA.csv:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        tk.Entry(main_frame, textvariable=self.gerencia_var, width=50, state="readonly").grid(row=0, column=1, padx=5, pady=5)
        tk.Button(main_frame, text="Examinar...", command=lambda: self.seleccionar_archivo("gerencia"),
                  bg="#e3f2fd", width=12).grid(row=0, column=2, padx=5, pady=5)

        tk.Label(main_frame, text="2. Archivo PERSONAS.csv:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w", pady=5)
        tk.Entry(main_frame, textvariable=self.personas_var, width=50, state="readonly").grid(row=1, column=1, padx=5, pady=5)
        tk.Button(main_frame, text="Examinar...", command=lambda: self.seleccionar_archivo("personas"),
                  bg="#e3f2fd", width=12).grid(row=1, column=2, padx=5, pady=5)

        ttk.Separator(main_frame, orient='horizontal').grid(row=2, column=0, columnspan=3, sticky="ew", pady=10)

        frame_periodo = tk.Frame(main_frame)
        frame_periodo.grid(row=3, column=0, columnspan=3, pady=5)
        tk.Label(frame_periodo, text="Periodo:", font=("Arial", 10)).pack(side="left", padx=5)
        tk.Entry(frame_periodo, textvariable=self.periodo_var, width=15, font=("Arial", 10)).pack(side="left", padx=5)
        tk.Label(frame_periodo, text="Año:", font=("Arial", 10)).pack(side="left", padx=5)
        tk.Entry(frame_periodo, textvariable=self.anio_var, width=10, font=("Arial", 10)).pack(side="left", padx=5)

        # Checkbutton bloqueado y siempre marcado
        tk.Checkbutton(
            main_frame,
            text="Agregar timestamp al nombre de los PDFs (siempre activo)",
            variable=self.timestamp_var,
            font=("Arial", 9),
            state="disabled",              # el usuario no puede interactuar
            disabledforeground="#555555",
            selectcolor="#4CAF50"
        ).grid(row=4, column=0, columnspan=3, pady=5)

        ttk.Separator(main_frame, orient='horizontal').grid(row=5, column=0, columnspan=3, sticky="ew", pady=10)

        tk.Label(main_frame, text="3. Carpeta de salida (raíz):", font=("Arial", 10, "bold")).grid(row=6, column=0, sticky="w", pady=5)
        tk.Entry(main_frame, textvariable=self.salida_var, width=50, state="readonly").grid(row=6, column=1, padx=5, pady=5)
        tk.Button(main_frame, text="Seleccionar...", command=self.seleccionar_carpeta_salida,
                  bg="#e3f2fd", width=12).grid(row=6, column=2, padx=5, pady=5)

        ttk.Separator(main_frame, orient='horizontal').grid(row=7, column=0, columnspan=3, sticky="ew", pady=10)

        frame_botones = tk.Frame(main_frame)
        frame_botones.grid(row=8, column=0, columnspan=3, pady=10)
        tk.Button(frame_botones, text="GENERAR PDFs", command=self.generar_pdfs,
                  bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), width=15, height=2).pack(side="left", padx=10)
        tk.Button(frame_botones, text="LIMPIAR", command=self.limpiar_campos,
                  bg="#FF9800", fg="white", font=("Arial", 10), width=10, height=2).pack(side="left", padx=10)
        tk.Button(frame_botones, text="SALIR", command=self.root.quit,
                  bg="#f44336", fg="white", font=("Arial", 10), width=10, height=2).pack(side="left", padx=10)

        self.status_var = tk.StringVar(value="✅ Listo para procesar...")
        status_bar = tk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w",
                              font=("Arial", 9), bg="#f0f0f0")
        status_bar.pack(side="bottom", fill="x", padx=5, pady=5)

    # ------------------------------------------------------------------
    # Selección de archivos y carpeta
    # ------------------------------------------------------------------
    def seleccionar_archivo(self, tipo):
        archivo = filedialog.askopenfilename(
            title=f"Selecciona el archivo {tipo.upper()}",
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
            title="Selecciona la carpeta raíz donde se guardarán los PDFs",
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

    # ------------------------------------------------------------------
    # Lectura segura de CSV
    # ------------------------------------------------------------------
    def _leer_csv(self, ruta):
        try:
            return pd.read_csv(ruta, sep=";", encoding="utf-8-sig", dtype=str)
        except UnicodeDecodeError:
            df = pd.read_csv(ruta, sep=";", encoding="latin1", dtype=str)
            for col in df.select_dtypes(include=['object']).columns:
                df[col] = df[col].apply(
                    lambda x: x.encode('latin1').decode('utf-8', errors='ignore') if isinstance(x, str) else x
                )
            return df

    # ------------------------------------------------------------------
    # Normalización de columnas de PERSONAS (DNI → CODIGO, NOMBRES → NOMBRE)
    # ------------------------------------------------------------------
    def _normalizar_personas(self, df):
        renombres = {}
        for col in list(df.columns):
            cu = col.strip().upper()
            if cu == "DNI" and "CODIGO" not in df.columns:
                renombres[col] = "CODIGO"
            elif cu == "NOMBRES":
                renombres[col] = "NOMBRE"
            elif cu == "NOMBRE":
                renombres[col] = "NOMBRE"
            elif cu == "DESCRIPCION":
                renombres[col] = "DESCRIPCION"
            elif cu in ("CANTIDAD DE CONSULTA", "CANTIDAD_CONSULTA", "CANTIDAD"):
                renombres[col] = "CANTIDAD DE CONSULTA"
        if renombres:
            df = df.rename(columns=renombres)
        return df

    # ------------------------------------------------------------------
    # Normalización de columnas de GERENCIA
    # ------------------------------------------------------------------
    def _normalizar_gerencia(self, df):
        columnas = list(df.columns)
        idx_codigo = [i for i, c in enumerate(columnas) if 'CODIGO' in c.upper()]

        if len(idx_codigo) >= 2:
            columnas[idx_codigo[0]] = "CODIGO_DNI"
            columnas[idx_codigo[1]] = "CODIGO_UND"
        elif len(idx_codigo) == 1:
            columnas[idx_codigo[0]] = "CODIGO_DNI"
            df["CODIGO_UND"] = "0000"
            columnas = list(df.columns)
        else:
            df["CODIGO_DNI"] = "00000000"
            df["CODIGO_UND"] = "0000"
            columnas = list(df.columns)

        df.columns = columnas

        if "GERENCIA" not in df.columns:
            df["GERENCIA"] = "OTROS"

        return df

    # ------------------------------------------------------------------
    # Proceso principal
    # ------------------------------------------------------------------
    def generar_pdfs(self):
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

            # Crear subcarpeta PDFS
            carpeta_pdfs = os.path.join(self.salida_var.get(), "PDFS")
            os.makedirs(carpeta_pdfs, exist_ok=True)

            # ----------------------------
            # Lectura de archivos
            # ----------------------------
            df_gerencia = self._leer_csv(self.gerencia_var.get())
            df_personas = self._leer_csv(self.personas_var.get())

            df_gerencia.columns = [col.replace('ï»¿', '').replace('\ufeff', '').strip() for col in df_gerencia.columns]
            df_personas.columns = [col.replace('ï»¿', '').replace('\ufeff', '').strip() for col in df_personas.columns]

            # ----------------------------
            # Normalizar columnas
            # ----------------------------
            df_personas = self._normalizar_personas(df_personas)
            df_gerencia = self._normalizar_gerencia(df_gerencia)

            print("Columnas GERENCIA:", df_gerencia.columns.tolist())
            print("Columnas PERSONAS:", df_personas.columns.tolist())

            # ----------------------------
            # Formatear códigos
            # ----------------------------
            df_gerencia["CODIGO_DNI"] = df_gerencia["CODIGO_DNI"].astype(str).str.strip().str.zfill(8)
            df_gerencia["CODIGO_UND"] = df_gerencia["CODIGO_UND"].astype(str).str.strip().str.zfill(4)

            if "CODIGO" not in df_personas.columns:
                df_personas["CODIGO"] = "00000000"
            df_personas["CODIGO"] = df_personas["CODIGO"].astype(str).str.strip().str.zfill(8)

            if "NOMBRE" not in df_personas.columns:
                df_personas["NOMBRE"] = "SIN NOMBRE"
            df_personas["NOMBRE"] = (
                df_personas["NOMBRE"].astype(str)
                .str.replace("■", "", regex=False)
                .str.replace(r"[^\x00-\x7F]+", "", regex=True)
                .str.strip()
            )

            if "DESCRIPCION" not in df_personas.columns:
                df_personas["DESCRIPCION"] = "SIN DESCRIPCION"

            if "CANTIDAD DE CONSULTA" not in df_personas.columns:
                df_personas["CANTIDAD DE CONSULTA"] = 0
            df_personas["CANTIDAD DE CONSULTA"] = pd.to_numeric(
                df_personas["CANTIDAD DE CONSULTA"], errors="coerce"
            ).fillna(0).astype(int)

            # ----------------------------
            # Merge
            # ----------------------------
            df_final = pd.merge(
                df_personas,
                df_gerencia[["CODIGO_UND", "CODIGO_DNI", "GERENCIA"]],
                left_on="CODIGO",
                right_on="CODIGO_DNI",
                how="left"
            )
            df_final["GERENCIA"] = df_final["GERENCIA"].fillna("OTROS")
            df_final["CODIGO_UND"] = df_final["CODIGO_UND"].fillna("0000").astype(str).str.zfill(4)
            df_final["NOMBRE"] = df_final["NOMBRE"].fillna("")
            df_final = df_final.drop(columns=["CODIGO_DNI"], errors="ignore")

            print("📊 Cantidad de registros por GERENCIA:")
            for g, c in df_final["GERENCIA"].value_counts().items():
                print(f"- {g}: {c}")

            # ----------------------------
            # Utilidades
            # ----------------------------
            def limpiar_nombre_archivo(nombre):
                nombre = str(nombre)
                nombre = unicodedata.normalize('NFC', nombre)
                nombre = re.sub(r'[\\/*?:"<>|]', '_', nombre)
                nombre = re.sub(r'\s+', ' ', nombre).strip()
                return nombre

            styles = getSampleStyleSheet()
            style_subtitle = ParagraphStyle("subtitle", parent=styles["Heading2"],
                                            alignment=TA_LEFT, fontSize=12, spaceAfter=10)
            style_total = ParagraphStyle("total", parent=styles["Normal"],
                                         alignment=TA_RIGHT, fontSize=10, textColor=colors.black)
            style_table_header = ParagraphStyle(
                name="table_header", fontSize=7, alignment=TA_CENTER,
                leading=8, textColor=colors.black
            )

            def encabezado(canvas, doc, gerencia):
                canvas.saveState()
                canvas.setFont("Helvetica-Bold", 10)
                canvas.drawCentredString(A4[0] / 2, A4[1] - 40, f" DE {gerencia}")
                canvas.setFont("Helvetica", 9)
                canvas.drawCentredString(A4[0] / 2, A4[1] - 55,
                                         f"Servicio Líneas Dedicadas - {periodo} {anio}")
                canvas.restoreState()

            # ----------------------------
            # Generar un PDF por (GERENCIA, DESCRIPCION)
            # ----------------------------
            contador = 0
            for (gerencia, descripcion), df_ger in df_final.groupby(["GERENCIA", "DESCRIPCION"]):
                gerencia_limpia = limpiar_nombre_archivo(gerencia)
                desc_limpia = limpiar_nombre_archivo(descripcion)

                # Si hay más de una descripción por gerencia, incluirla en el nombre
                conteo_desc = df_final[df_final["GERENCIA"] == gerencia]["DESCRIPCION"].nunique()
                if conteo_desc > 1:
                    nombre_base = f"{anio}-{periodo}-{gerencia_limpia} - {desc_limpia}.pdf"
                else:
                    nombre_base = f"{anio}-{periodo}-{gerencia_limpia}.pdf"

                # >>> Timestamp SIEMPRE aplicado (checkbox bloqueado en True)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                nombre_base = nombre_base.replace(".pdf", f"_{timestamp}.pdf")

                nombre_pdf = os.path.join(carpeta_pdfs, nombre_base)

                doc = BaseDocTemplate(nombre_pdf, pagesize=A4)
                frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height - 60, id="normal")
                template = PageTemplate(
                    id="reporte",
                    frames=frame,
                    onPage=lambda canvas, doc, g=gerencia: encabezado(canvas, doc, g)
                )
                doc.addPageTemplates([template])

                story = []
                df_ger = df_ger.sort_values(["CODIGO_UND", "NOMBRE"], ascending=[True, True])

                data = [[
                    Paragraph("Unidad Orgánica", style_table_header),
                    Paragraph("Código Usuario", style_table_header),
                    Paragraph("Nombre del Usuario", style_table_header),
                    Paragraph("Total<br/>Consultas", style_table_header)
                ]]

                for _, row in df_ger.iterrows():
                    data.append([
                        row["CODIGO_UND"],
                        row["CODIGO"],
                        row["NOMBRE"],
                        f"{int(row['CANTIDAD DE CONSULTA'])}"
                    ])

                table = Table(data, colWidths=[60, 70, 200, 90], repeatRows=1)
                table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                    ("ALIGN", (3, 1), (3, -1), "RIGHT"),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
                ]))

                story.append(Paragraph(f"Descripción: {descripcion}", style_subtitle))
                story.append(table)
                story.append(Spacer(1, 12))

                total_desc = df_ger["CANTIDAD DE CONSULTA"].sum()
                story.append(Paragraph(f"TOTAL {descripcion}: {int(total_desc):,}", style_total))

                doc.build(story)
                print(f"✅ PDF generado: {nombre_pdf}")
                contador += 1

            self.status_var.set(f"✅ {contador} PDF(s) generados en: {carpeta_pdfs}")
            messagebox.showinfo(
                "Éxito",
                f"✅ Se generaron {contador} PDF(s) en:\n{carpeta_pdfs}"
            )

        except Exception as e:
            self.status_var.set(f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFGeneratorGUI(root)
    root.mainloop()