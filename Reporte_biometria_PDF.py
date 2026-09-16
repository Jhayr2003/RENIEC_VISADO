import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (BaseDocTemplate, Frame, PageTemplate, Table,
                                TableStyle, Paragraph, Spacer)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_RIGHT
from datetime import datetime


def formatear_codigo(valor, longitud):
    """Convierte un valor a string, elimina decimales y rellena con ceros a la izquierda."""
    if pd.isna(valor):
        return '0' * longitud
    s = str(valor).strip()
    if s.endswith('.0'):
        s = s[:-2]
    elif '.' in s:
        try:
            s = str(int(float(s)))
        except (ValueError, TypeError):
            pass
    return s.zfill(longitud)


class GeneradorPDFBiometrico:
    def __init__(self, root):
        self.root = root
        self.root.title("Generador de Reporte PDF - Servicio Biométrico")
        self.root.geometry("750x480")
        self.root.resizable(False, False)

        self.biometrico_var = tk.StringVar()
        self.gerencia_var = tk.StringVar()
        self.pdf_var = tk.StringVar()

        self.periodo_var = tk.StringVar(value="MAYO")
        self.anio_var = tk.StringVar(value="2026")

        self.timestamp_var = tk.BooleanVar(value=True)

        self.crear_interfaz()

    def crear_interfaz(self):
        titulo = tk.Label(self.root, text="GENERADOR DE REPORTE PDF - SERVICIO BIOMÉTRICO",
                          font=("Arial", 14, "bold"), fg="darkblue")
        titulo.pack(pady=10)

        main_frame = tk.Frame(self.root, padx=20, pady=10)
        main_frame.pack(fill="both", expand=True)

        tk.Label(main_frame, text="1. Archivo Biometria.csv:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=5)
        tk.Entry(main_frame, textvariable=self.biometrico_var, width=50, state="readonly").grid(row=0, column=1, padx=5, pady=5)
        tk.Button(main_frame, text="Examinar...", command=lambda: self.seleccionar_archivo("biometrico"),
                  bg="#e3f2fd", width=12).grid(row=0, column=2, padx=5, pady=5)

        tk.Label(main_frame, text="2. Archivo GERENCIA.csv:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky="w", pady=5)
        tk.Entry(main_frame, textvariable=self.gerencia_var, width=50, state="readonly").grid(row=1, column=1, padx=5, pady=5)
        tk.Button(main_frame, text="Examinar...", command=lambda: self.seleccionar_archivo("gerencia"),
                  bg="#e3f2fd", width=12).grid(row=1, column=2, padx=5, pady=5)

        ttk.Separator(main_frame, orient='horizontal').grid(row=2, column=0, columnspan=3, sticky="ew", pady=10)

        frame_periodo = tk.Frame(main_frame)
        frame_periodo.grid(row=3, column=0, columnspan=3, pady=5)

        tk.Label(frame_periodo, text="Periodo:", font=("Arial", 10)).pack(side="left", padx=5)
        tk.Entry(frame_periodo, textvariable=self.periodo_var, width=15, font=("Arial", 10)).pack(side="left", padx=5)

        tk.Label(frame_periodo, text="Año:", font=("Arial", 10)).pack(side="left", padx=5)
        tk.Entry(frame_periodo, textvariable=self.anio_var, width=10, font=("Arial", 10)).pack(side="left", padx=5)

        tk.Checkbutton(main_frame, text="Agregar timestamp al nombre del PDF",
                       variable=self.timestamp_var, font=("Arial", 9)).grid(row=4, column=0, columnspan=3, pady=5)

        ttk.Separator(main_frame, orient='horizontal').grid(row=5, column=0, columnspan=3, sticky="ew", pady=10)

        tk.Label(main_frame, text="3. PDF de Salida:", font=("Arial", 10, "bold")).grid(row=6, column=0, sticky="w", pady=5)
        tk.Entry(main_frame, textvariable=self.pdf_var, width=50, state="readonly").grid(row=6, column=1, padx=5, pady=5)
        tk.Button(main_frame, text="Guardar como...", command=self.seleccionar_pdf_salida,
                  bg="#e3f2fd", width=12).grid(row=6, column=2, padx=5, pady=5)

        ttk.Separator(main_frame, orient='horizontal').grid(row=7, column=0, columnspan=3, sticky="ew", pady=10)

        frame_botones = tk.Frame(main_frame)
        frame_botones.grid(row=8, column=0, columnspan=3, pady=10)

        tk.Button(frame_botones, text="GENERAR PDF", command=self.generar_pdf,
                  bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), width=15, height=2).pack(side="left", padx=10)
        tk.Button(frame_botones, text="LIMPIAR", command=self.limpiar_campos,
                  bg="#FF9800", fg="white", font=("Arial", 10), width=10, height=2).pack(side="left", padx=10)
        tk.Button(frame_botones, text="SALIR", command=self.root.quit,
                  bg="#f44336", fg="white", font=("Arial", 10), width=10, height=2).pack(side="left", padx=10)

        self.status_var = tk.StringVar(value="✅ Listo para procesar...")
        status_bar = tk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w",
                              font=("Arial", 9), bg="#f0f0f0")
        status_bar.pack(side="bottom", fill="x", padx=5, pady=5)

    def generar_nombre_con_timestamp(self, nombre_base):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre, extension = os.path.splitext(nombre_base)
        return f"{nombre}_{timestamp}{extension}"

    def seleccionar_archivo(self, tipo):
        titulo = f"Selecciona el archivo {tipo.upper()}"
        if tipo == "biometrico":
            archivo = filedialog.askopenfilename(
                title=titulo,
                filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")]
            )
            if archivo:
                self.biometrico_var.set(archivo)
                self.actualizar_pdf_predeterminado()
        elif tipo == "gerencia":
            archivo = filedialog.askopenfilename(
                title=titulo,
                filetypes=[("Archivos CSV", "*.csv"), ("Todos los archivos", "*.*")]
            )
            if archivo:
                self.gerencia_var.set(archivo)
                self.actualizar_pdf_predeterminado()

    def actualizar_pdf_predeterminado(self):
        periodo = self.periodo_var.get()
        anio = self.anio_var.get()

        if all([self.biometrico_var.get(), self.gerencia_var.get()]):
            directorio = os.path.dirname(self.biometrico_var.get())
            if not directorio:
                directorio = os.path.dirname(self.gerencia_var.get())
            if not directorio:
                directorio = os.path.expanduser("~")

            nombre_pdf = f"APLICATIVO-BIOMETRICO {periodo} {anio}.pdf"

            if self.timestamp_var.get():
                nombre_pdf = self.generar_nombre_con_timestamp(nombre_pdf)

            ruta_pdf = os.path.join(directorio, nombre_pdf)
            self.pdf_var.set(ruta_pdf)

    def seleccionar_pdf_salida(self):
        periodo = self.periodo_var.get()
        anio = self.anio_var.get()
        nombre_defecto = f"APLICATIVO-BIOMETRICO {periodo} {anio}.pdf"

        if self.timestamp_var.get():
            nombre_defecto = self.generar_nombre_con_timestamp(nombre_defecto)

        archivo = filedialog.asksaveasfilename(
            title="Guardar PDF como...",
            defaultextension=".pdf",
            filetypes=[("Archivos PDF", "*.pdf"), ("Todos los archivos", "*.*")],
            initialfile=nombre_defecto,
            initialdir=os.path.dirname(self.pdf_var.get()) if self.pdf_var.get() else None
        )
        if archivo:
            self.pdf_var.set(archivo)

    def limpiar_campos(self):
        self.biometrico_var.set("")
        self.gerencia_var.set("")
        self.pdf_var.set("")
        self.status_var.set("✅ Campos limpiados")

    def _leer_csv_seguro(self, ruta):
        try:
            return pd.read_csv(ruta, sep=";", encoding="utf-8-sig")
        except UnicodeDecodeError:
            return pd.read_csv(ruta, sep=";", encoding="latin1")

    def generar_pdf(self):
        if not self.biometrico_var.get():
            messagebox.showerror("Error", "Selecciona el archivo Biometria.csv")
            return
        if not self.gerencia_var.get():
            messagebox.showerror("Error", "Selecciona el archivo GERENCIA.csv")
            return
        if not self.pdf_var.get():
            messagebox.showerror("Error", "Especifica la ubicación del PDF de salida")
            return

        try:
            self.status_var.set("⏳ Procesando archivos...")
            self.root.update()

            periodo = self.periodo_var.get()
            anio = self.anio_var.get()
            pdf_path = self.pdf_var.get()

            if self.timestamp_var.get():
                directorio = os.path.dirname(pdf_path)
                nombre_base = os.path.basename(pdf_path)
                nombre_sin_timestamp = f"APLICATIVO-BIOMETRICO {periodo} {anio}.pdf"
                if nombre_base == nombre_sin_timestamp:
                    pdf_path = os.path.join(directorio, self.generar_nombre_con_timestamp(nombre_base))
                    self.pdf_var.set(pdf_path)

            # --- Leer CSVs ---
            self.status_var.set("⏳ Leyendo archivos CSV...")
            self.root.update()
            df_biometrico = self._leer_csv_seguro(self.biometrico_var.get())
            df_gerencia = self._leer_csv_seguro(self.gerencia_var.get())

            df_gerencia.columns = df_gerencia.columns.str.replace('ï»¿', '').str.strip()
            df_biometrico.columns = df_biometrico.columns.str.replace('ï»¿', '').str.strip()

            # --- Procesar biometría ---
            self.status_var.set("⏳ Procesando datos biométricos...")
            self.root.update()

            df_biometrico["TOTAL"] = df_biometrico["TOTAL"].fillna(0)
            df_biometrico["DNI_CAJERO"] = df_biometrico["DNI_CAJERO"].fillna("")
            df_biometrico["DNI_CAJERO"] = df_biometrico["DNI_CAJERO"].apply(lambda x: formatear_codigo(x, 8))

            # --- Identificar columnas de código en GERENCIA ---
            columnas_codigo = [col for col in df_gerencia.columns if 'CODIGO' in col.upper()]

            if len(columnas_codigo) >= 2:
                df_gerencia[columnas_codigo[0]] = df_gerencia[columnas_codigo[0]].apply(lambda x: formatear_codigo(x, 8))
                df_gerencia[columnas_codigo[1]] = df_gerencia[columnas_codigo[1]].apply(lambda x: formatear_codigo(x, 4))
                df_gerencia = df_gerencia.rename(columns={
                    columnas_codigo[0]: "DNI_CAJERO",
                    columnas_codigo[1]: "COD_UND"
                })
            elif len(columnas_codigo) == 1:
                df_gerencia[columnas_codigo[0]] = df_gerencia[columnas_codigo[0]].apply(lambda x: formatear_codigo(x, 8))
                df_gerencia = df_gerencia.rename(columns={columnas_codigo[0]: "DNI_CAJERO"})
                df_gerencia["COD_UND"] = "0000"
            else:
                df_gerencia["DNI_CAJERO"] = "00000000"
                df_gerencia["COD_UND"] = "0000"

            if "NOMBRE" not in df_gerencia.columns:
                df_gerencia["NOMBRE"] = "SIN NOMBRE"
            else:
                df_gerencia["NOMBRE"] = df_gerencia["NOMBRE"].fillna("SIN NOMBRE").astype(str)

            # --- Merge ---
            self.status_var.set("⏳ Realizando merge de datos...")
            self.root.update()

            df_union = pd.merge(df_biometrico, df_gerencia, on="DNI_CAJERO", how="inner")

            if df_union.empty:
                messagebox.showwarning(
                    "Advertencia",
                    "No se encontraron coincidencias entre los DNI de ambos archivos.\n"
                    "Verifica que los datos sean correctos."
                )
                self.status_var.set("⚠️ No se encontraron coincidencias")
                return

            # --- Construir df_pdf (sin Empresa, sin separación) ---
            df_pdf = pd.DataFrame({
                'Unidad Organica': df_union['COD_UND'],
                'Código de Usuario': df_union['DNI_CAJERO'],
                'Nombre de Usuario': df_union['NOMBRE'],
                'Total': df_union['TOTAL']
            }).sort_values(by=['Unidad Organica', 'Nombre de Usuario']).reset_index(drop=True)

            total_general = df_pdf['Total'].sum()

            # --- Estilos ---
            styles = getSampleStyleSheet()
            right_style = styles['Normal']
            right_style.alignment = TA_RIGHT

            def make_header_footer(titulo_texto):
                def hf(canvas, doc):
                    canvas.saveState()
                    P = Paragraph(titulo_texto, styles['Title'])
                    w, h = P.wrap(doc.width, doc.topMargin)
                    P.drawOn(canvas, doc.leftMargin, A4[1] - h - 0.5 * cm)
                    page_num_text = f"Página {canvas.getPageNumber()}"
                    canvas.setFont("Helvetica", 9)
                    canvas.drawRightString(A4[0] - doc.rightMargin, 1 * cm, page_num_text)
                    canvas.restoreState()
                return hf

            titulo_general = (f"CONSULTAS RENIEC POR GERENCIA DE RED DE AGENCIAS // <br/>"
                              f"Servicio Biométrico - {periodo} {anio}")

            self.status_var.set("⏳ Generando PDF...")
            self.root.update()

            doc = BaseDocTemplate(pdf_path, pagesize=A4)

            frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width,
                          doc.height - 1 * cm, id='principal')

            template = PageTemplate(id='principal', frames=[frame],
                                    onPage=make_header_footer(titulo_general))

            doc.addPageTemplates([template])

            # --- Tabla única con todos los registros ---
            data = [['Unidad Organica', 'Código de Usuario', 'Nombre de Usuario', 'Total']]
            for _, row in df_pdf.iterrows():
                data.append([
                    row['Unidad Organica'],
                    row['Código de Usuario'],
                    row['Nombre de Usuario'],
                    row['Total']
                ])

            tabla = Table(data, colWidths=[3 * cm, 3.5 * cm, 9 * cm, 3 * cm], repeatRows=1)
            tabla.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('ALIGN', (3, 1), (3, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
            ]))

            elements = []
            elements.append(Spacer(1, 1 * cm))
            elements.append(tabla)
            elements.append(Spacer(1, 0.3 * cm))
            elements.append(Paragraph(f"<b>TOTAL  : {total_general:,}</b>", right_style))

            doc.build(elements)

            self.status_var.set(f"✅ PDF generado exitosamente: {os.path.basename(pdf_path)}")
            messagebox.showinfo(
                "Éxito",
                f"✅ PDF generado correctamente en:\n{pdf_path}\n\n"
                f"📊 Total de registros: {len(df_pdf)}\n"
                f"📈 Suma total: {total_general:,}"
            )

        except Exception as e:
            self.status_var.set(f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = GeneradorPDFBiometrico(root)
    root.mainloop()