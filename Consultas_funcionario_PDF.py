import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    PageBreak
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_RIGHT


# ============================================================
# FUNCIÓN PRINCIPAL - GENERAR PDF
# ============================================================

def generar_pdf(ruta_gerencia, ruta_funcionario, ruta_funcionario_1,
                ruta_salida, periodo, anio):

    try:

        # ====================================================
        # 📑 Leer CSVs  (ahora desde rutas individuales)
        # ====================================================

        df_gerencia = pd.read_csv(
            ruta_gerencia,
            sep=";",
            encoding="utf-8-sig"
        )

        df_funcionario = pd.read_csv(
            ruta_funcionario,
            sep=";",
            encoding="latin1"
        )

        df_funcionario_1 = pd.read_csv(
            ruta_funcionario_1,
            sep=";",
            encoding="latin1"
        )

        # ====================================================
        # 🔄 Normalizar códigos
        # ====================================================

        df_gerencia["CODIGO"] = (
            df_gerencia["CODIGO"]
            .astype(str)
            .str.zfill(8)
        )

        df_gerencia["CODIGO.1"] = (
            pd.to_numeric(
                df_gerencia["CODIGO.1"],
                errors="coerce"
            )
            .fillna(0)
            .astype(int)
            .astype(str)
            .str.zfill(4)
        )

        df_funcionario["F05_USER"] = (
            df_funcionario["F05_USER"]
            .astype(str)
            .str.strip()
        )

        df_funcionario["F05_DOC_IDENTIDAD"] = (
            df_funcionario["F05_DOC_IDENTIDAD"]
            .astype(str)
            .str.strip()
            .str.zfill(8)
        )

        df_funcionario_1["DNI"] = (
            df_funcionario_1["DNI"]
            .astype(str)
            .str.zfill(8)
        )

        # ====================================================
        # 🔗 Primer MERGE
        # ====================================================

        df_union = df_funcionario.merge(
            df_gerencia,
            left_on="F05_USER",
            right_on="CODIGO",
            how="inner"
        )

        df_union = df_union[
            [
                "F05_ALIAS_SISTEMA",
                "F05_DOC_IDENTIDAD",
                "F05_USER",
                "F05_MSG_RESUL",
                "F05_CONSULTA",
                "CODIGO",
                "NOMBRE",
                "GERENCIA",
                "CODIGO.1",
                "F05_FE_LOG_INI"
            ]
        ]

        df_union = df_union.rename(columns={
            "F05_ALIAS_SISTEMA": "F05_ALIAS_SISTEMA",
            "F05_DOC_IDENTIDAD": "F05_DOC_IDENTIDAD",
            "F05_USER": "F05_USER",
            "F05_MSG_RESUL": "F05_MSG_RESUL",
            "F05_CONSULTA": "F05_CONSULTA",
            "CODIGO": "F05_COD_USER",
            "NOMBRE": "F05_USER_NOMBRE",
            "GERENCIA": "F05_USER_GERENCIA",
            "CODIGO.1": "F05_USER_CODIGO",
            "F05_FE_LOG_INI": "F05_FE_LOG_INI"
        })

        # ====================================================
        # 🔗 Segundo MERGE
        # ====================================================

        df_union_1 = df_union.merge(
            df_funcionario_1,
            left_on="F05_DOC_IDENTIDAD",
            right_on="DNI",
            how="inner"
        )

        df_union_1 = df_union_1[
            [
                "F05_USER_CODIGO",
                "F05_COD_USER",
                "F05_USER_NOMBRE",
                "F05_CONSULTA",
                "F05_FE_LOG_INI",
                "APELLIDOS Y NOMBRES"
            ]
        ]

        df_union_1 = df_union_1.sort_values(
            by="APELLIDOS Y NOMBRES",
            ascending=True
        )

        df_union_1["F05_CONSULTA"] = df_union_1[
            "F05_CONSULTA"
        ].apply(
            lambda x:
            "" if pd.isna(x) or str(x).strip() == ""
            else str(x)
        )

        # ====================================================
        # 🎨 Estilos
        # ====================================================

        styles = getSampleStyleSheet()

        style_title = ParagraphStyle(
            name="Title",
            fontSize=13,
            alignment=TA_CENTER,
            spaceAfter=12,
            leading=16
        )

        # ====================================================
        # 📂 Salida del PDF
        # ====================================================

        pdf_path = ruta_salida

        doc = SimpleDocTemplate(
            pdf_path,
            pagesize=A4
        )

        elements = []

        # ====================================================
        # 👤 Agrupar por funcionario
        # ====================================================

        for nombre, grupo in df_union_1.groupby(
            "APELLIDOS Y NOMBRES"
        ):

            titulo = Paragraph(
                "CONSULTA RENIEC POR FUNCIONARIO<br/>"
                "SERVICIO DE LÍNEA DEDICADA - "
                + periodo
                + " "
                + anio,
                style_title
            )

            elements.append(titulo)
            elements.append(Spacer(1, 12))

            columnas = [
                "F05_USER_CODIGO",
                "F05_COD_USER",
                "F05_USER_NOMBRE",
                "F05_CONSULTA",
                "F05_FE_LOG_INI",
                "APELLIDOS Y NOMBRES"
            ]

            encabezados = {
                "F05_USER_CODIGO": "Unidad Organica",
                "F05_COD_USER": "Codigo Usuario",
                "F05_USER_NOMBRE": "Nombre Usuario",
                "F05_CONSULTA": "Datos de Consulta",
                "F05_FE_LOG_INI": "Fecha y Hora",
                "APELLIDOS Y NOMBRES": "Funcionario Relacionado"
            }

            data = [
                [encabezados[col] for col in columnas]
            ] + grupo[columnas].values.tolist()

            table = Table(
                data,
                repeatRows=1
            )

            table.setStyle(
                TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightblue),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 5),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (2, 0), (2, -1), "LEFT"),
                    ("ALIGN", (3, 0), (3, -1), "LEFT"),
                    ("ALIGN", (5, 0), (5, -1), "LEFT")
                ])
            )

            elements.append(table)
            elements.append(Spacer(1, 12))

            total_registros = len(grupo)

            resumen = Paragraph(
                f"<b>Total registros:</b> {total_registros}",
                styles["Normal"]
            )

            elements.append(resumen)
            elements.append(PageBreak())

        doc.build(elements)

        return pdf_path, len(df_union_1)

    except Exception as e:
        raise e


# ============================================================
# 📂 SELECTORES DE ARCHIVOS
# ============================================================

def seleccionar_csv(entrada, titulo):
    archivo = filedialog.askopenfilename(
        title=titulo,
        filetypes=[("Archivos CSV", "*.csv"), ("Todos", "*.*")]
    )
    if archivo:
        entrada.delete(0, tk.END)
        entrada.insert(0, archivo)


def seleccionar_gerencia():
    seleccionar_csv(entrada_gerencia, "Seleccionar archivo Gerencia.csv")


def seleccionar_funcionario():
    seleccionar_csv(entrada_funcionario, "Seleccionar archivo de Funcionario (Select ...)")


def seleccionar_funcionario_1():
    seleccionar_csv(entrada_funcionario_1, "Seleccionar archivo FUNCIONARIO.csv")


def seleccionar_salida():
    periodo = combo_periodo.get().strip() or "PERIODO"
    anio = entrada_anio.get().strip() or "AÑO"
    nombre_sugerido = f"CONSULTAS DE FUNCIONARIOS {periodo} {anio}.pdf"

    archivo = filedialog.asksaveasfilename(
        title="Guardar PDF como",
        defaultextension=".pdf",
        initialfile=nombre_sugerido,
        filetypes=[("Archivo PDF", "*.pdf")]
    )
    if archivo:
        entrada_salida.delete(0, tk.END)
        entrada_salida.insert(0, archivo)


# ============================================================
# ▶️ EJECUTAR PROCESO
# ============================================================

def ejecutar():

    ruta_gerencia = entrada_gerencia.get().strip()
    ruta_funcionario = entrada_funcionario.get().strip()
    ruta_funcionario_1 = entrada_funcionario_1.get().strip()
    ruta_salida = entrada_salida.get().strip()
    periodo = combo_periodo.get().strip()
    anio = entrada_anio.get().strip()

    # Validaciones
    if not ruta_gerencia or not os.path.exists(ruta_gerencia):
        messagebox.showwarning(
            "Falta información",
            "Seleccione el archivo Gerencia.csv"
        )
        return

    if not ruta_funcionario or not os.path.exists(ruta_funcionario):
        messagebox.showwarning(
            "Falta información",
            "Seleccione el archivo de Funcionario (Select ...)"
        )
        return

    if not ruta_funcionario_1 or not os.path.exists(ruta_funcionario_1):
        messagebox.showwarning(
            "Falta información",
            "Seleccione el archivo FUNCIONARIO.csv"
        )
        return

    if not ruta_salida:
        messagebox.showwarning(
            "Falta información",
            "Indique la ruta de salida del PDF."
        )
        return

    if not periodo:
        messagebox.showwarning(
            "Falta información",
            "Seleccione el periodo."
        )
        return

    if not anio:
        messagebox.showwarning(
            "Falta información",
            "Ingrese el año."
        )
        return

    try:

        boton_generar.config(state="disabled")
        estado.set("Procesando archivos y generando PDF...")
        ventana.update_idletasks()

        pdf_path, total = generar_pdf(
            ruta_gerencia,
            ruta_funcionario,
            ruta_funcionario_1,
            ruta_salida,
            periodo,
            anio
        )

        estado.set("Proceso finalizado correctamente.")

        messagebox.showinfo(
            "Proceso completado",
            "PDF generado correctamente.\n\n"
            f"Registros procesados: {total}\n\n"
            f"Archivo:\n{pdf_path}"
        )

    except Exception as e:

        estado.set("Se produjo un error.")

        messagebox.showerror(
            "Error durante el proceso",
            f"Ocurrió un error:\n\n{str(e)}"
        )

    finally:

        boton_generar.config(state="normal")


# ============================================================
# 🖥️ INTERFAZ GRÁFICA
# ============================================================

ventana = tk.Tk()
ventana.title("Generador de Reporte RENIEC")
ventana.geometry("820x560")
ventana.resizable(False, False)

style = ttk.Style()
try:
    style.theme_use("clam")
except:
    pass


# ============================================================
# 🏷️ TÍTULO
# ============================================================

titulo = ttk.Label(
    ventana,
    text="GENERADOR DE CONSULTAS RENIEC",
    font=("Arial", 16, "bold")
)
titulo.pack(pady=(20, 5))

subtitulo = ttk.Label(
    ventana,
    text="Reporte de consultas por funcionario"
)
subtitulo.pack(pady=(0, 15))


# ============================================================
# 📂 ARCHIVOS DE ENTRADA (3 CSV)
# ============================================================

frame_archivos = ttk.LabelFrame(ventana, text=" Archivos de entrada (CSV) ")
frame_archivos.pack(fill="x", padx=30, pady=8)

# --- CSV 1: Gerencia ---
ttk.Label(frame_archivos, text="1) Gerencia.csv:").grid(
    row=0, column=0, sticky="w", padx=8, pady=6
)
entrada_gerencia = ttk.Entry(frame_archivos, width=70)
entrada_gerencia.grid(row=0, column=1, padx=5, pady=6)
ttk.Button(frame_archivos, text="📂 Buscar", command=seleccionar_gerencia).grid(
    row=0, column=2, padx=8, pady=6
)

# --- CSV 2: Select (funcionario) ---
ttk.Label(frame_archivos, text="2) Select.csv:").grid(
    row=1, column=0, sticky="w", padx=8, pady=6
)
entrada_funcionario = ttk.Entry(frame_archivos, width=70)
entrada_funcionario.grid(row=1, column=1, padx=5, pady=6)
ttk.Button(frame_archivos, text="📂 Buscar", command=seleccionar_funcionario).grid(
    row=1, column=2, padx=8, pady=6
)

# --- CSV 3: FUNCIONARIO ---
ttk.Label(frame_archivos, text="3) Funcionarios.csv:").grid(
    row=2, column=0, sticky="w", padx=8, pady=6
)
entrada_funcionario_1 = ttk.Entry(frame_archivos, width=70)
entrada_funcionario_1.grid(row=2, column=1, padx=5, pady=6)
ttk.Button(frame_archivos, text="📂 Buscar", command=seleccionar_funcionario_1).grid(
    row=2, column=2, padx=8, pady=6
)


# ============================================================
# 💾 ARCHIVO DE SALIDA (PDF)
# ============================================================

frame_salida = ttk.LabelFrame(ventana, text=" Archivo de salida (PDF) ")
frame_salida.pack(fill="x", padx=30, pady=8)

ttk.Label(frame_salida, text="Guardar PDF como:").grid(
    row=0, column=0, sticky="w", padx=8, pady=6
)
entrada_salida = ttk.Entry(frame_salida, width=70)
entrada_salida.grid(row=0, column=1, padx=5, pady=6)
ttk.Button(frame_salida, text="💾 Guardar", command=seleccionar_salida).grid(
    row=0, column=2, padx=8, pady=6
)


# ============================================================
# 📅 PERIODO Y AÑO
# ============================================================

frame_datos = ttk.Frame(ventana)
frame_datos.pack(padx=40, pady=12)

ttk.Label(frame_datos, text="Periodo:").grid(row=0, column=0, padx=10)
combo_periodo = ttk.Combobox(
    frame_datos,
    values=[
        "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
        "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
    ],
    state="readonly",
    width=15
)
combo_periodo.grid(row=1, column=0, padx=10, pady=5)
combo_periodo.set("AGOSTO")

ttk.Label(frame_datos, text="Año:").grid(row=0, column=1, padx=10)
entrada_anio = ttk.Entry(frame_datos, width=15)
entrada_anio.grid(row=1, column=1, padx=10, pady=5)
entrada_anio.insert(0, "2026")


# ============================================================
# ▶️ BOTÓN GENERAR
# ============================================================

boton_generar = ttk.Button(
    ventana,
    text="▶  GENERAR PDF",
    command=ejecutar
)
boton_generar.pack(pady=15, ipadx=25, ipady=8)


# ============================================================
# 📊 ESTADO
# ============================================================

estado = tk.StringVar()
estado.set("Listo para generar el reporte.")

label_estado = ttk.Label(ventana, textvariable=estado)
label_estado.pack(pady=5)


# ============================================================
# ℹ️ PIE
# ============================================================

label_pie = ttk.Label(
    ventana,
    text="Servicio de Línea Dedicada - RENIEC",
    font=("Arial", 8)
)
label_pie.pack(side="bottom", pady=10)


# ============================================================
# 🚀 INICIAR
# ============================================================

ventana.mainloop()