from datetime import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


# ─── FUNCIONES DE INTERFAZ ──────────────────────────────────────────────────
def generar_nombre_destino(ruta_origen):
    """Genera ruta de destino con nombre fijo 'Gerencia.csv' en la misma carpeta."""
    directorio = os.path.dirname(ruta_origen)
    return os.path.join(directorio, "Gerencia.csv")


def seleccionar_archivo_entrada():
    ruta = filedialog.askopenfilename(
        title="Selecciona el archivo Excel de origen",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")],
    )
    if ruta:
        entrada_var.set(ruta)
        destino_sugerido = generar_nombre_destino(ruta)
        destino_var.set(destino_sugerido)


def seleccionar_destino():
    sugerido = destino_var.get()
    if not sugerido:
        sugerido = "Gerencia.csv"
    ruta = filedialog.asksaveasfilename(
        title="Guardar archivo modificado como CSV",
        defaultextension=".csv",
        filetypes=[("Archivos CSV", "*.csv")],
        initialfile=os.path.basename(sugerido),
        initialdir=(
            os.path.dirname(sugerido) if os.path.dirname(sugerido) else None
        ),
    )
    if ruta:
        destino_var.set(ruta)


def procesar():
    archivo_origen = entrada_var.get()
    archivo_destino = destino_var.get()
    if not archivo_origen or not archivo_destino:
        messagebox.showwarning(
            "Faltan datos", "Selecciona el archivo de origen y el destino."
        )
        return

    try:
        # ─── 1. Leer saltando 3 filas; DNI (col 0) y COD OFIC (col 3) como string ──
        df = pd.read_excel(
            archivo_origen,
            sheet_name="Planta",
            skiprows=3,
            converters={0: str, 3: str},
        )

        # ─── 2. Limpiar espacios en blanco ──
        df = df.map(lambda x: x.strip() if isinstance(x, str) else x)

        # ─── 3. Identificar columnas por posición ──
        col_dni = df.columns[0]           # A - N° DNI
        col_nombre = df.columns[1]        # B - APELLIDOS Y NOMBRES
        col_reparticion = df.columns[2]   # C - REPARTICIÓN
        col_cod = df.columns[3]           # D - COD OFIC
        col_oficina = df.columns[4]       # E - OFICINA
        col_empresa = df.columns[5] if len(df.columns) > 5 else None  # F - EMPRESA

        # ─── 4. Aplicar regla: código 0-999 OR empresa != BN (y no vacía) ──
        def actualizar_reparto(row):
            cod = row[col_cod]

            # Condición 1: código numérico entre 0 y 999
            cumple_codigo = False
            if pd.notna(cod):
                try:
                    cod_num = int(cod)
                    cumple_codigo = (0 <= cod_num <= 999)
                except (ValueError, TypeError):
                    cumple_codigo = False

            # Condición 2: empresa distinta a BN y NO vacía
            cumple_empresa = False
            if col_empresa is not None:
                val = row[col_empresa]
                if pd.notna(val):
                    val_str = str(val).strip().upper()
                    if val_str != "":
                        cumple_empresa = (val_str != "BN")

            if cumple_codigo or cumple_empresa:
                return "RED DE AGENCIAS"
            return row[col_reparticion]

        df[col_reparticion] = df.apply(actualizar_reparto, axis=1)

        # ─── 5. Formatear COD OFIC a 4 dígitos con ceros ──
        df[col_cod] = df[col_cod].astype(str).str.zfill(4)

        # ─── 6. Seleccionar SOLO las 4 columnas del reporte final ──
        df_reporte = df[[col_dni, col_nombre, col_reparticion, col_cod]].copy()
        df_reporte.columns = ["CODIGO", "NOMBRE", "GERENCIA", "CODIGO"]

        # ─── 7. Guardar CSV con separador ';' y UTF-8 con BOM ──
        df_reporte.to_csv(
            archivo_destino, sep=";", index=False, encoding="utf-8-sig"
        )
        messagebox.showinfo(
            "Éxito",
            "✅ Proceso completado.\n"
            "Regla aplicada: 'RED DE AGENCIAS' si el código está entre 0-999 "
            "o si la empresa es distinta a BN (excluyendo vacíos).\n"
            f"Archivo guardado como 'Gerencia.csv' en:\n{archivo_destino}",
        )
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error:\n{e}")


# ─── VENTANA PRINCIPAL ──────────────────────────────────────────────────────
root = tk.Tk()
root.title("Procesador de Plantilla UNION_HERMES_RRHH.xlsx → Gerencia.CSV")
root.geometry("550x220")

entrada_var = tk.StringVar()
destino_var = tk.StringVar()

tk.Label(root, text="Archivo de origen (Excel):").grid(
    row=0, column=0, padx=5, pady=10, sticky="e"
)
tk.Entry(root, textvariable=entrada_var, width=45).grid(
    row=0, column=1, padx=5, pady=10
)
tk.Button(
    root, text="Examinar...", command=seleccionar_archivo_entrada
).grid(row=0, column=2, padx=5, pady=10)

tk.Label(root, text="Destino (CSV):").grid(
    row=1, column=0, padx=5, pady=10, sticky="e"
)
tk.Entry(root, textvariable=destino_var, width=45).grid(
    row=1, column=1, padx=5, pady=10
)
tk.Button(root, text="Guardar como...", command=seleccionar_destino).grid(
    row=1, column=2, padx=5, pady=10
)

tk.Button(
    root, text="PROCESAR", command=procesar, bg="lightblue", width=20
).grid(row=2, column=1, pady=20)

root.mainloop()