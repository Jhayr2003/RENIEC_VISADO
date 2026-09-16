from pathlib import Path
from datetime import datetime
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd


# ============================================================
# RUTAS POR DEFECTO (solo para abrir el explorador)
# ============================================================
CARPETA_DEFECTO = Path(r"C:\Users\pra_jcerna\Downloads\pythoncodevisado\archivos para visar")


# ============================================================
# AUXILIARES
# ============================================================
def normalizar(v):
    return str(v).strip().upper()


def obtener_hoja_gerentes(xls):
    for hoja in xls.sheet_names:
        if normalizar(hoja) == "GERENTES":
            return hoja
    raise ValueError("No se encontró la hoja 'Gerentes'.")


# ============================================================
# LÓGICA PRINCIPAL (sin cambios)
# ============================================================
def procesar(archivo: Path, salida: Path, log=print):
    if not archivo.exists():
        raise FileNotFoundError(f"No existe el archivo: {archivo}")

    xls = pd.ExcelFile(archivo)
    hoja = obtener_hoja_gerentes(xls)
    log(f"Hoja encontrada: {hoja}")

    df_raw = pd.read_excel(archivo, sheet_name=hoja, header=None, dtype=str)
    log(f"Dimensiones leídas: {df_raw.shape} (filas x columnas)")

    fila_enc = None
    for i in range(min(15, len(df_raw))):
        valores = [normalizar(v) for v in df_raw.iloc[i].tolist()]
        if any(v == "DNI" for v in valores):
            fila_enc = i
            break

    if fila_enc is None:
        raise ValueError("No se encontró la fila de encabezados (DNI).")

    log(f"Fila de encabezado detectada (índice): {fila_enc}")

    df = df_raw.iloc[fila_enc + 1:].reset_index(drop=True)

    n_cols = min(5, df.shape[1])
    log(f"Columnas disponibles: {df.shape[1]} -> usando: {n_cols}")

    if n_cols < 5:
        log("ADVERTENCIA: la hoja tiene menos de 5 columnas. "
            "No se podrá extraer 'COD OFIC' desde la columna E.")

    df = df.iloc[:, :n_cols].copy()

    nombres = ["N°", "DNI", "APELLIDOS Y NOMBRES", "GERENCIA", "COD OFIC"]
    df.columns = nombres[:n_cols]

    if "COD OFIC" not in df.columns:
        df["COD OFIC"] = ""

    for c in df.columns:
        df[c] = (
            df[c].astype(str).str.strip()
            .replace({"nan": "", "None": "", "NaT": ""})
        )

    df = df[df["N°"].str.match(r"^\d+$", na=False)]
    df = df[["N°", "DNI", "APELLIDOS Y NOMBRES", "GERENCIA", "COD OFIC"]]

    df.to_csv(salida, index=False, sep=";", encoding="utf-8-sig")

    log(f"✓ CSV generado: {salida}")
    log(f"  Filas exportadas: {len(df)}")
    return salida, len(df)


def generar_nombre_salida(carpeta: Path) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return carpeta / f"Funcionarios_{ts}.csv"


# ============================================================
# INTERFAZ GRÁFICA
# ============================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Extractor de Funcionarios -> CSV")
        self.geometry("780x520")
        self.resizable(False, False)

        # ---- Archivo de entrada ----
        frame_top = tk.Frame(self, padx=10, pady=10)
        frame_top.pack(fill="x")

        tk.Label(frame_top, text="Archivo Excel (RRHH Planta de Personal y Funcionarios.xlsx) :").grid(row=0, column=0, sticky="w")
        self.var_archivo = tk.StringVar(value="")   # ← VACÍO
        tk.Entry(frame_top, textvariable=self.var_archivo, width=70).grid(
            row=0, column=1, padx=5, sticky="we"
        )
        tk.Button(frame_top, text="Examinar...", command=self.seleccionar_archivo).grid(
            row=0, column=2, padx=5
        )

        # ---- Carpeta de salida ----
        tk.Label(frame_top, text="Carpeta salida:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.var_carpeta = tk.StringVar(value="")   # ← VACÍO
        tk.Entry(frame_top, textvariable=self.var_carpeta, width=70).grid(
            row=1, column=1, padx=5, pady=(8, 0), sticky="we"
        )
        tk.Button(frame_top, text="Examinar...", command=self.seleccionar_carpeta).grid(
            row=1, column=2, padx=5, pady=(8, 0)
        )

        frame_top.columnconfigure(1, weight=1)

        # ---- Botón ejecutar ----
        frame_btn = tk.Frame(self, padx=10, pady=5)
        frame_btn.pack(fill="x")
        self.btn_ejecutar = tk.Button(
            frame_btn, text="Ejecutar", width=20,
            bg="#2E7D32", fg="white",
            command=self.ejecutar
        )
        self.btn_ejecutar.pack(side="left")

        self.lbl_estado = tk.Label(frame_btn, text="Selecciona un archivo y una carpeta.", fg="gray")
        self.lbl_estado.pack(side="left", padx=10)

        # ---- Área de log ----
        frame_log = tk.Frame(self, padx=10, pady=5)
        frame_log.pack(fill="both", expand=True)

        tk.Label(frame_log, text="Salida:").pack(anchor="w")
        self.txt_log = scrolledtext.ScrolledText(frame_log, height=20, wrap="word")
        self.txt_log.pack(fill="both", expand=True)

    # ---------- Acciones ----------
    def seleccionar_archivo(self):
        initial = str(CARPETA_DEFECTO) if CARPETA_DEFECTO.exists() else "."
        ruta = filedialog.askopenfilename(
            title="Selecciona el archivo Excel",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
            initialdir=initial
        )
        if ruta:
            self.var_archivo.set(ruta)
            # Autocompletar carpeta de salida con la carpeta del archivo
            if not self.var_carpeta.get().strip():
                self.var_carpeta.set(str(Path(ruta).parent))

    def seleccionar_carpeta(self):
        initial = str(CARPETA_DEFECTO) if CARPETA_DEFECTO.exists() else "."
        carpeta = filedialog.askdirectory(
            title="Selecciona la carpeta de salida",
            initialdir=initial
        )
        if carpeta:
            self.var_carpeta.set(carpeta)

    def log(self, mensaje):
        self.txt_log.insert("end", mensaje + "\n")
        self.txt_log.see("end")
        self.update_idletasks()

    def ejecutar(self):
        archivo_txt = self.var_archivo.get().strip()
        carpeta_txt = self.var_carpeta.get().strip()

        # ---- Validaciones ----
        if not archivo_txt:
            messagebox.showwarning("Atención", "Debes seleccionar un archivo Excel.")
            return
        if not carpeta_txt:
            messagebox.showwarning("Atención", "Debes seleccionar una carpeta de salida.")
            return

        archivo = Path(archivo_txt)
        carpeta = Path(carpeta_txt)

        if not archivo.exists():
            messagebox.showerror("Error", f"No existe el archivo:\n{archivo}")
            return
        if not carpeta.exists():
            try:
                carpeta.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo crear la carpeta:\n{e}")
                return

        salida = generar_nombre_salida(carpeta)

        self.txt_log.delete("1.0", "end")
        self.lbl_estado.config(text="Procesando...", fg="blue")
        self.btn_ejecutar.config(state="disabled")

        threading.Thread(
            target=self._run,
            args=(archivo, salida),
            daemon=True
        ).start()

    def _run(self, archivo, salida):
        try:
            ruta, filas = procesar(archivo, salida, log=self.log)
            self.lbl_estado.config(text=f"✓ Listo. {filas} filas.", fg="green")
            messagebox.showinfo("Éxito", f"CSV generado:\n{ruta}")
        except Exception as e:
            self.log(f"✗ ERROR: {e}")
            self.lbl_estado.config(text="Error.", fg="red")
            messagebox.showerror("Error", str(e))
        finally:
            self.btn_ejecutar.config(state="normal")


# ============================================================
# EJECUCIÓN
# ============================================================
if __name__ == "__main__":
    App().mainloop()