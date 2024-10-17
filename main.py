import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

# Función para leer el archivo Excel
def cargar_excel():
     archivo_path = filedialog.askopenfilename(
        filetypes=[("Archivo Excel", "*.xlsx *.xls")]
    )
     if archivo_path:
        try:
            # Cargar datos del archivo Excel
            df = pd.read_excel(archivo_path)
            mostrar_datos(df)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo: {str(e)}")