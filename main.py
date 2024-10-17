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

# Función para mostrar los datos del archivo Excel en el Treeview
def mostrar_datos(df):
    # Limpiar el Treeview antes de insertar nuevos datos
    for i in tree.get_children():
        tree.delete(i)
     # Insertar los encabezados
    tree["columns"] = list(df.columns)
    tree["show"] = "headings"
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, anchor="center", width=100)
     # Insertar filas en el Treeview
    for index, row in df.iterrows():
        tree.insert("", "end", values=list(row))

# Configurar la ventana principal de Tkinter
root = tk.Tk()
root.title("Lector de Excel")
root.geometry("800x400")
root.configure(bg="#e0f7fa")
# Estilos personalizados
style = ttk.Style()
style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"), background="#26a69a", foreground="white")
style.configure("Treeview", font=("Helvetica", 10), rowheight=25, fieldbackground="white", background="#ffffff", foreground="#000000")

