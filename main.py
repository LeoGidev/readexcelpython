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
# Función para seleccionar una fila y habilitar la edición
def seleccionar_fila(event):
    item = tree.selection()[0]
    valores = tree.item(item, "values")

    # Limpiar las entradas antes de cargar la nueva fila seleccionada
    for entry in entries:
        entry.delete(0, tk.END)
        # Insertar los valores en los campos de edición
    for i, value in enumerate(valores):
        entries[i].insert(0, value)
# Función para actualizar los datos en el Treeview
def actualizar_fila():
    selected_item = tree.selection()[0]
    # Obtener los nuevos valores de los campos de entrada
    nuevos_valores = [entry.get() for entry in entries]
    # Actualizar la fila seleccionada en el Treeview
    tree.item(selected_item, values=nuevos_valores)
    # Actualizar el DataFrame
    index = int(tree.index(selected_item))
    df.iloc[index] = nuevos_valores
# Función para guardar el DataFrame modificado en un archivo Excel
def guardar_excel():
    global df



# Configurar la ventana principal de Tkinter
root = tk.Tk()
root.title("Lector de Excel")
root.geometry("800x400")
root.configure(bg="#e0f7fa")

# Estilos personalizados
style = ttk.Style()
style.theme_use("clam")
style.configure("Treeview.Heading", font=("Helvetica", 10, "bold"), background="#26a69a", foreground="white")
style.configure("Treeview", font=("Helvetica", 10), rowheight=25, fieldbackground="white", background="#ffffff", foreground="#000000")
style.map("Treeview", background=[("selected", "#81c784")])

# Crear botón para cargar archivo Excel
btn_cargar = tk.Button(root, text="Cargar Excel", command=cargar_excel, font=("Helvetica", 12, "bold"), bg="#26a69a", fg="white", padx=10, pady=5)
btn_cargar.pack(pady=20)

# Crear Treeview para mostrar los datos del archivo Excel
tree_frame = tk.Frame(root)
tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

tree = ttk.Treeview(tree_frame)
tree.pack(fill="both", expand=True, side="left")

# Barra de desplazamiento para el Treeview
scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=scrollbar.set)
scrollbar.pack(side="right", fill="y")

# Iniciar la aplicación
root.mainloop()
