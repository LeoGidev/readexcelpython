import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

# Variables globales
df = None
archivo_path = None  # Guardar la ruta del archivo cargado
entradas = []

# Función para leer el archivo Excel
def cargar_excel():
    global df, archivo_path
    archivo_path = filedialog.askopenfilename(
        filetypes=[("Archivo Excel", "*.xlsx *.xls")]
    )
    
    if archivo_path:
        try:
            df = pd.read_excel(archivo_path)
            mostrar_datos(df)
            crear_campos_entrada(df)  # Crear campos de entrada después de cargar el archivo Excel
            actualizar_contador(df)  # Actualizar el contador de filas y columnas
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo: {str(e)}")

# Función para guardar el DataFrame actualizado en el archivo Excel
def guardar_excel():
    global df, archivo_path
    if df is not None and archivo_path is not None:
        try:
            df.to_excel(archivo_path, index=False)  # Guarda los cambios en el archivo original
            messagebox.showinfo("Éxito", "El archivo se ha guardado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo: {str(e)}")
    else:
        messagebox.showerror("Error", "No hay datos cargados para guardar.")

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

# Función para seleccionar una fila del Treeview
def seleccionar_fila(event):
    selected_item = tree.selection()
    if selected_item:
        index = tree.index(selected_item[0])
        valores_fila = df.iloc[index].tolist()

        # Limpiar los campos de entrada
        for i in range(len(entradas)):
            entradas[i].delete(0, tk.END)
            entradas[i].insert(0, valores_fila[i])

# Función para actualizar una fila en el DataFrame y en el Treeview
def actualizar_fila():
    global df
    selected_item = tree.selection()
    if selected_item:
        index = tree.index(selected_item[0])

        # Obtener los valores nuevos desde los campos de entrada
        nuevos_valores = [entrada.get() for entrada in entradas]

        if len(nuevos_valores) == len(df.columns):  # Verifica que coincida el número de columnas
            df.iloc[index] = nuevos_valores
            mostrar_datos(df)  # Refresca los datos en el Treeview
        else:
            messagebox.showerror("Error", "El número de columnas no coincide con el número de valores.")

# Función para crear los campos de entrada dinámicamente según las columnas del Excel
def crear_campos_entrada(df):
    global entradas
    # Limpiar el frame de entrada si ya hay entradas creadas
    for widget in entrada_frame.winfo_children():
        widget.destroy()

    entradas = []
    for col in df.columns:
        entrada = tk.Entry(entrada_frame, font=("Helvetica", 10))
        entrada.pack(side="left", padx=5)
        entradas.append(entrada)
# Función para actualizar el contador de filas y columnas
def actualizar_contador(df):
    num_filas = len(df)
    num_columnas = len(df.columns)
    contador_label.config(text=f"Filas: {num_filas}, Columnas: {num_columnas}")


# Configurar la ventana principal de Tkinter
root = tk.Tk()
root.title("Lector de Excel")
root.geometry("800x600")
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

# Vincular la selección del Treeview con la función de selección de fila
tree.bind("<<TreeviewSelect>>", seleccionar_fila)

# Crear frame para los campos de entrada
entrada_frame = tk.Frame(root)
entrada_frame.pack(fill="x", padx=10, pady=10)

# Botón para actualizar la fila seleccionada
btn_actualizar = tk.Button(root, text="Actualizar Fila", command=actualizar_fila, font=("Helvetica", 12, "bold"), bg="#26a69a", fg="white", padx=10, pady=5)
btn_actualizar.pack(pady=10)

# Botón para guardar los cambios
btn_guardar = tk.Button(root, text="Guardar Cambios", command=guardar_excel, font=("Helvetica", 12, "bold"), bg="#26a69a", fg="white", padx=10, pady=5)
btn_guardar.pack(pady=10)



# Iniciar la aplicación
root.mainloop()



