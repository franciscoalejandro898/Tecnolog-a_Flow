import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.style import Style
import tkinter as tk
from tkinter import messagebox  # Importa messagebox para mostrar ventanas de diálogo
import os

# Crear la aplicación y configurar estilo
app = ttk.Window()
app.geometry("600x500")
style = Style(theme="minty")

# Etiqueta superior
label = ttk.Label(app, text="Inserta tu ruta para Mapaear los archivos")
label.pack(pady=30)
label.config(font=("Ubuntu", 20, "bold"))

# Entrada para la ruta
name_frame = ttk.Frame(app)
name_frame.pack(pady=15, padx=10, fill="x")
ttk.Label(name_frame, text="Ruta    ").pack(side=tk.LEFT, padx=5)

ruta_var = tk.StringVar()  # Variable para almacenar la ruta
ruta_entry = ttk.Entry(name_frame, textvariable=ruta_var)
ruta_entry.pack(side=tk.LEFT, fill="x", expand=True, padx=5)

# Función para procesar la ruta
def procesar_ruta():
    ruta = ruta_var.get()  # Obtener la ruta de la variable
    if os.path.exists(ruta):
        os.startfile(ruta)  # Abrir la carpeta en el explorador de archivos
    else:
        # Mostrar mensaje de error
        messagebox.showerror("Error", "La ruta no existe", parent=app)

# Botones
button_frame = ttk.Frame(app)
button_frame.pack(pady=50, padx=10, fill="x")
ttk.Button(button_frame, text="Procesar", bootstyle=SUCCESS, command=procesar_ruta).pack(side=tk.LEFT, padx=10)
ttk.Button(button_frame, text="Cancelar", bootstyle=SECONDARY).pack(side=tk.LEFT, padx=10)

app.mainloop()

