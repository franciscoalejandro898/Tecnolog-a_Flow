# main.py
import tkinter as tk
from tkinter import filedialog, messagebox
from ttkbootstrap import Style
from ttkbootstrap.constants import *
import ttkbootstrap as ttk  # Importar ttk desde ttkbootstrap
import pandas as pd
from PIL import Image, ImageTk  # Para manejar imágenes
from process import process_excel_files

class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Procesador de Archivos Excel")

        # Configurar tamaño y centrar ventana
        window_width = 900
        window_height = 600
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        position_top = int(screen_height / 2 - window_height / 2)
        position_right = int(screen_width / 2 - window_width / 2)
        root.geometry(f"{window_width}x{window_height}+{position_right}+{position_top}")

        # Aplicar estilo de ttkbootstrap
        self.style = Style(theme="minty")

        # Crear estilos personalizados
        self.style.configure('Custom.TButton', font=('Helvetica', 12, 'bold'), foreground='black', borderwidth=2, relief="solid", bordercolor='#02a4d3')
        self.style.map('Custom.TButton', 
                       background=[('active', '#d9d9d9'), ('!active', '#f0f0f0')],
                       relief=[('pressed', 'groove'), ('!pressed', 'solid')])

        # Frame para el contenido
        content_frame = ttk.Frame(root)
        content_frame.pack(pady=20, padx=10, fill=tk.BOTH, expand=True)

        # Frame para el logo
        logo_frame = ttk.Frame(content_frame)
        logo_frame.pack(side=tk.LEFT, anchor="n", padx=0)

        # Cargar y mostrar el logo
        self.logo_image = Image.open("logo.png")
        self.logo_image = self.logo_image.resize((120, 120), Image.LANCZOS)  # Ajustar el tamaño del logo
        self.logo_photo = ImageTk.PhotoImage(self.logo_image)
        logo_label = ttk.Label(logo_frame, image=self.logo_photo)
        logo_label.pack()

        # Frame para los widgets centrados
        center_frame = ttk.Frame(content_frame)
        center_frame.pack(side=tk.LEFT, anchor="n", padx=20, expand=True)

        # Sub-frame para centrar los widgets verticalmente
        inner_widget_frame = ttk.Frame(center_frame)
        inner_widget_frame.pack(anchor="center")

        # Etiqueta de título
        title_label = ttk.Label(inner_widget_frame, text="Procesador de Archivos Excel", font=("Helvetica", 18))
        title_label.pack(pady=10)

        # Botón para seleccionar directorio
        select_button = ttk.Button(inner_widget_frame, text="Seleccionar Directorio", command=self.select_directory, style='Custom.TButton')
        select_button.pack(pady=10)

        # Botón para guardar archivo
        self.save_button = ttk.Button(inner_widget_frame, text="Guardar Archivo", command=self.save_file, style='Custom.TButton')
        self.save_button.pack(pady=10)
        self.save_button.config(state=tk.DISABLED)

        # Botón para restablecer
        reset_button = ttk.Button(inner_widget_frame, text="Restablecer", command=self.reset_app, style='Custom.TButton')
        reset_button.pack(pady=10)

        # Frame para Treeview con barra de desplazamiento
        self.frame = ttk.Frame(root)
        self.frame.pack(pady=20, fill=tk.BOTH, expand=True)

        # Crear Treeview con barra de desplazamiento
        self.tree_scroll = ttk.Scrollbar(self.frame)
        self.tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(self.frame, yscrollcommand=self.tree_scroll.set)
        self.tree.pack(pady=20, fill=tk.BOTH, expand=True)

        self.tree_scroll.config(command=self.tree.yview)

        self.data = None

    def select_directory(self):
        directory_path = filedialog.askdirectory()
        if directory_path:
            try:
                self.data = process_excel_files(directory_path)
                if not self.data.empty:
                    self.preview_data(self.data)
                    messagebox.showinfo("Proceso Completado", f"Se han procesado los archivos en el directorio: {directory_path}")
                    self.save_button.config(state=tk.NORMAL)
                else:
                    messagebox.showwarning("Sin Datos", "No se encontraron datos procesables en los archivos.")
                    self.save_button.config(state=tk.DISABLED)
            except Exception as e:
                messagebox.showerror("Error", f"Ocurrió un error: {str(e)}")
                self.save_button.config(state=tk.DISABLED)

    def preview_data(self, data):
        for widget in self.frame.winfo_children():
            widget.destroy()

        self.tree_scroll = ttk.Scrollbar(self.frame)
        self.tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(self.frame, yscrollcommand=self.tree_scroll.set)
        self.tree.pack(pady=20, fill=tk.BOTH, expand=True)

        self.tree_scroll.config(command=self.tree.yview)

        # Define columnas
        self.tree["column"] = list(data.columns)
        self.tree["show"] = "headings"

        for column in self.tree["columns"]:
            self.tree.heading(column, text=column)
            self.tree.column(column, width=100, anchor="center")

        # Inserta filas
        for _, row in data.iterrows():
            self.tree.insert("", "end", values=list(row))

    def save_file(self):
        if self.data is not None:
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if save_path:
                self.data.to_excel(save_path, index=False)
                messagebox.showinfo("Archivo Guardado", f"El archivo ha sido guardado en: {save_path}")

    def reset_app(self):
        for widget in self.frame.winfo_children():
            widget.destroy()
        self.data = None
        self.save_button.config(state=tk.DISABLED)
        messagebox.showinfo("Restablecer", "La aplicación ha sido restablecida")

# Configuración de la ventana principal de Tkinter
root = tk.Tk()
app = ExcelProcessorApp(root)
root.mainloop()
