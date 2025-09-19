import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, Scrollbar
import pdfplumber
import pandas as pd

def convert_pdf_to_excel(pdf_path):
    all_tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table:  # Asegura que la tabla no esté vacía
                    # Asume que la primera fila es el encabezado
                    df = pd.DataFrame(table[1:], columns=table[0])
                    all_tables.append(df)
    
    if all_tables:
        excel_path = pdf_path.replace('.pdf', '.xlsx')
        with pd.ExcelWriter(excel_path) as writer:
            for i, df in enumerate(all_tables):
                df.to_excel(writer, sheet_name=f'Tabla_{i+1}', index=False)
        return f"Convertido {pdf_path} a {excel_path}"
    else:
        return f"No se encontraron tablas en {pdf_path}"

def select_files():
    files = filedialog.askopenfilenames(
        title="Seleccionar Archivos PDF",
        filetypes=[("Archivos PDF", "*.pdf")]
    )
    if files:
        file_list.delete(0, tk.END)
        for file in files:
            file_list.insert(tk.END, file)
        status_label.config(text=f"{len(files)} archivos seleccionados.")

def convert_files():
    files = file_list.get(0, tk.END)
    if not files:
        messagebox.showwarning("Sin Archivos", "Por favor, selecciona archivos PDF primero.")
        return
    
    results = []
    for pdf_file in files:
        result = convert_pdf_to_excel(pdf_file)
        results.append(result)
    
    messagebox.showinfo("Conversión Completada", "\n".join(results))
    status_label.config(text="Conversión completada. Listo para nuevos archivos.")
    file_list.delete(0, tk.END)

# Configuración de la GUI
root = tk.Tk()
root.title("Convertidor de PDF a Excel")
root.geometry("600x400")

# Instrucciones
instructions = tk.Label(root, text="¡Bienvenido! Esta herramienta convierte archivos PDF a Excel.\n"
                                   "1. Haz clic en 'Seleccionar PDFs' para elegir tus archivos.\n"
                                   "2. Los archivos seleccionados aparecerán en la lista abajo.\n"
                                   "3. Haz clic en 'Convertir' para procesarlos.\n"
                                   "Cada PDF se convertirá en un archivo Excel separado en el mismo directorio.",
                        justify=tk.LEFT, padx=10, pady=10)
instructions.pack(anchor=tk.W)

# Botón de selección de archivos
select_button = tk.Button(root, text="Seleccionar PDFs", command=select_files)
select_button.pack(pady=10)

# Lista de archivos
frame = tk.Frame(root)
frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

scrollbar = Scrollbar(frame)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

file_list = Listbox(frame, yscrollcommand=scrollbar.set)
file_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar.config(command=file_list.yview)

# Botón de conversión
convert_button = tk.Button(root, text="Convertir", command=convert_files)
convert_button.pack(pady=10)

# Etiqueta de estado
status_label = tk.Label(root, text="Aún no se han seleccionado archivos.")
status_label.pack(pady=10)

# Verificar si se ejecuta desde línea de comandos con argumentos
if len(sys.argv) > 1 and sys.argv[1] != '':
    # Modo CLI
    for pdf_file in sys.argv[1:]:
        print(convert_pdf_to_excel(pdf_file))
else:
    # Modo GUI
    root.mainloop()