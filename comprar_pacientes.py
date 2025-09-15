import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import os
from pathlib import Path

def clean_monto(value):
    """Limpia el monto para convertirlo a float (ej. '9.528,62 $' -> 9528.62)."""
    if pd.isna(value):
        return 0.0
    if isinstance(value, str):
        value = value.strip().replace('$', '').replace(' ', '').replace('.', '').replace(',', '.')
        try:
            return float(value)
        except ValueError:
            return 0.0
    return float(value)

def compare_records(user_file, hospital_files, output_dir=None):
    """Compara registro con los del hospital y genera un Excel con discrepancias."""
    try:
        # Leer registro del usuario
        user_df = pd.read_excel(user_file)
        user_df.columns = user_df.columns.str.strip().str.lower().str.replace(' ', '_')
        rename_dict_user = {
            'hc': 'HC',
            'paciente': 'Nombre',
            'fecha': 'Fecha',
            'hora': 'Hora',
            'monto': 'Monto'
        }
        user_df = user_df.rename(columns=rename_dict_user)
        user_df['HC'] = user_df['HC'].astype(str).str.strip()
        user_df['Fecha'] = pd.to_datetime(user_df['Fecha'], errors='coerce', dayfirst=True).dt.date
        if 'Monto' in user_df.columns:
            user_df['Monto'] = user_df['Monto'].apply(clean_monto)
        if 'Hora' not in user_df.columns:
            user_df['Hora'] = np.nan

        # Leer y combinar registros del hospital
        hospital_dfs = []
        for file in hospital_files:
            df = pd.read_excel(file)
            df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
            rename_dict_hospital = {
                'hc': 'HC',
                'historia': 'HC',
                'hono_impu1': 'Monto',
                'fecha': 'Fecha',
                'nombre': 'Nombre',
                'apellido_nombre': 'Doctor'
            }
            df = df.rename(columns=rename_dict_hospital)
            if 'HC' in df.columns:
                df['HC'] = df['HC'].astype(str).str.strip()
            if 'Fecha' in df.columns:
                df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce', dayfirst=True).dt.date
            if 'Monto' in df.columns:
                df['Monto'] = df['Monto'].apply(clean_monto)
            if 'Hora' not in df.columns:
                df['Hora'] = np.nan
            relevant_cols = [col for col in ['HC', 'Nombre', 'Fecha', 'Hora', 'Monto'] if col in df.columns]
            hospital_dfs.append(df[relevant_cols])
        
        hospital_df = pd.concat(hospital_dfs, ignore_index=True).drop_duplicates(subset=['HC', 'Fecha'])

        # Merge para encontrar discrepancias
        key_cols = ['HC', 'Fecha']
        merged = user_df[key_cols + [col for col in ['Nombre', 'Hora', 'Monto'] if col in user_df.columns]].merge(
            hospital_df[key_cols + [col for col in ['Nombre', 'Hora', 'Monto'] if col in hospital_df.columns]],
            on=key_cols, how='outer', suffixes=('_user', '_hospital'), indicator=True)

        # Extra en el registro propio (a favor)
        extra_user_mask = merged['_merge'] == 'left_only'
        extra_user = merged[extra_user_mask].copy()
        if not extra_user.empty:
            if 'Nombre_user' in extra_user.columns:
                extra_user['Nombre'] = extra_user['Nombre_user'].combine_first(extra_user.get('Nombre_hospital', pd.Series(dtype=object)))
            else:
                extra_user['Nombre'] = extra_user.get('Nombre', np.nan)
            if 'Monto_user' in extra_user.columns:
                extra_user['Monto'] = extra_user['Monto_user'].combine_first(extra_user.get('Monto_hospital', pd.Series(dtype=float)))
            else:
                extra_user['Monto'] = extra_user.get('Monto', 0.0)
            if 'Hora_user' in extra_user.columns:
                extra_user['Hora'] = extra_user['Hora_user'].combine_first(extra_user.get('Hora_hospital', pd.Series(dtype=object)))
            else:
                extra_user['Hora'] = extra_user.get('Hora', np.nan)
            extra_user = extra_user[key_cols + ['Nombre', 'Hora', 'Monto']]
        else:
            extra_user = pd.DataFrame(columns=['HC', 'Fecha', 'Nombre', 'Hora', 'Monto'])

        # Extra en el registro del hospital (en contra)
        extra_hospital_mask = merged['_merge'] == 'right_only'
        extra_hospital = merged[extra_hospital_mask].copy()
        if not extra_hospital.empty:
            if 'Nombre_hospital' in extra_hospital.columns:
                extra_hospital['Nombre'] = extra_hospital['Nombre_hospital'].combine_first(extra_hospital.get('Nombre_user', pd.Series(dtype=object)))
            else:
                extra_hospital['Nombre'] = extra_hospital.get('Nombre', np.nan)
            if 'Monto_hospital' in extra_hospital.columns:
                extra_hospital['Monto'] = extra_hospital['Monto_hospital'].combine_first(extra_hospital.get('Monto_user', pd.Series(dtype=float)))
            else:
                extra_hospital['Monto'] = extra_hospital.get('Monto', 0.0)
            if 'Hora_hospital' in extra_hospital.columns:
                extra_hospital['Hora'] = extra_hospital['Hora_hospital'].combine_first(extra_hospital.get('Hora_user', pd.Series(dtype=object)))
            else:
                extra_hospital['Hora'] = extra_hospital.get('Hora', np.nan)
            extra_hospital = extra_hospital[key_cols + ['Nombre', 'Hora', 'Monto']]
        else:
            extra_hospital = pd.DataFrame(columns=['HC', 'Fecha', 'Nombre', 'Hora', 'Monto'])

        # Ordenar
        cols = ['Nombre', 'Fecha', 'Hora', 'HC', 'Monto']
        extra_user = extra_user[cols].sort_values(by=['Nombre', 'Fecha'])
        extra_hospital = extra_hospital[cols].sort_values(by=['Nombre', 'Fecha'])

        # Resumen
        summary = pd.DataFrame({
            'Tipo': ['Extra en mi registro (a favor)', 'Extra en hospital (en contra)'],
            'Cantidad': [len(extra_user), len(extra_hospital)],
            'Monto Total': [extra_user['Monto'].sum() if len(extra_user) > 0 else 0.0, 
                           extra_hospital['Monto'].sum() if len(extra_hospital) > 0 else 0.0]
        })

        # Determinar ruta de salida
        if output_dir is None:
            output_dir = os.path.dirname(user_file)
        
        output_file = os.path.join(output_dir, 'discrepancias_pacientes.xlsx')

        # Exportar
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            extra_user.to_excel(writer, sheet_name='Extra_Mi_Registro', index=False)
            extra_hospital.to_excel(writer, sheet_name='Extra_Hospital', index=False)
            summary.to_excel(writer, sheet_name='Resumen', index=False)

        return output_file

    except Exception as e:
        raise Exception(f"Error procesando archivos: {str(e)}")

class ComparadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparador de Pacientes 1.3.2 - by RenzoRossiBrun")
        self.root.geometry("600x500")
        
        # Variables
        self.user_file = tk.StringVar()
        self.hospital_files = []
        
        self.setup_ui()
    
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        title_label = ttk.Label(main_frame, text="Comparador de Registros de Pacientes", 
                               font=('Arial', 14, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Archivo del usuario
        ttk.Label(main_frame, text="1. Selecciona tu archivo de registro:").grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=(0, 5))
        
        user_frame = ttk.Frame(main_frame)
        user_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Entry(user_frame, textvariable=self.user_file, width=50, state='readonly').grid(row=0, column=0, padx=(0, 10))
        ttk.Button(user_frame, text="Examinar...", command=self.select_user_file).grid(row=0, column=1)
        
        # Archivos del hospital
        ttk.Label(main_frame, text="2. Selecciona archivos del hospital:").grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(10, 5))
        
        hospital_frame = ttk.Frame(main_frame)
        hospital_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(hospital_frame, text="Agregar archivos del hospital", 
                  command=self.select_hospital_files).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(hospital_frame, text="Limpiar lista", 
                  command=self.clear_hospital_files).grid(row=0, column=1)
        
        # Lista de archivos del hospital
        self.hospital_listbox = tk.Listbox(main_frame, height=6)
        self.hospital_listbox.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Scrollbar para la lista
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.hospital_listbox.yview)
        scrollbar.grid(row=5, column=3, sticky=(tk.N, tk.S))
        self.hospital_listbox.configure(yscrollcommand=scrollbar.set)
        
        # Botón de procesamiento
        process_frame = ttk.Frame(main_frame)
        process_frame.grid(row=6, column=0, columnspan=3, pady=20)
        
        self.process_btn = ttk.Button(process_frame, text="Comparar y Generar Reporte", 
                                     command=self.process_files, style='Accent.TButton')
        self.process_btn.pack()
        
        # Barra de progreso
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Mensaje de estado
        self.status_label = ttk.Label(main_frame, text="Selecciona los archivos para comenzar")
        self.status_label.grid(row=8, column=0, columnspan=3, pady=(10, 0))
        
        # Configurar redimensionamiento
        main_frame.columnconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def select_user_file(self):
        filename = filedialog.askopenfilename(
            title="Selecciona tu archivo de registro",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        if filename:
            self.user_file.set(filename)
            self.update_status("Archivo de usuario seleccionado")
    
    def select_hospital_files(self):
        filenames = filedialog.askopenfilenames(
            title="Selecciona archivos del hospital",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        if filenames:
            for filename in filenames:
                if filename not in self.hospital_files:
                    self.hospital_files.append(filename)
                    self.hospital_listbox.insert(tk.END, os.path.basename(filename))
            self.update_status(f"Archivos del hospital: {len(self.hospital_files)}")
    
    def clear_hospital_files(self):
        self.hospital_files.clear()
        self.hospital_listbox.delete(0, tk.END)
        self.update_status("Lista de archivos del hospital limpiada")
    
    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def process_files(self):
        # Validar archivos
        if not self.user_file.get():
            messagebox.showerror("Error", "Selecciona tu archivo de registro")
            return
        
        if not self.hospital_files:
            messagebox.showerror("Error", "Selecciona al menos un archivo del hospital")
            return
        
        # Mostrar progreso
        self.progress.start(10)
        self.process_btn.config(state='disabled')
        self.update_status("Procesando archivos...")
        
        try:
            # Procesar archivos
            output_file = compare_records(self.user_file.get(), self.hospital_files)
            
            # Mostrar resultado
            messagebox.showinfo(
                "¡Completado!", 
                f"Reporte generado exitosamente:\n\n{output_file}\n\nEl archivo se guardó en la misma carpeta que tu registro."
            )
            
            # Preguntar si abrir el archivo
            if messagebox.askyesno("Abrir archivo", "¿Deseas abrir el archivo generado?"):
                os.startfile(output_file)  # Windows
            
            self.update_status("Proceso completado exitosamente")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al procesar los archivos:\n\n{str(e)}")
            self.update_status("Error en el procesamiento")
        
        finally:
            self.progress.stop()
            self.process_btn.config(state='normal')

def main():
    root = tk.Tk()
    app = ComparadorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()