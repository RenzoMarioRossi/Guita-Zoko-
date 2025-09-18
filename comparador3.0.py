"""
#   V1

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import os
import re

class PatientControlApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Control de Pacientes - Hospital vs Usuario")
        self.root.geometry("800x600")
        
        # Variables para almacenar archivos
        self.hospital_files = []
        self.user_files = []
        
        self.setup_ui()
    
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        title_label = ttk.Label(main_frame, text="Sistema de Control de Pacientes", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)
        
        # Sección archivos del hospital
        hospital_frame = ttk.LabelFrame(main_frame, text="Archivos del Hospital", padding="10")
        hospital_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(hospital_frame, text="Agregar Archivos del Hospital", 
                  command=self.add_hospital_files).grid(row=0, column=0, padx=5)
        ttk.Button(hospital_frame, text="Limpiar Lista", 
                  command=self.clear_hospital_files).grid(row=0, column=1, padx=5)
        
        self.hospital_listbox = tk.Listbox(hospital_frame, height=6)
        self.hospital_listbox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        hospital_scrollbar = ttk.Scrollbar(hospital_frame, orient="vertical", 
                                         command=self.hospital_listbox.yview)
        hospital_scrollbar.grid(row=1, column=2, sticky=(tk.N, tk.S))
        self.hospital_listbox.configure(yscrollcommand=hospital_scrollbar.set)
        
        # Sección archivos del usuario
        user_frame = ttk.LabelFrame(main_frame, text="Archivos del Usuario", padding="10")
        user_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(user_frame, text="Agregar Archivos del Usuario", 
                  command=self.add_user_files).grid(row=0, column=0, padx=5)
        ttk.Button(user_frame, text="Limpiar Lista", 
                  command=self.clear_user_files).grid(row=0, column=1, padx=5)
        
        self.user_listbox = tk.Listbox(user_frame, height=6)
        self.user_listbox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        user_scrollbar = ttk.Scrollbar(user_frame, orient="vertical", 
                                     command=self.user_listbox.yview)
        user_scrollbar.grid(row=1, column=2, sticky=(tk.N, tk.S))
        self.user_listbox.configure(yscrollcommand=user_scrollbar.set)
        
        # Botones de acción
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=3, column=0, columnspan=3, pady=20)
        
        ttk.Button(action_frame, text="Procesar Archivos", 
                  command=self.process_files, style="Accent.TButton").grid(row=0, column=0, padx=10)
        ttk.Button(action_frame, text="Salir", 
                  command=self.root.quit).grid(row=0, column=1, padx=10)
        
        # Barra de progreso
        self.progress = ttk.Progressbar(main_frame, length=300, mode='determinate')
        self.progress.grid(row=4, column=0, columnspan=3, pady=10)
        
        # Label de estado
        self.status_label = ttk.Label(main_frame, text="Listo para procesar archivos")
        self.status_label.grid(row=5, column=0, columnspan=3, pady=5)
        
        # Configurar pesos para redimensionamiento
        main_frame.columnconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def add_hospital_files(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos del hospital",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        for file in files:
            if file not in self.hospital_files:
                self.hospital_files.append(file)
                self.hospital_listbox.insert(tk.END, os.path.basename(file))
    
    def add_user_files(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos del usuario",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        for file in files:
            if file not in self.user_files:
                self.user_files.append(file)
                self.user_listbox.insert(tk.END, os.path.basename(file))
    
    def clear_hospital_files(self):
        self.hospital_files.clear()
        self.hospital_listbox.delete(0, tk.END)
    
    def clear_user_files(self):
        self.user_files.clear()
        self.user_listbox.delete(0, tk.END)
    
    def normalize_text(self, text):
        #Normaliza texto para comparación insensible a mayúsculas/minúsculas

        if pd.isna(text) or text is None:
            return ""
        return str(text).strip().upper()
    
    def normalize_date(self, date_value):
        #Normaliza fechas a formato comparable

        if pd.isna(date_value) or date_value is None:
            return None
        
        # Si ya es datetime
        if isinstance(date_value, datetime):
            return date_value.date()
        
        # Si es string, intentar parsearlo
        if isinstance(date_value, str):
            date_value = date_value.strip()
            # Intentar varios formatos comunes
            formats = ['%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d', '%d/%m/%y', '%d-%m-%y']
            for fmt in formats:
                try:
                    return datetime.strptime(date_value, fmt).date()
                except ValueError:
                    continue
        
        return None
    
    def find_hc_column(self, df):
        #Encuentra la columna de historia clínica

        hc_patterns = ['hc', 'historia', 'hist', 'h.c', 'historia clinica', 'historia clínica']
        
        for col in df.columns:
            col_name = str(col).lower().strip()
            for pattern in hc_patterns:
                if pattern in col_name:
                    return col
        return None
    
    def find_patient_column(self, df):
        #Encuentra la columna de paciente/nombre

        patient_patterns = ['paciente', 'nombre', 'apellido', 'patient', 'name']
        
        for col in df.columns:
            col_name = str(col).lower().strip()
            for pattern in patient_patterns:
                if pattern in col_name:
                    return col
        return None
    
    def find_date_column(self, df):
        #Encuentra la columna de fecha
        date_patterns = ['fecha', 'date', 'dia', 'day']
        
        for col in df.columns:
            col_name = str(col).lower().strip()
            for pattern in date_patterns:
                if pattern in col_name:
                    return col
        return None
    
    def load_excel_file(self, filepath):
        #Carga un archivo Excel y retorna DataFrame con columnas identificadas
        try:
            # Intentar leer el archivo
            df = pd.read_excel(filepath)
            
            if df.empty:
                return None, "Archivo vacío"
            
            # Identificar columnas importantes
            hc_col = self.find_hc_column(df)
            patient_col = self.find_patient_column(df)
            date_col = self.find_date_column(df)
            
            return {
                'dataframe': df,
                'hc_column': hc_col,
                'patient_column': patient_col,
                'date_column': date_col,
                'filename': os.path.basename(filepath)
            }, None
            
        except Exception as e:
            return None, f"Error al leer {filepath}: {str(e)}"
    
    def search_patient_in_user_files(self, hospital_row, hospital_info, user_data_list):
        #Busca un paciente del hospital en los archivos del usuario
        
        hospital_hc = None
        hospital_patient = None
        hospital_date = None
        
        # Obtener datos del hospital
        if hospital_info['hc_column'] and hospital_info['hc_column'] in hospital_row:
            hospital_hc = self.normalize_text(hospital_row[hospital_info['hc_column']])
        
        if hospital_info['patient_column'] and hospital_info['patient_column'] in hospital_row:
            hospital_patient = self.normalize_text(hospital_row[hospital_info['patient_column']])
        
        if hospital_info['date_column'] and hospital_info['date_column'] in hospital_row:
            hospital_date = self.normalize_date(hospital_row[hospital_info['date_column']])
        
        # Buscar en archivos de usuario
        for user_info in user_data_list:
            user_df = user_info['dataframe']
            
            for _, user_row in user_df.iterrows():
                user_hc = None
                user_patient = None
                user_date = None
                
                # Obtener datos del usuario
                if user_info['hc_column'] and user_info['hc_column'] in user_row:
                    user_hc = self.normalize_text(user_row[user_info['hc_column']])
                
                if user_info['patient_column'] and user_info['patient_column'] in user_row:
                    user_patient = self.normalize_text(user_row[user_info['patient_column']])
                
                if user_info['date_column'] and user_info['date_column'] in user_row:
                    user_date = self.normalize_date(user_row[user_info['date_column']])
                
                # Comparar - priorizar HC si existe, sino usar nombre
                match_found = False
                
                # Si ambos tienen HC y no están vacías
                if hospital_hc and user_hc and hospital_hc != "" and user_hc != "":
                    if hospital_hc == user_hc:
                        # HC coincide, verificar fecha
                        if hospital_date and user_date and hospital_date == user_date:
                            return True  # Encontrado con fecha coincidente
                        elif not hospital_date or not user_date:
                            return True  # Al menos uno no tiene fecha, considerar encontrado
                        # Si tienen fechas diferentes, continuar buscando
                
                # Si no hay HC o no coincide, comparar por nombre
                elif hospital_patient and user_patient and hospital_patient != "" and user_patient != "":
                    if hospital_patient == user_patient:
                        # Nombre coincide, verificar fecha
                        if hospital_date and user_date and hospital_date == user_date:
                            return True  # Encontrado con fecha coincidente
                        elif not hospital_date or not user_date:
                            return True  # Al menos uno no tiene fecha, considerar encontrado
                        # Si tienen fechas diferentes, continuar buscando
        
        return False  # No encontrado
    
    def process_files(self):
        if not self.hospital_files:
            messagebox.showerror("Error", "Debe seleccionar al menos un archivo del hospital")
            return
        
        if not self.user_files:
            messagebox.showerror("Error", "Debe seleccionar al menos un archivo del usuario")
            return
        
        try:
            self.status_label.config(text="Cargando archivos...")
            self.progress['value'] = 0
            self.root.update()
            
            # Cargar archivos del hospital
            hospital_data_list = []
            for i, filepath in enumerate(self.hospital_files):
                self.status_label.config(text=f"Cargando archivo del hospital {i+1}/{len(self.hospital_files)}")
                self.root.update()
                
                data, error = self.load_excel_file(filepath)
                if error:
                    messagebox.showerror("Error", error)
                    return
                hospital_data_list.append(data)
            
            # Cargar archivos del usuario
            user_data_list = []
            for i, filepath in enumerate(self.user_files):
                self.status_label.config(text=f"Cargando archivo del usuario {i+1}/{len(self.user_files)}")
                self.root.update()
                
                data, error = self.load_excel_file(filepath)
                if error:
                    messagebox.showerror("Error", error)
                    return
                user_data_list.append(data)
            
            # Procesar comparaciones
            missing_patients = []
            total_rows = sum(len(data['dataframe']) for data in hospital_data_list)
            processed_rows = 0
            
            self.status_label.config(text="Procesando comparaciones...")
            
            for hospital_info in hospital_data_list:
                hospital_df = hospital_info['dataframe']
                
                for _, hospital_row in hospital_df.iterrows():
                    processed_rows += 1
                    progress_percent = (processed_rows / total_rows) * 100
                    self.progress['value'] = progress_percent
                    
                    if processed_rows % 10 == 0:  # Actualizar cada 10 filas
                        self.status_label.config(text=f"Procesando: {processed_rows}/{total_rows}")
                        self.root.update()
                    
                    # Buscar paciente en archivos de usuario
                    found = self.search_patient_in_user_files(hospital_row, hospital_info, user_data_list)
                    
                    if not found:
                        # Crear registro del paciente faltante
                        missing_record = {}
                        for col in hospital_df.columns:
                            missing_record[col] = hospital_row[col]
                        missing_record['Archivo_Origen'] = hospital_info['filename']
                        missing_patients.append(missing_record)
            
            # Generar reporte
            if missing_patients:
                self.generate_report(missing_patients)
                messagebox.showinfo("Proceso Completado", 
                                  f"Proceso completado. Se encontraron {len(missing_patients)} pacientes no registrados.\n"
                                  f"El reporte se ha guardado como 'reporte_pacientes_faltantes.xlsx'")
            else:
                messagebox.showinfo("Proceso Completado", 
                                  "Proceso completado. Todos los pacientes del hospital están registrados en los archivos del usuario.")
            
            self.progress['value'] = 100
            self.status_label.config(text="Proceso completado")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error durante el procesamiento: {str(e)}")
            self.status_label.config(text="Error en el procesamiento")
    
    def generate_report(self, missing_patients):
        #Genera el reporte de pacientes faltantes

        if not missing_patients:
            return
        
        # Crear DataFrame con los pacientes faltantes
        df_report = pd.DataFrame(missing_patients)
        
        # Guardar archivo
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"reporte_pacientes_faltantes_{timestamp}.xlsx"
        
        try:
            df_report.to_excel(filename, index=False)
            self.status_label.config(text=f"Reporte guardado como: {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el reporte: {str(e)}")

def main():
    root = tk.Tk()
    app = PatientControlApp(root)
    
    # Configurar estilo
    style = ttk.Style()
    style.theme_use('clam')
    
    root.mainloop()

if __name__ == "__main__":
    main()

"""

"""

#   V2

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import os
import re

class PatientControlApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Control de Pacientes - Hospital vs Usuario")
        self.root.geometry("800x600")
        
        # Variables para almacenar archivos
        self.hospital_files = []
        self.user_files = []
        
        self.setup_ui()
    
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        title_label = ttk.Label(main_frame, text="Sistema de Control de Pacientes", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)
        
        # Sección archivos del hospital
        hospital_frame = ttk.LabelFrame(main_frame, text="Archivos del Hospital", padding="10")
        hospital_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(hospital_frame, text="Agregar Archivos del Hospital", 
                  command=self.add_hospital_files).grid(row=0, column=0, padx=5)
        ttk.Button(hospital_frame, text="Limpiar Lista", 
                  command=self.clear_hospital_files).grid(row=0, column=1, padx=5)
        
        self.hospital_listbox = tk.Listbox(hospital_frame, height=6)
        self.hospital_listbox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        hospital_scrollbar = ttk.Scrollbar(hospital_frame, orient="vertical", 
                                         command=self.hospital_listbox.yview)
        hospital_scrollbar.grid(row=1, column=2, sticky=(tk.N, tk.S))
        self.hospital_listbox.configure(yscrollcommand=hospital_scrollbar.set)
        
        # Sección archivos del usuario
        user_frame = ttk.LabelFrame(main_frame, text="Archivos del Usuario", padding="10")
        user_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(user_frame, text="Agregar Archivos del Usuario", 
                  command=self.add_user_files).grid(row=0, column=0, padx=5)
        ttk.Button(user_frame, text="Limpiar Lista", 
                  command=self.clear_user_files).grid(row=0, column=1, padx=5)
        
        self.user_listbox = tk.Listbox(user_frame, height=6)
        self.user_listbox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        user_scrollbar = ttk.Scrollbar(user_frame, orient="vertical", 
                                     command=self.user_listbox.yview)
        user_scrollbar.grid(row=1, column=2, sticky=(tk.N, tk.S))
        self.user_listbox.configure(yscrollcommand=user_scrollbar.set)
        
        # Botones de acción
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=3, column=0, columnspan=3, pady=20)
        
        ttk.Button(action_frame, text="Procesar Archivos", 
                  command=self.process_files, style="Accent.TButton").grid(row=0, column=0, padx=10)
        ttk.Button(action_frame, text="Salir", 
                  command=self.root.quit).grid(row=0, column=1, padx=10)
        
        # Barra de progreso
        self.progress = ttk.Progressbar(main_frame, length=300, mode='determinate')
        self.progress.grid(row=4, column=0, columnspan=3, pady=10)
        
        # Label de estado
        self.status_label = ttk.Label(main_frame, text="Listo para procesar archivos")
        self.status_label.grid(row=5, column=0, columnspan=3, pady=5)
        
        # Configurar pesos para redimensionamiento
        main_frame.columnconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def add_hospital_files(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos del hospital",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        for file in files:
            if file not in self.hospital_files:
                self.hospital_files.append(file)
                self.hospital_listbox.insert(tk.END, os.path.basename(file))
    
    def add_user_files(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos del usuario",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        for file in files:
            if file not in self.user_files:
                self.user_files.append(file)
                self.user_listbox.insert(tk.END, os.path.basename(file))
    
    def clear_hospital_files(self):
        self.hospital_files.clear()
        self.hospital_listbox.delete(0, tk.END)
    
    def clear_user_files(self):
        self.user_files.clear()
        self.user_listbox.delete(0, tk.END)
    
    def normalize_text(self, text):
        #Normaliza texto para comparación insensible a mayúsculas/minúsculas
        if pd.isna(text) or text is None:
            return ""
        return str(text).strip().upper()
    
    def normalize_date(self, date_value):
        #Normaliza fechas a formato comparable
        if pd.isna(date_value) or date_value is None:
            return None
        
        # Si ya es datetime
        if isinstance(date_value, datetime):
            return date_value.date()
        
        # Si es string, intentar parsearlo
        if isinstance(date_value, str):
            date_value = date_value.strip()
            # Intentar varios formatos comunes
            formats = ['%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d', '%d/%m/%y', '%d-%m-%y']
            for fmt in formats:
                try:
                    return datetime.strptime(date_value, fmt).date()
                except ValueError:
                    continue
        
        return None
    
    def find_hc_column(self, df):
        #Encuentra la columna de historia clínica
        hc_patterns = ['hc', 'historia', 'hist', 'h.c', 'historia clinica', 'historia clínica']
        
        for col in df.columns:
            col_name = str(col).lower().strip()
            for pattern in hc_patterns:
                if pattern in col_name:
                    return col
        return None
    
    def find_patient_column(self, df):
        #Encuentra la columna de paciente/nombre
        patient_patterns = ['paciente', 'nombre', 'apellido', 'patient', 'name']
        
        for col in df.columns:
            col_name = str(col).lower().strip()
            for pattern in patient_patterns:
                if pattern in col_name:
                    return col
        return None
    
    def find_date_column(self, df):
        #Encuentra la columna de fecha
        date_patterns = ['fecha', 'date', 'dia', 'day']
        
        for col in df.columns:
            col_name = str(col).lower().strip()
            for pattern in date_patterns:
                if pattern in col_name:
                    return col
        return None
    
    def load_excel_file(self, filepath):
        #Carga un archivo Excel y retorna DataFrame con columnas identificadas
        try:
            # Intentar leer el archivo
            df = pd.read_excel(filepath)
            
            if df.empty:
                return None, "Archivo vacío"
            
            # Identificar columnas importantes
            hc_col = self.find_hc_column(df)
            patient_col = self.find_patient_column(df)
            date_col = self.find_date_column(df)
            
            return {
                'dataframe': df,
                'hc_column': hc_col,
                'patient_column': patient_col,
                'date_column': date_col,
                'filename': os.path.basename(filepath)
            }, None
            
        except Exception as e:
            return None, f"Error al leer {filepath}: {str(e)}"
    
    def search_patient_in_user_files(self, hospital_row, hospital_info, user_data_list):
        #Busca un paciente del hospital en los archivos del usuario
        
        hospital_hc = None
        hospital_patient = None
        hospital_date = None
        
        # Obtener datos del hospital
        if hospital_info['hc_column'] and hospital_info['hc_column'] in hospital_row:
            hospital_hc = self.normalize_text(hospital_row[hospital_info['hc_column']])
        
        if hospital_info['patient_column'] and hospital_info['patient_column'] in hospital_row:
            hospital_patient = self.normalize_text(hospital_row[hospital_info['patient_column']])
        
        if hospital_info['date_column'] and hospital_info['date_column'] in hospital_row:
            hospital_date = self.normalize_date(hospital_row[hospital_info['date_column']])
        
        # Buscar en archivos de usuario
        for user_info in user_data_list:
            user_df = user_info['dataframe']
            
            for _, user_row in user_df.iterrows():
                user_hc = None
                user_patient = None
                user_date = None
                
                # Obtener datos del usuario
                if user_info['hc_column'] and user_info['hc_column'] in user_row:
                    user_hc = self.normalize_text(user_row[user_info['hc_column']])
                
                if user_info['patient_column'] and user_info['patient_column'] in user_row:
                    user_patient = self.normalize_text(user_row[user_info['patient_column']])
                
                if user_info['date_column'] and user_info['date_column'] in user_row:
                    user_date = self.normalize_date(user_row[user_info['date_column']])
                
                # Comparar - priorizar HC si existe, sino usar nombre
                match_found = False
                
                # Si ambos tienen HC y no están vacías
                if hospital_hc and user_hc and hospital_hc != "" and user_hc != "":
                    if hospital_hc == user_hc:
                        # HC coincide, verificar fecha
                        if hospital_date and user_date and hospital_date == user_date:
                            return True  # Encontrado con fecha coincidente
                        elif not hospital_date or not user_date:
                            return True  # Al menos uno no tiene fecha, considerar encontrado
                        # Si tienen fechas diferentes, continuar buscando
                
                # Si no hay HC o no coincide, comparar por nombre
                elif hospital_patient and user_patient and hospital_patient != "" and user_patient != "":
                    if hospital_patient == user_patient:
                        # Nombre coincide, verificar fecha
                        if hospital_date and user_date and hospital_date == user_date:
                            return True  # Encontrado con fecha coincidente
                        elif not hospital_date or not user_date:
                            return True  # Al menos uno no tiene fecha, considerar encontrado
                        # Si tienen fechas diferentes, continuar buscando
        
        return False  # No encontrado
    
    def process_files(self):
        if not self.hospital_files:
            messagebox.showerror("Error", "Debe seleccionar al menos un archivo del hospital")
            return
        
        if not self.user_files:
            messagebox.showerror("Error", "Debe seleccionar al menos un archivo del usuario")
            return
        
        try:
            self.status_label.config(text="Cargando archivos...")
            self.progress['value'] = 0
            self.root.update()
            
            # Cargar archivos del hospital
            hospital_data_list = []
            for i, filepath in enumerate(self.hospital_files):
                self.status_label.config(text=f"Cargando archivo del hospital {i+1}/{len(self.hospital_files)}")
                self.root.update()
                
                data, error = self.load_excel_file(filepath)
                if error:
                    messagebox.showerror("Error", error)
                    return
                hospital_data_list.append(data)
            
            # Cargar archivos del usuario
            user_data_list = []
            for i, filepath in enumerate(self.user_files):
                self.status_label.config(text=f"Cargando archivo del usuario {i+1}/{len(self.user_files)}")
                self.root.update()
                
                data, error = self.load_excel_file(filepath)
                if error:
                    messagebox.showerror("Error", error)
                    return
                user_data_list.append(data)
            
            # Procesar comparaciones
            missing_patients = []
            total_rows = sum(len(data['dataframe']) for data in hospital_data_list)
            processed_rows = 0
            
            self.status_label.config(text="Procesando comparaciones...")
            
            for hospital_info in hospital_data_list:
                hospital_df = hospital_info['dataframe']
                
                for _, hospital_row in hospital_df.iterrows():
                    processed_rows += 1
                    progress_percent = (processed_rows / total_rows) * 100
                    self.progress['value'] = progress_percent
                    
                    if processed_rows % 10 == 0:  # Actualizar cada 10 filas
                        self.status_label.config(text=f"Procesando: {processed_rows}/{total_rows}")
                        self.root.update()
                    
                    # Buscar paciente en archivos de usuario
                    found = self.search_patient_in_user_files(hospital_row, hospital_info, user_data_list)
                    
                    if not found:
                        # Crear registro del paciente faltante
                        missing_record = {}
                        for col in hospital_df.columns:
                            missing_record[col] = hospital_row[col]
                        missing_record['Archivo_Origen'] = hospital_info['filename']
                        missing_patients.append(missing_record)
            
            # Generar reporte
            if missing_patients:
                self.generate_report(missing_patients)
                messagebox.showinfo("Proceso Completado", 
                                  f"Proceso completado. Se encontraron {len(missing_patients)} pacientes no registrados.")
            else:
                messagebox.showinfo("Proceso Completado", 
                                  "Proceso completado. Todos los pacientes del hospital están registrados en los archivos del usuario.")
            
            self.progress['value'] = 100
            self.status_label.config(text="Proceso completado")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error durante el procesamiento: {str(e)}")
            self.status_label.config(text="Error en el procesamiento")
    
    def generate_report(self, missing_patients):
        #Genera el reporte de pacientes faltantes
        if not missing_patients:
            return
        
        # Crear DataFrame con los pacientes faltantes
        df_report = pd.DataFrame(missing_patients)
        
        # Reordenar columnas según el formato del archivo de usuario
        desired_columns = ['Hc', 'Paciente', 'Cobertura', 'Consultorio', 'Estado', 'Fecha']
        
        # Crear un nuevo DataFrame con las columnas deseadas
        report_data = []
        for patient in missing_patients:
            # Mapear columnas del hospital a formato del usuario
            row_data = {}
            
            # Buscar HC
            hc_value = ""
            for key, value in patient.items():
                if any(hc_pattern in str(key).lower() for hc_pattern in ['hc', 'historia', 'hist']):
                    hc_value = value if not pd.isna(value) else ""
                    break
            row_data['Hc'] = hc_value
            
            # Buscar Paciente/Nombre
            patient_value = ""
            for key, value in patient.items():
                if any(patient_pattern in str(key).lower() for patient_pattern in ['paciente', 'nombre', 'apellido']):
                    patient_value = value if not pd.isna(value) else ""
                    break
            row_data['Paciente'] = patient_value
            
            # Buscar Cobertura (obra social, plan, etc.)
            cobertura_value = ""
            for key, value in patient.items():
                if any(cob_pattern in str(key).lower() for cob_pattern in ['cobertura', 'obra', 'social', 'plan', 'seguro']):
                    cobertura_value = value if not pd.isna(value) else ""
                    break
            row_data['Cobertura'] = cobertura_value
            
            # Buscar Consultorio
            consultorio_value = ""
            for key, value in patient.items():
                if any(cons_pattern in str(key).lower() for cons_pattern in ['consultorio', 'consulta', 'atencion', 'servicio']):
                    consultorio_value = value if not pd.isna(value) else ""
                    break
            row_data['Consultorio'] = consultorio_value
            
            # Buscar Estado
            estado_value = ""
            for key, value in patient.items():
                if any(est_pattern in str(key).lower() for est_pattern in ['estado', 'status', 'situacion']):
                    estado_value = value if not pd.isna(value) else ""
                    break
            row_data['Estado'] = estado_value
            
            # Buscar Fecha
            fecha_value = ""
            for key, value in patient.items():
                if any(fecha_pattern in str(key).lower() for fecha_pattern in ['fecha', 'date', 'dia', 'day']):
                    if not pd.isna(value):
                        # Formatear fecha si es posible
                        try:
                            if isinstance(value, datetime):
                                fecha_value = value.strftime('%d/%m/%Y')
                            else:
                                fecha_value = str(value)
                        except:
                            fecha_value = str(value)
                    break
            row_data['Fecha'] = fecha_value
            
            report_data.append(row_data)
        
        # Crear DataFrame con formato estandarizado
        df_report = pd.DataFrame(report_data, columns=desired_columns)
        
        # Guardar archivo
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"reporte_pacientes_faltantes_{timestamp}.xlsx"
        
        try:
            df_report.to_excel(filename, index=False)
            self.status_label.config(text=f"Reporte guardado como: {filename}")
            
            # Preguntar si quiere abrir el archivo
            response = messagebox.askyesno("Archivo Generado", 
                                         f"El reporte se ha guardado como '{filename}'.\n\n¿Desea abrir el archivo ahora?")
            if response:
                self.open_file(filename)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el reporte: {str(e)}")
    
    def open_file(self, filename):
        #Abre el archivo generado con la aplicación por defecto del sistema
        try:
            import subprocess
            import sys
            
            if sys.platform == "win32":
                os.startfile(filename)
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", filename])
            else:  # Linux y otros
                subprocess.run(["xdg-open", filename])
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo: {str(e)}\n"
                                f"El archivo se encuentra en: {os.path.abspath(filename)}")

def main():
    root = tk.Tk()
    app = PatientControlApp(root)
    
    # Configurar estilo
    style = ttk.Style()
    style.theme_use('clam')
    
    root.mainloop()

if __name__ == "__main__":
    main()

    """

"""
#   V3

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import os
import re

class PatientControlApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Control de Pacientes - Hospital vs Usuario")
        self.root.geometry("800x600")
        
        # Variables para almacenar archivos
        self.hospital_files = []
        self.user_files = []
        
        self.setup_ui()
    
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        title_label = ttk.Label(main_frame, text="Sistema de Control de Pacientes", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)
        
        # Sección archivos del hospital
        hospital_frame = ttk.LabelFrame(main_frame, text="Archivos del Hospital", padding="10")
        hospital_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(hospital_frame, text="Agregar Archivos del Hospital", 
                  command=self.add_hospital_files).grid(row=0, column=0, padx=5)
        ttk.Button(hospital_frame, text="Limpiar Lista", 
                  command=self.clear_hospital_files).grid(row=0, column=1, padx=5)
        
        self.hospital_listbox = tk.Listbox(hospital_frame, height=6)
        self.hospital_listbox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        hospital_scrollbar = ttk.Scrollbar(hospital_frame, orient="vertical", 
                                         command=self.hospital_listbox.yview)
        hospital_scrollbar.grid(row=1, column=2, sticky=(tk.N, tk.S))
        self.hospital_listbox.configure(yscrollcommand=hospital_scrollbar.set)
        
        # Sección archivos del usuario
        user_frame = ttk.LabelFrame(main_frame, text="Archivos del Usuario", padding="10")
        user_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(user_frame, text="Agregar Archivos del Usuario", 
                  command=self.add_user_files).grid(row=0, column=0, padx=5)
        ttk.Button(user_frame, text="Limpiar Lista", 
                  command=self.clear_user_files).grid(row=0, column=1, padx=5)
        
        self.user_listbox = tk.Listbox(user_frame, height=6)
        self.user_listbox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        user_scrollbar = ttk.Scrollbar(user_frame, orient="vertical", 
                                     command=self.user_listbox.yview)
        user_scrollbar.grid(row=1, column=2, sticky=(tk.N, tk.S))
        self.user_listbox.configure(yscrollcommand=user_scrollbar.set)
        
        # Botones de acción
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=3, column=0, columnspan=3, pady=20)
        
        ttk.Button(action_frame, text="Procesar Archivos", 
                  command=self.process_files, style="Accent.TButton").grid(row=0, column=0, padx=10)
        ttk.Button(action_frame, text="Mostrar Información de Archivos", 
                  command=self.show_file_info).grid(row=0, column=1, padx=10)
        ttk.Button(action_frame, text="Salir", 
                  command=self.root.quit).grid(row=0, column=2, padx=10)
        
        # Barra de progreso
        self.progress = ttk.Progressbar(main_frame, length=300, mode='determinate')
        self.progress.grid(row=4, column=0, columnspan=3, pady=10)
        
        # Label de estado
        self.status_label = ttk.Label(main_frame, text="Listo para procesar archivos")
        self.status_label.grid(row=5, column=0, columnspan=3, pady=5)
        
        # Configurar pesos para redimensionamiento
        main_frame.columnconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def add_hospital_files(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos del hospital",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        for file in files:
            if file not in self.hospital_files:
                self.hospital_files.append(file)
                self.hospital_listbox.insert(tk.END, os.path.basename(file))
    
    def add_user_files(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos del usuario",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        for file in files:
            if file not in self.user_files:
                self.user_files.append(file)
                self.user_listbox.insert(tk.END, os.path.basename(file))
    
    def clear_hospital_files(self):
        self.hospital_files.clear()
        self.hospital_listbox.delete(0, tk.END)
    
    def clear_user_files(self):
        self.user_files.clear()
        self.user_listbox.delete(0, tk.END)
    
    def normalize_text(self, text):
        #Normaliza texto para comparación insensible a mayúsculas/minúsculas
        if pd.isna(text) or text is None:
            return ""
        return str(text).strip().upper()
    
    def normalize_date(self, date_value):
        #Normaliza fechas a formato comparable
        if pd.isna(date_value) or date_value is None:
            return None
        
        # Si ya es datetime
        if isinstance(date_value, datetime):
            return date_value.date()
        
        # Si es string, intentar parsearlo
        if isinstance(date_value, str):
            date_value = date_value.strip()
            # Intentar varios formatos comunes
            formats = ['%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d', '%d/%m/%y', '%d-%m-%y']
            for fmt in formats:
                try:
                    return datetime.strptime(date_value, fmt).date()
                except ValueError:
                    continue
        
        return None
    
    def find_hc_column(self, df):
        #Encuentra la columna de historia clínica
        hc_patterns = ['hc', 'historia', 'hist', 'h.c', 'historia clinica', 'historia clínica']
        
        for col in df.columns:
            col_name = str(col).lower().strip()
            for pattern in hc_patterns:
                if pattern in col_name:
                    return col
        return None
    
    def find_patient_column(self, df):
        #Encuentra la columna de paciente/nombre
        patient_patterns = ['paciente', 'nombre', 'apellido', 'patient', 'name']
        
        for col in df.columns:
            col_name = str(col).lower().strip()
            for pattern in patient_patterns:
                if pattern in col_name:
                    return col
        return None
    
    def find_date_column(self, df):
        #Encuentra la columna de fecha"
        date_patterns = ['fecha', 'date', 'dia', 'day']
        
        for col in df.columns:
            col_name = str(col).lower().strip()
            for pattern in date_patterns:
                if pattern in col_name:
                    return col
        return None
    
    def load_excel_file(self, filepath):
        #Carga un archivo Excel y retorna DataFrame con columnas identificadas
        try:
            # Intentar leer el archivo con diferentes configuraciones
            df = None
            
            # Intentar lectura estándar
            try:
                df = pd.read_excel(filepath)
            except:
                pass
            
            # Si falla, intentar saltando filas iniciales
            if df is None or df.empty:
                for skip_rows in range(1, 5):
                    try:
                        df = pd.read_excel(filepath, skiprows=skip_rows)
                        if not df.empty and len(df.columns) > 2:
                            break
                    except:
                        continue
            
            if df is None or df.empty:
                return None, "No se pudo leer el archivo o está vacío"
            
            # Limpiar DataFrame
            # Eliminar filas completamente vacías
            df = df.dropna(how='all')
            
            # Limpiar nombres de columnas
            df.columns = [str(col).strip() if col is not None else f'Col_{i}' 
                         for i, col in enumerate(df.columns)]
            
            # Identificar columnas importantes
            hc_col = self.find_hc_column(df)
            patient_col = self.find_patient_column(df)
            date_col = self.find_date_column(df)
            
            return {
                'dataframe': df,
                'hc_column': hc_col,
                'patient_column': patient_col,
                'date_column': date_col,
                'filename': os.path.basename(filepath)
            }, None
            
        except Exception as e:
            return None, f"Error al leer {filepath}: {str(e)}"
    
    def search_patient_in_user_files(self, hospital_row, hospital_info, user_data_list):
        #Busca un paciente del hospital en los archivos del usuario
        
        hospital_hc = None
        hospital_patient = None
        hospital_date = None
        
        # Obtener datos del hospital
        if hospital_info['hc_column'] and hospital_info['hc_column'] in hospital_row:
            hospital_hc = self.normalize_text(hospital_row[hospital_info['hc_column']])
        
        if hospital_info['patient_column'] and hospital_info['patient_column'] in hospital_row:
            hospital_patient = self.normalize_text(hospital_row[hospital_info['patient_column']])
        
        if hospital_info['date_column'] and hospital_info['date_column'] in hospital_row:
            hospital_date = self.normalize_date(hospital_row[hospital_info['date_column']])
        
        # Skip si no hay datos suficientes para comparar
        if (not hospital_hc or hospital_hc == "") and (not hospital_patient or hospital_patient == ""):
            return True  # No se puede verificar, asumir que existe
        
        # Buscar en archivos de usuario
        for user_info in user_data_list:
            user_df = user_info['dataframe']
            
            for _, user_row in user_df.iterrows():
                user_hc = None
                user_patient = None
                user_date = None
                
                # Obtener datos del usuario
                if user_info['hc_column'] and user_info['hc_column'] in user_row:
                    user_hc = self.normalize_text(user_row[user_info['hc_column']])
                
                if user_info['patient_column'] and user_info['patient_column'] in user_row:
                    user_patient = self.normalize_text(user_row[user_info['patient_column']])
                
                if user_info['date_column'] and user_info['date_column'] in user_row:
                    user_date = self.normalize_date(user_row[user_info['date_column']])
                
                # Skip filas vacías del usuario
                if (not user_hc or user_hc == "") and (not user_patient or user_patient == ""):
                    continue
                
                # Comparar - priorizar HC si existe
                match_found = False
                
                # Caso 1: Ambos tienen HC válida
                if (hospital_hc and hospital_hc != "" and 
                    user_hc and user_hc != ""):
                    if hospital_hc == user_hc:
                        # HC coincide, verificar fecha si ambas existen
                        if hospital_date and user_date:
                            if hospital_date == user_date:
                                return True  # Encontrado con fecha coincidente
                            # Si fechas no coinciden, continuar buscando
                        else:
                            return True  # HC coincide y al menos una no tiene fecha
                
                # Caso 2: Comparar por nombre si no hay HC o no coincide HC
                elif (hospital_patient and hospital_patient != "" and 
                      user_patient and user_patient != ""):
                    if hospital_patient == user_patient:
                        # Nombre coincide, verificar fecha si ambas existen
                        if hospital_date and user_date:
                            if hospital_date == user_date:
                                return True  # Encontrado con fecha coincidente
                            # Si fechas no coinciden, continuar buscando
                        else:
                            return True  # Nombre coincide y al menos una no tiene fecha
        
        return False  # No encontrado
    
    def process_files(self):
        if not self.hospital_files:
            messagebox.showerror("Error", "Debe seleccionar al menos un archivo del hospital")
            return
        
        if not self.user_files:
            messagebox.showerror("Error", "Debe seleccionar al menos un archivo del usuario")
            return
        
        try:
            self.status_label.config(text="Cargando archivos...")
            self.progress['value'] = 0
            self.root.update()
            
            # Cargar archivos del hospital
            hospital_data_list = []
            for i, filepath in enumerate(self.hospital_files):
                self.status_label.config(text=f"Cargando archivo del hospital {i+1}/{len(self.hospital_files)}")
                self.root.update()
                
                data, error = self.load_excel_file(filepath)
                if error:
                    messagebox.showerror("Error", error)
                    return
                hospital_data_list.append(data)
            
            # Cargar archivos del usuario
            user_data_list = []
            for i, filepath in enumerate(self.user_files):
                self.status_label.config(text=f"Cargando archivo del usuario {i+1}/{len(self.user_files)}")
                self.root.update()
                
                data, error = self.load_excel_file(filepath)
                if error:
                    messagebox.showerror("Error", error)
                    return
                user_data_list.append(data)
            
            # Procesar comparaciones
            missing_patients = []
            total_rows = sum(len(data['dataframe']) for data in hospital_data_list)
            processed_rows = 0
            
            self.status_label.config(text="Procesando comparaciones...")
            
            for hospital_info in hospital_data_list:
                hospital_df = hospital_info['dataframe']
                
                for _, hospital_row in hospital_df.iterrows():
                    processed_rows += 1
                    progress_percent = (processed_rows / total_rows) * 100
                    self.progress['value'] = progress_percent
                    
                    if processed_rows % 10 == 0:  # Actualizar cada 10 filas
                        self.status_label.config(text=f"Procesando: {processed_rows}/{total_rows}")
                        self.root.update()
                    
                    # Verificar que la fila tenga datos válidos
                    has_valid_data = False
                    for col in hospital_df.columns:
                        if not pd.isna(hospital_row[col]) and str(hospital_row[col]).strip() != "":
                            has_valid_data = True
                            break
                    
                    if not has_valid_data:
                        continue  # Skip filas completamente vacías
                    
                    # Buscar paciente en archivos de usuario
                    found = self.search_patient_in_user_files(hospital_row, hospital_info, user_data_list)
                    
                    if not found:
                        # Crear registro del paciente faltante
                        missing_record = {}
                        for col in hospital_df.columns:
                            missing_record[col] = hospital_row[col]
                        missing_record['Archivo_Origen'] = hospital_info['filename']
                        missing_patients.append(missing_record)
            
            # Generar reporte
            if missing_patients:
                self.generate_report(missing_patients)
                messagebox.showinfo("Proceso Completado", 
                                  f"Proceso completado. Se encontraron {len(missing_patients)} pacientes no registrados.")
            else:
                messagebox.showinfo("Proceso Completado", 
                                  "Proceso completado. Todos los pacientes del hospital están registrados en los archivos del usuario.")
            
            self.progress['value'] = 100
            self.status_label.config(text="Proceso completado")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error durante el procesamiento: {str(e)}")
            self.status_label.config(text="Error en el procesamiento")
    
    def generate_report(self, missing_patients):
        #Genera el reporte de pacientes faltantes
        if not missing_patients:
            return
        
        # Crear DataFrame con los pacientes faltantes
        df_report = pd.DataFrame(missing_patients)
        
        # Reordenar columnas según el formato del archivo de usuario
        desired_columns = ['Hc', 'Paciente', 'Cobertura', 'Consultorio', 'Estado', 'Fecha']
        
        # Crear un nuevo DataFrame con las columnas deseadas
        report_data = []
        for patient in missing_patients:
            # Mapear columnas del hospital a formato del usuario
            row_data = {}
            
            # Buscar HC
            hc_value = ""
            for key, value in patient.items():
                if any(hc_pattern in str(key).lower() for hc_pattern in ['hc', 'historia', 'hist']):
                    hc_value = value if not pd.isna(value) else ""
                    break
            row_data['Hc'] = hc_value
            
            # Buscar Paciente/Nombre
            patient_value = ""
            for key, value in patient.items():
                if any(patient_pattern in str(key).lower() for patient_pattern in ['paciente', 'nombre', 'apellido']):
                    patient_value = value if not pd.isna(value) else ""
                    break
            row_data['Paciente'] = patient_value
            
            # Buscar Cobertura (obra social, plan, etc.)
            cobertura_value = ""
            for key, value in patient.items():
                if any(cob_pattern in str(key).lower() for cob_pattern in ['cobertura', 'obra', 'social', 'plan', 'seguro']):
                    cobertura_value = value if not pd.isna(value) else ""
                    break
            row_data['Cobertura'] = cobertura_value
            
            # Buscar Consultorio
            consultorio_value = ""
            for key, value in patient.items():
                if any(cons_pattern in str(key).lower() for cons_pattern in ['consultorio', 'consulta', 'atencion', 'servicio']):
                    consultorio_value = value if not pd.isna(value) else ""
                    break
            row_data['Consultorio'] = consultorio_value
            
            # Buscar Estado
            estado_value = ""
            for key, value in patient.items():
                if any(est_pattern in str(key).lower() for est_pattern in ['estado', 'status', 'situacion']):
                    estado_value = value if not pd.isna(value) else ""
                    break
            row_data['Estado'] = estado_value
            
            # Buscar Fecha
            fecha_value = ""
            for key, value in patient.items():
                if any(fecha_pattern in str(key).lower() for fecha_pattern in ['fecha', 'date', 'dia', 'day']):
                    if not pd.isna(value):
                        # Formatear fecha si es posible
                        try:
                            if isinstance(value, datetime):
                                fecha_value = value.strftime('%d/%m/%Y')
                            else:
                                fecha_value = str(value)
                        except:
                            fecha_value = str(value)
                    break
            row_data['Fecha'] = fecha_value
            
            report_data.append(row_data)
        
        # Crear DataFrame con formato estandarizado
        df_report = pd.DataFrame(report_data, columns=desired_columns)
        
        # Guardar archivo
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"reporte_pacientes_faltantes_{timestamp}.xlsx"
        
        try:
            df_report.to_excel(filename, index=False)
            self.status_label.config(text=f"Reporte guardado como: {filename}")
            
            # Preguntar si quiere abrir el archivo
            response = messagebox.askyesno("Archivo Generado", 
                                         f"El reporte se ha guardado como '{filename}'.\n\n¿Desea abrir el archivo ahora?")
            if response:
                self.open_file(filename)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el reporte: {str(e)}")
    
    def open_file(self, filename):
        #Abre el archivo generado con la aplicación por defecto del sistema
        try:
            import subprocess
            import sys
            
            if sys.platform == "win32":
                os.startfile(filename)
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", filename])
            else:  # Linux y otros
                subprocess.run(["xdg-open", filename])
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo: {str(e)}\n"
                                f"El archivo se encuentra en: {os.path.abspath(filename)}")
    
    def show_file_info(self):
        #Muestra información detallada sobre los archivos cargados
        if not self.hospital_files and not self.user_files:
            messagebox.showinfo("Información", "No hay archivos cargados")
            return
        
        info_text = "INFORMACIÓN DE ARCHIVOS CARGADOS\n" + "="*50 + "\n\n"
        
        # Información de archivos del hospital
        if self.hospital_files:
            info_text += "ARCHIVOS DEL HOSPITAL:\n" + "-"*25 + "\n"
            for i, filepath in enumerate(self.hospital_files):
                try:
                    data, error = self.load_excel_file(filepath)
                    if error:
                        info_text += f"{i+1}. {os.path.basename(filepath)} - ERROR: {error}\n"
                    else:
                        df = data['dataframe']
                        info_text += f"{i+1}. {os.path.basename(filepath)}\n"
                        info_text += f"   - Filas: {len(df)}\n"
                        info_text += f"   - Columnas: {list(df.columns)}\n"
                        info_text += f"   - HC detectada: {data['hc_column']}\n"
                        info_text += f"   - Paciente detectado: {data['patient_column']}\n"
                        info_text += f"   - Fecha detectada: {data['date_column']}\n\n"
                except Exception as e:
                    info_text += f"{i+1}. {os.path.basename(filepath)} - ERROR: {str(e)}\n\n"
        
        # Información de archivos del usuario
        if self.user_files:
            info_text += "ARCHIVOS DEL USUARIO:\n" + "-"*25 + "\n"
            for i, filepath in enumerate(self.user_files):
                try:
                    data, error = self.load_excel_file(filepath)
                    if error:
                        info_text += f"{i+1}. {os.path.basename(filepath)} - ERROR: {error}\n"
                    else:
                        df = data['dataframe']
                        info_text += f"{i+1}. {os.path.basename(filepath)}\n"
                        info_text += f"   - Filas: {len(df)}\n"
                        info_text += f"   - Columnas: {list(df.columns)}\n"
                        info_text += f"   - HC detectada: {data['hc_column']}\n"
                        info_text += f"   - Paciente detectado: {data['patient_column']}\n"
                        info_text += f"   - Fecha detectada: {data['date_column']}\n\n"
                except Exception as e:
                    info_text += f"{i+1}. {os.path.basename(filepath)} - ERROR: {str(e)}\n\n"
        
        # Mostrar en ventana popup
        info_window = tk.Toplevel(self.root)
        info_window.title("Información de Archivos")
        info_window.geometry("800x600")
        
        text_widget = tk.Text(info_window, wrap=tk.WORD, font=("Consolas", 10))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        scrollbar = ttk.Scrollbar(info_window, orient="vertical", command=text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.insert(tk.END, info_text)
        text_widget.configure(state=tk.DISABLED)

def main():
    root = tk.Tk()
    app = PatientControlApp(root)
    
    # Configurar estilo
    style = ttk.Style()
    style.theme_use('clam')
    
    root.mainloop()

if __name__ == "__main__":
    main()

"""

    #   V4

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import os
import re

class PatientControlApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Control de Pacientes - Hospital vs Usuario")
        self.root.geometry("800x600")
        
        # Variables para almacenar archivos
        self.hospital_files = []
        self.user_files = []
        
        self.setup_ui()
    
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        title_label = ttk.Label(main_frame, text="Sistema de Control de Pacientes", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)
        
        # Sección archivos del hospital
        hospital_frame = ttk.LabelFrame(main_frame, text="Archivos del Hospital", padding="10")
        hospital_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(hospital_frame, text="Agregar Archivos del Hospital", 
                  command=self.add_hospital_files).grid(row=0, column=0, padx=5)
        ttk.Button(hospital_frame, text="Limpiar Lista", 
                  command=self.clear_hospital_files).grid(row=0, column=1, padx=5)
        
        self.hospital_listbox = tk.Listbox(hospital_frame, height=6)
        self.hospital_listbox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        hospital_scrollbar = ttk.Scrollbar(hospital_frame, orient="vertical", 
                                         command=self.hospital_listbox.yview)
        hospital_scrollbar.grid(row=1, column=2, sticky=(tk.N, tk.S))
        self.hospital_listbox.configure(yscrollcommand=hospital_scrollbar.set)
        
        # Sección archivos del usuario
        user_frame = ttk.LabelFrame(main_frame, text="Archivos del Usuario", padding="10")
        user_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(user_frame, text="Agregar Archivos del Usuario", 
                  command=self.add_user_files).grid(row=0, column=0, padx=5)
        ttk.Button(user_frame, text="Limpiar Lista", 
                  command=self.clear_user_files).grid(row=0, column=1, padx=5)
        
        self.user_listbox = tk.Listbox(user_frame, height=6)
        self.user_listbox.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        user_scrollbar = ttk.Scrollbar(user_frame, orient="vertical", 
                                     command=self.user_listbox.yview)
        user_scrollbar.grid(row=1, column=2, sticky=(tk.N, tk.S))
        self.user_listbox.configure(yscrollcommand=user_scrollbar.set)
        
        # Botones de acción
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=3, column=0, columnspan=3, pady=20)
        
        ttk.Button(action_frame, text="Procesar Archivos", 
                  command=self.process_files, style="Accent.TButton").grid(row=0, column=0, padx=10)
        ttk.Button(action_frame, text="Mostrar Información de Archivos", 
                  command=self.show_file_info).grid(row=0, column=1, padx=10)
        ttk.Button(action_frame, text="Salir", 
                  command=self.root.quit).grid(row=0, column=2, padx=10)
        
        # Barra de progreso
        self.progress = ttk.Progressbar(main_frame, length=300, mode='determinate')
        self.progress.grid(row=4, column=0, columnspan=3, pady=10)
        
        # Label de estado
        self.status_label = ttk.Label(main_frame, text="Listo para procesar archivos")
        self.status_label.grid(row=5, column=0, columnspan=3, pady=5)
        
        # Configurar pesos para redimensionamiento
        main_frame.columnconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def add_hospital_files(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos del hospital",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        for file in files:
            if file not in self.hospital_files:
                self.hospital_files.append(file)
                self.hospital_listbox.insert(tk.END, os.path.basename(file))
    
    def add_user_files(self):
        files = filedialog.askopenfilenames(
            title="Seleccionar archivos del usuario",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        for file in files:
            if file not in self.user_files:
                self.user_files.append(file)
                self.user_listbox.insert(tk.END, os.path.basename(file))
    
    def clear_hospital_files(self):
        self.hospital_files.clear()
        self.hospital_listbox.delete(0, tk.END)
    
    def clear_user_files(self):
        self.user_files.clear()
        self.user_listbox.delete(0, tk.END)
    
    def normalize_text(self, text):
        """Normaliza texto para comparación insensible a mayúsculas/minúsculas"""
        if pd.isna(text) or text is None:
            return ""
        return str(text).strip().upper()
    
    def normalize_date(self, date_value):
        """Normaliza fechas a formato comparable"""
        if pd.isna(date_value) or date_value is None:
            return None
        
        # Si ya es datetime
        if isinstance(date_value, datetime):
            return date_value.date()
        
        # Si es string, intentar parsearlo
        if isinstance(date_value, str):
            date_value = date_value.strip()
            # Intentar varios formatos comunes
            formats = ['%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d', '%d/%m/%y', '%d-%m-%y']
            for fmt in formats:
                try:
                    return datetime.strptime(date_value, fmt).date()
                except ValueError:
                    continue
        
        return None
    
    def find_hc_column(self, df):
        """Encuentra la columna de historia clínica"""
        hc_patterns = ['hc', 'historia', 'hist', 'h.c', 'historia clinica', 'historia clínica']
        
        for col in df.columns:
            col_name = str(col).lower().strip()
            for pattern in hc_patterns:
                if pattern in col_name:
                    return col
        return None
    
    def find_patient_column(self, df):
        """Encuentra la columna de paciente/nombre"""
        patient_patterns = ['paciente', 'nombre', 'apellido', 'patient', 'name']
        
        for col in df.columns:
            col_name = str(col).lower().strip()
            for pattern in patient_patterns:
                if pattern in col_name:
                    return col
        return None
    
    def find_date_column(self, df):
        """Encuentra la columna de fecha"""
        date_patterns = ['fecha', 'date', 'dia', 'day']
        
        for col in df.columns:
            col_name = str(col).lower().strip()
            for pattern in date_patterns:
                if pattern in col_name:
                    return col
        return None
    
    def load_excel_file(self, filepath):
        """Carga un archivo Excel y retorna DataFrame con columnas identificadas"""
        try:
            # Intentar leer el archivo con diferentes configuraciones
            df = None
            
            # Intentar lectura estándar
            try:
                df = pd.read_excel(filepath)
            except:
                pass
            
            # Si falla, intentar saltando filas iniciales
            if df is None or df.empty:
                for skip_rows in range(1, 5):
                    try:
                        df = pd.read_excel(filepath, skiprows=skip_rows)
                        if not df.empty and len(df.columns) > 2:
                            break
                    except:
                        continue
            
            if df is None or df.empty:
                return None, "No se pudo leer el archivo o está vacío"
            
            # Limpiar DataFrame
            # Eliminar filas completamente vacías
            df = df.dropna(how='all')
            
            # Limpiar nombres de columnas
            df.columns = [str(col).strip() if col is not None else f'Col_{i}' 
                         for i, col in enumerate(df.columns)]
            
            # Identificar columnas importantes
            hc_col = self.find_hc_column(df)
            patient_col = self.find_patient_column(df)
            date_col = self.find_date_column(df)
            
            return {
                'dataframe': df,
                'hc_column': hc_col,
                'patient_column': patient_col,
                'date_column': date_col,
                'filename': os.path.basename(filepath)
            }, None
            
        except Exception as e:
            return None, f"Error al leer {filepath}: {str(e)}"
    
    def search_user_patient_in_hospital_files(self, user_row, user_info, hospital_data_list):
        """Busca un paciente del USUARIO en los archivos del HOSPITAL"""
        
        user_hc = None
        user_patient = None
        user_date = None
        
        # Obtener datos del USUARIO
        if user_info['hc_column'] and user_info['hc_column'] in user_row:
            user_hc = self.normalize_text(user_row[user_info['hc_column']])
        
        if user_info['patient_column'] and user_info['patient_column'] in user_row:
            user_patient = self.normalize_text(user_row[user_info['patient_column']])
        
        if user_info['date_column'] and user_info['date_column'] in user_row:
            user_date = self.normalize_date(user_row[user_info['date_column']])
        
        # Skip si no hay datos suficientes para comparar
        if (not user_hc or user_hc == "") and (not user_patient or user_patient == ""):
            return True  # No se puede verificar, asumir que existe (pagado)
        
        # Buscar en archivos del HOSPITAL
        for hospital_info in hospital_data_list:
            hospital_df = hospital_info['dataframe']
            
            for _, hospital_row in hospital_df.iterrows():
                hospital_hc = None
                hospital_patient = None
                hospital_date = None
                
                # Obtener datos del HOSPITAL
                if hospital_info['hc_column'] and hospital_info['hc_column'] in hospital_row:
                    hospital_hc = self.normalize_text(hospital_row[hospital_info['hc_column']])
                
                if hospital_info['patient_column'] and hospital_info['patient_column'] in hospital_row:
                    hospital_patient = self.normalize_text(hospital_row[hospital_info['patient_column']])
                
                if hospital_info['date_column'] and hospital_info['date_column'] in hospital_row:
                    hospital_date = self.normalize_date(hospital_row[hospital_info['date_column']])
                
                # Skip filas vacías del hospital
                if (not hospital_hc or hospital_hc == "") and (not hospital_patient or hospital_patient == ""):
                    continue
                
                # Comparar - priorizar HC si existe
                match_found = False
                
                # Caso 1: Ambos tienen HC válida
                if (user_hc and user_hc != "" and 
                    hospital_hc and hospital_hc != ""):
                    if user_hc == hospital_hc:
                        # HC coincide, verificar fecha si ambas existen
                        if user_date and hospital_date:
                            if user_date == hospital_date:
                                return True  # Encontrado con fecha coincidente (PAGADO)
                            # Si fechas no coinciden, continuar buscando
                        else:
                            return True  # HC coincide y al menos una no tiene fecha (PAGADO)
                
                # Caso 2: Comparar por nombre si no hay HC o no coincide HC
                elif (user_patient and user_patient != "" and 
                      hospital_patient and hospital_patient != ""):
                    if user_patient == hospital_patient:
                        # Nombre coincide, verificar fecha si ambas existen
                        if user_date and hospital_date:
                            if user_date == hospital_date:
                                return True  # Encontrado con fecha coincidente (PAGADO)
                            # Si fechas no coinciden, continuar buscando
                        else:
                            return True  # Nombre coincide y al menos una no tiene fecha (PAGADO)
        
        return False  # No encontrado en hospital (NO PAGADO)
    
    def process_files(self):
        if not self.hospital_files:
            messagebox.showerror("Error", "Debe seleccionar al menos un archivo del hospital")
            return
        
        if not self.user_files:
            messagebox.showerror("Error", "Debe seleccionar al menos un archivo del usuario")
            return
        
        try:
            self.status_label.config(text="Cargando archivos...")
            self.progress['value'] = 0
            self.root.update()
            
            # Cargar archivos del hospital
            hospital_data_list = []
            for i, filepath in enumerate(self.hospital_files):
                self.status_label.config(text=f"Cargando archivo del hospital {i+1}/{len(self.hospital_files)}")
                self.root.update()
                
                data, error = self.load_excel_file(filepath)
                if error:
                    messagebox.showerror("Error", error)
                    return
                hospital_data_list.append(data)
            
            # Cargar archivos del usuario
            user_data_list = []
            for i, filepath in enumerate(self.user_files):
                self.status_label.config(text=f"Cargando archivo del usuario {i+1}/{len(self.user_files)}")
                self.root.update()
                
                data, error = self.load_excel_file(filepath)
                if error:
                    messagebox.showerror("Error", error)
                    return
                user_data_list.append(data)
            
            # Procesar comparaciones - LÓGICA CORREGIDA
            missing_patients = []
            total_rows = sum(len(data['dataframe']) for data in user_data_list)  # Total de filas del USUARIO
            processed_rows = 0
            
            self.status_label.config(text="Procesando comparaciones...")
            
            # LÓGICA CORREGIDA: Buscar pacientes del USUARIO en archivos del HOSPITAL
            for user_info in user_data_list:
                user_df = user_info['dataframe']
                
                for _, user_row in user_df.iterrows():
                    processed_rows += 1
                    progress_percent = (processed_rows / total_rows) * 100
                    self.progress['value'] = progress_percent
                    
                    if processed_rows % 10 == 0:  # Actualizar cada 10 filas
                        self.status_label.config(text=f"Procesando: {processed_rows}/{total_rows}")
                        self.root.update()
                    
                    # Verificar que la fila del USUARIO tenga datos válidos
                    has_valid_data = False
                    for col in user_df.columns:
                        if not pd.isna(user_row[col]) and str(user_row[col]).strip() != "":
                            has_valid_data = True
                            break
                    
                    if not has_valid_data:
                        continue  # Skip filas completamente vacías
                    
                    # Buscar este paciente del USUARIO en los archivos del HOSPITAL
                    found = self.search_user_patient_in_hospital_files(user_row, user_info, hospital_data_list)
                    
                    if not found:
                        # Crear registro del paciente del USUARIO que NO fue encontrado en hospital (no pagado)
                        missing_record = {}
                        for col in user_df.columns:
                            missing_record[col] = user_row[col]
                        missing_record['Archivo_Origen_Usuario'] = user_info['filename']
                        missing_patients.append(missing_record)
            
            # Generar reporte
            if missing_patients:
                self.generate_report(missing_patients)
                messagebox.showinfo("Proceso Completado", 
                                  f"Proceso completado. Se encontraron {len(missing_patients)} pacientes atendidos por el usuario que NO aparecen en los archivos del hospital (posiblemente no pagados).")
            else:
                messagebox.showinfo("Proceso Completado", 
                                  "Proceso completado. Todos los pacientes atendidos por el usuario aparecen en los archivos del hospital.")
            
            self.progress['value'] = 100
            self.status_label.config(text="Proceso completado")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error durante el procesamiento: {str(e)}")
            self.status_label.config(text="Error en el procesamiento")
    
    def generate_report(self, missing_patients):
        """Genera el reporte de pacientes del usuario que NO aparecen en hospital (no pagados)"""
        if not missing_patients:
            return
        
        # Los datos ya vienen del archivo del USUARIO con las columnas correctas
        # Solo necesitamos mapear a las columnas estándar si es necesario
        
        desired_columns = ['Hc', 'Paciente', 'Cobertura', 'Consultorio', 'Estado', 'Fecha']
        
        # Crear un nuevo DataFrame con las columnas deseadas
        report_data = []
        for patient in missing_patients:
            # Los datos ya vienen del usuario, solo necesitamos mapearlos si tienen nombres diferentes
            row_data = {}
            
            # Mapear directamente si las columnas ya existen con los nombres correctos
            for desired_col in desired_columns:
                row_data[desired_col] = ""
                
                # Buscar la columna en los datos originales del usuario
                for key, value in patient.items():
                    if key == 'Archivo_Origen_Usuario':
                        continue
                        
                    key_lower = str(key).lower().strip()
                    
                    if desired_col == 'Hc':
                        if any(hc_pattern in key_lower for hc_pattern in ['hc', 'historia', 'hist']):
                            row_data[desired_col] = value if not pd.isna(value) else ""
                            break
                    elif desired_col == 'Paciente':
                        if any(patient_pattern in key_lower for patient_pattern in ['paciente', 'nombre', 'apellido']):
                            row_data[desired_col] = value if not pd.isna(value) else ""
                            break
                    elif desired_col == 'Cobertura':
                        if any(cob_pattern in key_lower for cob_pattern in ['cobertura', 'obra', 'social', 'plan', 'seguro']):
                            row_data[desired_col] = value if not pd.isna(value) else ""
                            break
                    elif desired_col == 'Consultorio':
                        if any(cons_pattern in key_lower for cons_pattern in ['consultorio', 'consulta', 'atencion', 'servicio']):
                            row_data[desired_col] = value if not pd.isna(value) else ""
                            break
                    elif desired_col == 'Estado':
                        if any(est_pattern in key_lower for est_pattern in ['estado', 'status', 'situacion']):
                            row_data[desired_col] = value if not pd.isna(value) else ""
                            break
                    elif desired_col == 'Fecha':
                        if any(fecha_pattern in key_lower for fecha_pattern in ['fecha', 'date', 'dia', 'day']):
                            if not pd.isna(value):
                                try:
                                    if isinstance(value, datetime):
                                        row_data[desired_col] = value.strftime('%d/%m/%Y')
                                    else:
                                        row_data[desired_col] = str(value)
                                except:
                                    row_data[desired_col] = str(value)
                            break
            
            report_data.append(row_data)
        
        # Crear DataFrame con formato estandarizado
        df_report = pd.DataFrame(report_data, columns=desired_columns)
        
        # Guardar archivo
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"pacientes_usuario_no_pagados_{timestamp}.xlsx"
        
        try:
            df_report.to_excel(filename, index=False)
            self.status_label.config(text=f"Reporte guardado como: {filename}")
            
            # Preguntar si quiere abrir el archivo
            response = messagebox.askyesno("Archivo Generado", 
                                         f"El reporte se ha guardado como '{filename}'.\n\n¿Desea abrir el archivo ahora?")
            if response:
                self.open_file(filename)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al guardar el reporte: {str(e)}")
    
    def open_file(self, filename):
        """Abre el archivo generado con la aplicación por defecto del sistema"""
        try:
            import subprocess
            import sys
            
            if sys.platform == "win32":
                os.startfile(filename)
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", filename])
            else:  # Linux y otros
                subprocess.run(["xdg-open", filename])
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo: {str(e)}\n"
                                f"El archivo se encuentra en: {os.path.abspath(filename)}")
    
    def show_file_info(self):
        """Muestra información detallada sobre los archivos cargados"""
        if not self.hospital_files and not self.user_files:
            messagebox.showinfo("Información", "No hay archivos cargados")
            return
        
        info_text = "INFORMACIÓN DE ARCHIVOS CARGADOS\n" + "="*50 + "\n\n"
        
        # Información de archivos del hospital
        if self.hospital_files:
            info_text += "ARCHIVOS DEL HOSPITAL:\n" + "-"*25 + "\n"
            for i, filepath in enumerate(self.hospital_files):
                try:
                    data, error = self.load_excel_file(filepath)
                    if error:
                        info_text += f"{i+1}. {os.path.basename(filepath)} - ERROR: {error}\n"
                    else:
                        df = data['dataframe']
                        info_text += f"{i+1}. {os.path.basename(filepath)}\n"
                        info_text += f"   - Filas: {len(df)}\n"
                        info_text += f"   - Columnas: {list(df.columns)}\n"
                        info_text += f"   - HC detectada: {data['hc_column']}\n"
                        info_text += f"   - Paciente detectado: {data['patient_column']}\n"
                        info_text += f"   - Fecha detectada: {data['date_column']}\n\n"
                except Exception as e:
                    info_text += f"{i+1}. {os.path.basename(filepath)} - ERROR: {str(e)}\n\n"
        
        # Información de archivos del usuario
        if self.user_files:
            info_text += "ARCHIVOS DEL USUARIO:\n" + "-"*25 + "\n"
            for i, filepath in enumerate(self.user_files):
                try:
                    data, error = self.load_excel_file(filepath)
                    if error:
                        info_text += f"{i+1}. {os.path.basename(filepath)} - ERROR: {error}\n"
                    else:
                        df = data['dataframe']
                        info_text += f"{i+1}. {os.path.basename(filepath)}\n"
                        info_text += f"   - Filas: {len(df)}\n"
                        info_text += f"   - Columnas: {list(df.columns)}\n"
                        info_text += f"   - HC detectada: {data['hc_column']}\n"
                        info_text += f"   - Paciente detectado: {data['patient_column']}\n"
                        info_text += f"   - Fecha detectada: {data['date_column']}\n\n"
                except Exception as e:
                    info_text += f"{i+1}. {os.path.basename(filepath)} - ERROR: {str(e)}\n\n"
        
        # Mostrar en ventana popup
        info_window = tk.Toplevel(self.root)
        info_window.title("Información de Archivos")
        info_window.geometry("800x600")
        
        text_widget = tk.Text(info_window, wrap=tk.WORD, font=("Consolas", 10))
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        scrollbar = ttk.Scrollbar(info_window, orient="vertical", command=text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.insert(tk.END, info_text)
        text_widget.configure(state=tk.DISABLED)

def main():
    root = tk.Tk()
    app = PatientControlApp(root)
    
    # Configurar estilo
    style = ttk.Style()
    style.theme_use('clam')
    
    root.mainloop()

if __name__ == "__main__":
    main()
 