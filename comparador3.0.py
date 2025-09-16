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
        """Busca un paciente del hospital en los archivos del usuario"""
        
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
        """Genera el reporte de pacientes faltantes"""
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