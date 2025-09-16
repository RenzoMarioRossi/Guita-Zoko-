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

def detect_file_type(filename):
    """Detecta el tipo de archivo basado en el nombre."""
    filename_lower = filename.lower()
    
    if 'planes' in filename_lower:
        return 'planes'
    elif 'pami' in filename_lower:
        return 'pami'
    elif 'ooss' in filename_lower:
        return 'ooss'
    else:
        return 'usuario'

def get_column_mapping(file_type):
    """Retorna el mapeo de columnas según el tipo de archivo."""
    
    base_mapping = {
        'nombre': 'Nombre',
        'fecha': 'Fecha',
        'paciente': 'Nombre'  # Agregado para archivos de usuario
    }
    
    if file_type == 'planes':
        return {
            **base_mapping,
            'hc': 'HC',
            'historia': 'HC',
            'historia_clinica': 'HC',
            'hono_impu1': 'Monto',
            'honorarios': 'Monto',
            'monto': 'Monto',
            'cobertura': 'Cobertura',
            'plan': 'Cobertura'
        }
    
    elif file_type == 'pami':
        return {
            **base_mapping,
            'hc': 'HC',
            'historia': 'HC',
            'historia_clinica': 'HC',
            'hono_impu1': 'Monto',
            'honorarios': 'Monto',
            'monto': 'Monto',
            'desgrupo': 'Desgrupo',
            'desc_cob': 'Desc_Cob'
        }
    
    elif file_type == 'ooss':
        return {
            **base_mapping,
            'historia': 'HC',  # En OOSS la HC se llama "historia"
            'hc': 'HC',
            'historia_clinica': 'HC',
            'hono_impu1': 'Monto',
            'honorarios': 'Monto',
            'monto': 'Monto',
            'desc_cob': 'Desc_Cob',
            'obra_social': 'Obra_Social'
        }
    
    else:  # archivo usuario
        return {
            **base_mapping,
            'hc': 'HC',
            'historia': 'HC',
            'historia_clinica': 'HC',
            'monto': 'Monto',
            'honorarios': 'Monto',
            'hono_impu1': 'Monto',
            'hora': 'Hora',
            'plan': 'Plan',
            'obra_social': 'Obra_Social'
        }

def process_dataframe(df, file_path):
    """Procesa un DataFrame según el tipo de archivo."""
    if df is None or df.empty:
        return pd.DataFrame()
    
    filename = os.path.basename(file_path)
    file_type = detect_file_type(filename)
    
    # Hacer una copia para evitar modificar el original
    df = df.copy()
    
    # Limpiar nombres de columnas
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')
    
    # Obtener mapeo para este tipo de archivo
    column_mapping = get_column_mapping(file_type)
    
    # Renombrar columnas
    df = df.rename(columns=column_mapping)
    
    # Verificar columnas esenciales
    required_cols = ['Nombre', 'HC', 'Fecha']
    
    # Crear columnas faltantes con valores nulos
    for col in required_cols:
        if col not in df.columns:
            df[col] = np.nan
    
    # Procesar HC - convertir a string y limpiar
    if 'HC' in df.columns:
        df['HC'] = df['HC'].astype(str).str.strip()
        # Filtrar registros con HC válida (no vacía, no 'nan')
        df = df[df['HC'].notna() & (df['HC'] != '') & (df['HC'] != 'nan')]
    
    # Procesar Fecha
    if 'Fecha' in df.columns:
        df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce', dayfirst=True).dt.date
        # Filtrar registros con fecha válida
        df = df[df['Fecha'].notna()]
    
    # Procesar Monto si existe y tiene datos
    if 'Monto' in df.columns:
        df['Monto'] = df['Monto'].apply(clean_monto)
    else:
        df['Monto'] = 0.0
    
    # Agregar columnas de identificación SIEMPRE
    df['Archivo_Origen'] = filename
    df['Tipo_Archivo'] = file_type
    
    # Seleccionar columnas relevantes según el tipo
    base_cols = ['HC', 'Nombre', 'Fecha', 'Monto', 'Archivo_Origen', 'Tipo_Archivo']
    
    # Agregar columnas específicas según el tipo
    specific_cols = []
    if file_type == 'planes':
        if 'Cobertura' in df.columns:
            specific_cols.append('Cobertura')
    elif file_type == 'pami':
        if 'Desgrupo' in df.columns:
            specific_cols.append('Desgrupo')
        if 'Desc_Cob' in df.columns:
            specific_cols.append('Desc_Cob')
    elif file_type == 'ooss':
        if 'Desc_Cob' in df.columns:
            specific_cols.append('Desc_Cob')
        if 'Obra_Social' in df.columns:
            specific_cols.append('Obra_Social')
    elif file_type == 'usuario':
        if 'Hora' in df.columns:
            specific_cols.append('Hora')
        if 'Plan' in df.columns:
            specific_cols.append('Plan')
        if 'Obra_Social' in df.columns:
            specific_cols.append('Obra_Social')
    
    final_cols = base_cols + specific_cols
    
    # Filtrar solo las columnas que existen
    existing_cols = [col for col in final_cols if col in df.columns]
    
    # Retornar DataFrame procesado
    result_df = df[existing_cols].copy()
    
    # Eliminar filas completamente vacías
    result_df = result_df.dropna(how='all')
    
    return result_df

def compare_records(user_file, hospital_files, output_dir=None):
    """Compara registro con los del hospital y genera un Excel con discrepancias."""
    try:
        print(f"Procesando archivo de usuario: {user_file}")
        
        # Procesar archivo del usuario
        user_df = pd.read_excel(user_file)
        print(f"Archivo de usuario cargado. Filas: {len(user_df)}")
        
        user_df = process_dataframe(user_df, user_file)
        print(f"Archivo de usuario procesado. Filas: {len(user_df)}")
        
        if len(user_df) == 0:
            raise Exception("El archivo de usuario no contiene registros válidos")
        
        # Procesar archivos del hospital
        hospital_dfs = []
        
        for file_path in hospital_files:
            try:
                print(f"Procesando archivo de hospital: {os.path.basename(file_path)}")
                df = pd.read_excel(file_path)
                print(f"  Cargado. Filas: {len(df)}")
                
                processed_df = process_dataframe(df, file_path)
                print(f"  Procesado. Filas: {len(processed_df)}")
                
                if len(processed_df) > 0:
                    hospital_dfs.append(processed_df)
                    
            except Exception as e:
                print(f"  Error procesando {file_path}: {str(e)}")
                continue
        
        if not hospital_dfs:
            raise Exception("No se pudieron procesar archivos del hospital")
        
        # Combinar archivos del hospital
        hospital_df = pd.concat(hospital_dfs, ignore_index=True)
        print(f"Archivos de hospital combinados. Total filas: {len(hospital_df)}")
        
        # Verificar que ambos DataFrames tengan las columnas necesarias
        required_user_cols = ['HC', 'Fecha']
        required_hospital_cols = ['HC', 'Fecha', 'Archivo_Origen', 'Tipo_Archivo']
        
        missing_user_cols = [col for col in required_user_cols if col not in user_df.columns]
        missing_hospital_cols = [col for col in required_hospital_cols if col not in hospital_df.columns]
        
        if missing_user_cols:
            raise Exception(f"Faltan columnas en archivo usuario: {missing_user_cols}")
        if missing_hospital_cols:
            raise Exception(f"Faltan columnas en archivos hospital: {missing_hospital_cols}")
        
        # Realizar merge para encontrar discrepancias (HC + Fecha como clave compuesta)
        key_cols = ['HC', 'Fecha']
        
        # Preparar columnas para el merge
        user_merge_cols = ['HC', 'Fecha', 'Nombre', 'Monto']
        hospital_merge_cols = ['HC', 'Fecha', 'Nombre', 'Monto', 'Archivo_Origen', 'Tipo_Archivo']
        
        # Filtrar solo las columnas que existen
        user_merge_cols = [col for col in user_merge_cols if col in user_df.columns]
        hospital_merge_cols = [col for col in hospital_merge_cols if col in hospital_df.columns]
        
        print(f"Realizando merge con columnas usuario: {user_merge_cols}")
        print(f"Realizando merge con columnas hospital: {hospital_merge_cols}")
        
        merged = user_df[user_merge_cols].merge(
            hospital_df[hospital_merge_cols],
            on=key_cols, 
            how='outer', 
            suffixes=('_usuario', '_hospital'), 
            indicator=True
        )
        
        print(f"Merge completado. Total registros: {len(merged)}")
        
        # Extra en registro del usuario (a favor)
        extra_user_mask = merged['_merge'] == 'left_only'
        extra_user = merged[extra_user_mask].copy()
        
        print(f"Registros extra en usuario: {len(extra_user)}")
        
        if not extra_user.empty:
            # Limpiar y organizar columnas
            extra_user['Nombre'] = extra_user.get('Nombre_usuario', extra_user.get('Nombre_hospital', ''))
            extra_user['Monto'] = extra_user.get('Monto_usuario', extra_user.get('Monto_hospital', 0.0))
            
            # Agregar información adicional del archivo usuario
            user_additional = user_df.copy()
            if not user_additional.empty:
                extra_user = extra_user.merge(
                    user_additional[['HC', 'Fecha'] + [col for col in user_additional.columns 
                                                     if col not in ['HC', 'Fecha', 'Nombre', 'Monto']]],
                    on=['HC', 'Fecha'], 
                    how='left'
                )
            
            # Seleccionar columnas relevantes
            cols_to_keep = ['HC', 'Fecha', 'Nombre', 'Monto']
            for col in ['Hora', 'Plan', 'Obra_Social']:
                if col in extra_user.columns:
                    cols_to_keep.append(col)
            
            extra_user = extra_user[[col for col in cols_to_keep if col in extra_user.columns]]
        else:
            extra_user = pd.DataFrame(columns=['HC', 'Fecha', 'Nombre', 'Monto'])
        
        # Extra en registros del hospital (en contra)
        extra_hospital_mask = merged['_merge'] == 'right_only'
        extra_hospital = merged[extra_hospital_mask].copy()
        
        print(f"Registros extra en hospital: {len(extra_hospital)}")
        
        if not extra_hospital.empty:
            extra_hospital['Nombre'] = extra_hospital.get('Nombre_hospital', extra_hospital.get('Nombre_usuario', ''))
            extra_hospital['Monto'] = extra_hospital.get('Monto_hospital', extra_hospital.get('Monto_usuario', 0.0))
            
            # Agregar información adicional del hospital
            hospital_additional = hospital_df.copy()
            if not hospital_additional.empty:
                extra_hospital = extra_hospital.merge(
                    hospital_additional,
                    on=['HC', 'Fecha'], 
                    how='left',
                    suffixes=('', '_extra')
                )
            
            # Seleccionar columnas relevantes
            cols_to_keep = ['HC', 'Fecha', 'Nombre', 'Monto', 'Archivo_Origen', 'Tipo_Archivo']
            for col in ['Cobertura', 'Desgrupo', 'Desc_Cob', 'Obra_Social']:
                if col in extra_hospital.columns:
                    cols_to_keep.append(col)
            
            extra_hospital = extra_hospital[[col for col in cols_to_keep if col in extra_hospital.columns]]
        else:
            extra_hospital = pd.DataFrame(columns=['HC', 'Fecha', 'Nombre', 'Monto', 'Archivo_Origen', 'Tipo_Archivo'])
        
        # Ordenar resultados
        extra_user = extra_user.sort_values(by=['Nombre', 'Fecha']).reset_index(drop=True)
        extra_hospital = extra_hospital.sort_values(by=['Nombre', 'Fecha']).reset_index(drop=True)
        
        # Crear estadísticas por historia clínica
        user_stats = pd.DataFrame()
        hospital_stats = pd.DataFrame()
        
        if not extra_user.empty:
            user_stats = extra_user.groupby('HC').agg({
                'Fecha': 'count',
                'Monto': 'sum',
                'Nombre': 'first'
            }).rename(columns={'Fecha': 'Cantidad_Fechas'}).reset_index()
        
        if not extra_hospital.empty:
            hospital_stats = extra_hospital.groupby('HC').agg({
                'Fecha': 'count', 
                'Monto': 'sum',
                'Nombre': 'first'
            }).rename(columns={'Fecha': 'Cantidad_Fechas'}).reset_index()
        
        # Crear resumen general
        summary = pd.DataFrame({
            'Concepto': [
                'Registros en mi archivo',
                'Registros en hospital (total)',
                'Extra en mi registro (registros)',
                'Extra en hospital (registros)',
                'Extra en mi registro (pacientes únicos)',
                'Extra en hospital (pacientes únicos)',
                'Diferencia neta (registros)',
                'Diferencia neta (monto)'
            ],
            'Cantidad': [
                len(user_df),
                len(hospital_df),
                len(extra_user),
                len(extra_hospital),
                len(user_stats),
                len(hospital_stats),
                len(extra_user) - len(extra_hospital),
                extra_user['Monto'].sum() - extra_hospital['Monto'].sum()
            ]
        })
        
        # Crear resumen por tipo de archivo del hospital
        hospital_summary = pd.DataFrame(columns=['Tipo_Archivo', 'Cantidad_Registros', 'Monto_Total'])
        if not extra_hospital.empty and 'Tipo_Archivo' in extra_hospital.columns:
            hospital_summary = extra_hospital.groupby('Tipo_Archivo').agg({
                'HC': 'count',
                'Monto': 'sum'
            }).reset_index()
            hospital_summary.columns = ['Tipo_Archivo', 'Cantidad_Registros', 'Monto_Total']

        # Determinar ruta de salida
        if output_dir is None:
            output_dir = os.path.dirname(user_file)
        
        output_file = os.path.join(output_dir, 'discrepancias_pacientes.xlsx')

        # Exportar resultados
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # Resumen general
            summary.to_excel(writer, sheet_name='Resumen_General', index=False)
            
            # Resumen por tipo de hospital
            hospital_summary.to_excel(writer, sheet_name='Resumen_Hospital', index=False)
            
            # Registros extra del usuario
            extra_user.to_excel(writer, sheet_name='Extra_Mi_Registro', index=False)
            
            # Registros extra del hospital
            extra_hospital.to_excel(writer, sheet_name='Extra_Hospital', index=False)
            
            # Estadísticas por paciente - usuario
            if not user_stats.empty:
                user_stats.to_excel(writer, sheet_name='Stats_Mi_Registro', index=False)
            
            # Estadísticas por paciente - hospital  
            if not hospital_stats.empty:
                hospital_stats.to_excel(writer, sheet_name='Stats_Hospital', index=False)
            
            # Datos originales para referencia
            user_df.to_excel(writer, sheet_name='Datos_Usuario', index=False)
            hospital_df.to_excel(writer, sheet_name='Datos_Hospital', index=False)

        print(f"Archivo de salida generado: {output_file}")
        return output_file

    except Exception as e:
        print(f"Error en compare_records: {str(e)}")
        raise Exception(f"Error procesando archivos: {str(e)}")

class ComparadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Comparador de Registros de Pacientes - V: 2.3.2 - by RenzoRossiBrun")
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
                try:
                    os.startfile(output_file)  # Windows
                except:
                    try:
                        os.system(f'open "{output_file}"')  # macOS
                    except:
                        os.system(f'xdg-open "{output_file}"')  # Linux
            
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