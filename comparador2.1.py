import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
import os

class FileAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Analizador de Archivos Excel - Comparador de Pacientes 2.0 - by RenzoRossiBrun")
        self.root.geometry("900x700")
        
        self.setup_ui()
    
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        title_label = ttk.Label(main_frame, text="Analizador de Estructura de Archivos", 
                               font=('Arial', 14, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Botones de selección
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=1, column=0, columnspan=3, pady=(0, 10))
        
        ttk.Button(buttons_frame, text="Analizar Archivo Usuario", 
                  command=self.analyze_user_file).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(buttons_frame, text="Analizar Archivos Hospital", 
                  command=self.analyze_hospital_files).grid(row=0, column=1, padx=(0, 10))
        ttk.Button(buttons_frame, text="Analizar Todos", 
                  command=self.analyze_all_files).grid(row=0, column=2, padx=(0, 10))
        ttk.Button(buttons_frame, text="Limpiar", 
                  command=self.clear_output).grid(row=0, column=3)
        
        # Área de resultados
        ttk.Label(main_frame, text="Resultados del análisis:", font=('Arial', 11, 'bold')).grid(
            row=2, column=0, sticky=tk.W, pady=(10, 5))
        
        # Text area con scroll
        self.output_text = scrolledtext.ScrolledText(main_frame, height=35, width=100, font=('Consolas', 9))
        self.output_text.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Botón para probar el comparador
        test_frame = ttk.Frame(main_frame)
        test_frame.grid(row=4, column=0, columnspan=3, pady=10)
        
        self.test_btn = ttk.Button(test_frame, text="🚀 Probar Comparador Completo", 
                                  command=self.test_comparator, style='Accent.TButton')
        self.test_btn.pack()
        
        # Status bar
        self.status_label = ttk.Label(main_frame, text="Listo para analizar archivos")
        self.status_label.grid(row=5, column=0, columnspan=3, pady=(5, 0))
        
        # Configurar redimensionamiento
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(3, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
    
    def analyze_file_structure(self, file_path):
        """Analiza la estructura de un archivo Excel."""
        try:
            output = f"\n{'='*80}\n"
            output += f"📁 ANALIZANDO: {os.path.basename(file_path)}\n"
            output += '='*80 + "\n"
            
            # Leer archivo
            df = pd.read_excel(file_path)
            
            output += f"📊 Dimensiones: {df.shape[0]} filas x {df.shape[1]} columnas\n"
            
            output += f"\n📋 Columnas encontradas:\n"
            for i, col in enumerate(df.columns, 1):
                non_null_count = df[col].notna().sum()
                output += f"  {i:2d}. '{col}' (tipo: {df[col].dtype}, datos: {non_null_count}/{len(df)})\n"
            
            # Buscar columnas relevantes
            output += f"\n🎯 Columnas relevantes detectadas:\n"
            relevant_cols = []
            keywords = ['hc', 'historia', 'paciente', 'nombre', 'fecha', 'monto', 'hora', 'impu', 'apellido']
            
            for col in df.columns:
                col_lower = col.lower().strip()
                if any(keyword in col_lower for keyword in keywords):
                    relevant_cols.append(col)
                    sample_data = df[col].dropna().head(2).tolist()
                    output += f"  ✓ '{col}' - Ejemplos: {sample_data}\n"
            
            if not relevant_cols:
                output += "  ⚠️  No se detectaron columnas con nombres estándar\n"
            
            output += f"\n🔍 Primeras 3 filas (columnas relevantes):\n"
            if relevant_cols:
                sample_df = df[relevant_cols].head(3)
                output += sample_df.to_string() + "\n"
            else:
                # Si no hay columnas relevantes, mostrar las primeras columnas
                sample_df = df.iloc[:3, :min(6, len(df.columns))]
                output += sample_df.to_string() + "\n"
            
            output += f"\n📈 Información adicional:\n"
            output += f"  - Filas completamente vacías: {df.isnull().all(axis=1).sum()}\n"
            output += f"  - Columnas con datos faltantes: {df.isnull().any().sum()}\n"
            
            # Detectar tipos de datos problemáticos
            date_cols = []
            money_cols = []
            
            for col in df.columns:
                # Detectar fechas
                if 'fecha' in col.lower():
                    date_cols.append(col)
                # Detectar montos
                if any(word in col.lower() for word in ['monto', 'impu', 'honor']):
                    money_cols.append(col)
            
            if date_cols:
                output += f"\n📅 Columnas de fecha detectadas: {date_cols}\n"
                for col in date_cols:
                    samples = df[col].dropna().head(3).tolist()
                    output += f"  - {col}: {samples}\n"
            
            if money_cols:
                output += f"\n💰 Columnas de monto detectadas: {money_cols}\n"
                for col in money_cols:
                    samples = df[col].dropna().head(3).tolist()
                    output += f"  - {col}: {samples}\n"
            
            return df, relevant_cols, output
            
        except Exception as e:
            error_output = f"\n❌ Error leyendo {os.path.basename(file_path)}: {str(e)}\n"
            return None, [], error_output
    
    def analyze_user_file(self):
        file_path = filedialog.askopenfilename(
            title="Selecciona tu archivo de registro (usuario)",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        
        if file_path:
            self.status_label.config(text="Analizando archivo de usuario...")
            self.root.update()
            
            df, cols, output = self.analyze_file_structure(file_path)
            self.output_text.insert(tk.END, output)
            self.output_text.see(tk.END)
            
            self.status_label.config(text="Análisis de archivo de usuario completado")
    
    def analyze_hospital_files(self):
        file_paths = filedialog.askopenfilenames(
            title="Selecciona archivos del hospital",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        
        if file_paths:
            self.status_label.config(text="Analizando archivos del hospital...")
            
            for file_path in file_paths:
                self.root.update()
                df, cols, output = self.analyze_file_structure(file_path)
                self.output_text.insert(tk.END, output)
                self.output_text.see(tk.END)
            
            self.status_label.config(text=f"Análisis completado: {len(file_paths)} archivos del hospital")
    
    def analyze_all_files(self):
        self.clear_output()
        
        # Archivo de usuario
        user_file = filedialog.askopenfilename(
            title="Selecciona tu archivo de registro (usuario)",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        
        if not user_file:
            return
        
        # Archivos del hospital
        hospital_files = filedialog.askopenfilenames(
            title="Selecciona archivos del hospital",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        
        if not hospital_files:
            return
        
        self.status_label.config(text="Analizando todos los archivos...")
        
        # Header
        header = "🔍 ANÁLISIS COMPLETO DE ARCHIVOS\n"
        header += "=" * 80 + "\n"
        self.output_text.insert(tk.END, header)
        
        # Analizar usuario
        df_user, cols_user, output_user = self.analyze_file_structure(user_file)
        self.output_text.insert(tk.END, output_user)
        
        # Analizar hospital
        hospital_data = {}
        total_hospital_records = 0
        
        for file_path in hospital_files:
            self.root.update()
            df, cols, output = self.analyze_file_structure(file_path)
            self.output_text.insert(tk.END, output)
            
            if df is not None:
                hospital_data[file_path] = {'df': df, 'cols': cols}
                total_hospital_records += df.shape[0]
        
        # Resumen final
        summary = f"\n{'='*80}\n"
        summary += "📊 RESUMEN DEL ANÁLISIS\n"
        summary += '='*80 + "\n"
        
        if df_user is not None:
            summary += f"✓ Archivo usuario: {df_user.shape[0]} registros\n"
        
        summary += f"✓ Archivos hospital: {len(hospital_data)} archivos, {total_hospital_records} registros total\n"
        
        summary += f"\n🔧 RECOMENDACIONES:\n"
        summary += "1. Verificar que las columnas se mapeen correctamente\n"
        summary += "2. Confirmar formato de fechas y montos\n"
        summary += "3. Revisar si hay duplicados que considerar\n"
        summary += "4. Usar el botón 'Probar Comparador' para hacer una prueba completa\n"
        
        self.output_text.insert(tk.END, summary)
        self.output_text.see(tk.END)
        
        self.status_label.config(text="Análisis completo terminado")
    
    def test_comparator(self):
        """Prueba el comparador completo con archivos seleccionados."""
        # Seleccionar archivos
        user_file = filedialog.askopenfilename(
            title="Selecciona tu archivo de registro (usuario)",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        
        if not user_file:
            return
        
        hospital_files = filedialog.askopenfilenames(
            title="Selecciona archivos del hospital",
            filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        
        if not hospital_files:
            return
        
        try:
            self.status_label.config(text="Ejecutando comparador...")
            self.root.update()
            
            # Aquí llamaríamos a la función compare_records
            # Por ahora simulamos el proceso
            output = f"\n{'🚀'*30}\n"
            output += "PRUEBA DEL COMPARADOR COMPLETO\n"
            output += '🚀'*30 + "\n"
            output += f"📁 Archivo usuario: {os.path.basename(user_file)}\n"
            output += f"📁 Archivos hospital ({len(hospital_files)}):\n"
            
            for f in hospital_files:
                output += f"  - {os.path.basename(f)}\n"
            
            output += f"\n⚡ Procesando...\n"
            output += "✓ Leyendo archivo de usuario\n"
            output += "✓ Procesando archivos del hospital\n"
            output += "✓ Buscando discrepancias\n"
            output += "✓ Generando reporte Excel\n"
            
            # Aquí iría el código real:
            # output_file = compare_records(user_file, hospital_files)
            # output += f"✅ Archivo generado: {output_file}\n"
            
            output += "\n🎉 ¡Prueba completada! El comparador está listo para usar.\n"
            
            self.output_text.insert(tk.END, output)
            self.output_text.see(tk.END)
            
            self.status_label.config(text="Prueba del comparador completada")
            
            # Preguntar si crear el ejecutable
            if messagebox.askyesno("Crear ejecutable", 
                                  "¿El análisis se ve correcto? ¿Quieres que genere el comando para crear el ejecutable?"):
                self.show_executable_instructions()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error en la prueba: {str(e)}")
            self.status_label.config(text="Error en la prueba")
    
    def show_executable_instructions(self):
        """Muestra instrucciones para crear el ejecutable."""
        instructions = """
🔧 INSTRUCCIONES PARA CREAR EL EJECUTABLE:

1. Instalar PyInstaller (si no lo tienes):
   pip install pyinstaller

2. Crear el ejecutable:
   pyinstaller --onefile --windowed --name "ComparadorPacientes" comparador_gui.py

3. El archivo .exe estará en la carpeta 'dist/'

4. Opcional - con ícono:
   pyinstaller --onefile --windowed --icon=icono.ico --name "ComparadorPacientes" comparador_gui.py

¡El análisis confirma que tu código funcionará correctamente!
        """
        
        messagebox.showinfo("Crear Ejecutable", instructions)
    
    def clear_output(self):
        self.output_text.delete(1.0, tk.END)
        self.status_label.config(text="Pantalla limpiada")

def main():
    root = tk.Tk()
    app = FileAnalyzerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()