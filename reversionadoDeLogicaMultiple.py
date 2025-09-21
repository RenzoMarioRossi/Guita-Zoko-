import pandas as pd
import os
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import subprocess
import sys
import threading
from datetime import datetime

class HistoriaClinicaProcessor:
    def __init__(self):
        self.archivos_control = []  
        self.archivos_hospital = []
        self.df_presentes = None
        self.df_pagos_en_contra = None  
        self.archivo_salida = "presentes_no_pagados.xlsx"
        self.archivo_salida_contra = "pagos_en_contra.xlsx"
        
    def normalizar_hc(self, valor):
        """
        Normaliza los valores de historia clínica para comparación
        Maneja casos especiales como HC=0 y pacientes sin HC
        """
        if pd.isna(valor):
            return None
        
        # Convertir a string y limpiar
        valor_str = str(valor).strip().lower()
        
        # Pacientes sin HC
        if any(palabra in valor_str for palabra in ['sin h.c', 'sin hc', 'sin historia']):
            return f"SIN_HC_{valor_str.replace(' ', '_')}"
        
        # Extraer números de la cadena
        numeros = re.findall(r'\d+', valor_str)
        
        if numeros:
            numero = int(numeros[0])
            # Manejar HC = 0 como caso especial
            if numero == 0:
                return f"HC_0_{valor_str.replace(' ', '_')}"
            return numero
        
        # Si no hay números pero hay contenido, tratarlo como caso especial
        if valor_str:
            return f"ESPECIAL_{valor_str.replace(' ', '_')}"
        
        return None
    
    def encontrar_columna_hc(self, df):
        """
        Encuentra la columna que contiene las historias clínicas
        """
        posibles_nombres = ['hc', 'HC', 'historia', 'Historia', 'historia clinica', 
                           'Historia Clinica', 'historia_clinica', 'Historia_Clinica',
                           'HC_PACIENTE', 'hc_paciente', 'numero_historia', 'Numero_Historia', 'HISTORIA']
        
        # Buscar por nombre exacto
        for col in df.columns:
            if str(col).strip().lower() in [nombre.lower() for nombre in posibles_nombres]:
                return col
        
        # Buscar por contenido parcial
        for col in df.columns:
            col_lower = str(col).lower()
            if any(palabra in col_lower for palabra in ['historia', 'hc']):
                return col
                
        return None
    
    def encontrar_columna_estado(self, df):
        """
        Encuentra la columna que contiene el estado
        """
        posibles_nombres = ['estado', 'Estado', 'ESTADO', 'status', 'Status', 'STATUS']
        
        for col in df.columns:
            if str(col).strip().lower() in [nombre.lower() for nombre in posibles_nombres]:
                return col
                
        return None
    
    def procesar_archivos_control(self, callback=None):
        """
        Procesa múltiples archivos de control y extrae solo los presentes
        Cada fila representa una visita independiente que debe ser pagada
        """
        try:
            if callback:
                callback(f"Procesando {len(self.archivos_control)} archivos de control...")
            
            df_todos_presentes = []
            
            for archivo in self.archivos_control:
                if callback:
                    callback(f"Procesando: {os.path.basename(archivo)}")
                
                df = pd.read_excel(archivo)
                
                # Encontrar columnas relevantes
                col_hc = self.encontrar_columna_hc(df)
                col_estado = self.encontrar_columna_estado(df)
                
                if not col_hc:
                    if callback:
                        callback(f"❌ No se encontró columna HC en {os.path.basename(archivo)}")
                    continue
                if not col_estado:
                    if callback:
                        callback(f"❌ No se encontró columna Estado en {os.path.basename(archivo)}")
                    continue
                
                # Filtrar solo los presentes
                df_filtrado = df[df[col_estado].str.upper() == 'P'].copy()
                
                # Normalizar HC para comparar
                df_filtrado['HC_NORMALIZADA'] = df_filtrado[col_hc].apply(self.normalizar_hc)
                
                # Agregar identificador de archivo y fila única
                df_filtrado['ARCHIVO_ORIGEN'] = os.path.basename(archivo)
                df_filtrado['ID_FILA'] = (os.path.basename(archivo) + '_' + 
                                        df_filtrado.index.astype(str) + '_' + 
                                        df_filtrado['HC_NORMALIZADA'].astype(str))
                
                # Eliminar filas donde no se pudo normalizar la HC
                antes_filtro = len(df_filtrado)
                df_filtrado = df_filtrado.dropna(subset=['HC_NORMALIZADA'])
                despues_filtro = len(df_filtrado)
                
                if antes_filtro != despues_filtro:
                    if callback:
                        callback(f"⚠️  {antes_filtro - despues_filtro} filas eliminadas por HC inválida")
                
                df_todos_presentes.append(df_filtrado)
                if callback:
                    callback(f"✅ {len(df_filtrado)} registros presentes")
            
            # Combinar todos los archivos
            if df_todos_presentes:
                self.df_presentes = pd.concat(df_todos_presentes, ignore_index=True)
                
                total_presentes = len(self.df_presentes)
                archivos_procesados = len([df for df in df_todos_presentes if len(df) > 0])
                
                if callback:
                    callback(f"📊 Resumen archivos de control:")
                    callback(f"  Archivos procesados exitosamente: {archivos_procesados}")
                    callback(f"  Total registros presentes: {total_presentes}")
                
                # Mostrar estadísticas de HC especiales
                hc_especiales = self.df_presentes[self.df_presentes['HC_NORMALIZADA'].astype(str).str.contains('SIN_HC|HC_0|ESPECIAL', na=False)]
                if len(hc_especiales) > 0:
                    if callback:
                        callback(f"  HC especiales (HC=0, sin HC, etc.): {len(hc_especiales)}")
                
                return True
            else:
                if callback:
                    callback("❌ No se pudieron procesar archivos de control")
                return False
                
        except Exception as e:
            if callback:
                callback(f"Error procesando archivos de control: {str(e)}")
            return False
    
    def procesar_archivos_hospital(self, callback=None):
        """
        Procesa los archivos del hospital y elimina las HC que coinciden
        También identifica pagos "en contra" (en hospital pero no en control)
        """
        if self.df_presentes is None:
            if callback:
                callback("Error: No hay datos de presentes para procesar")
            return False
        
        hc_presentes = set(self.df_presentes['HC_NORMALIZADA'].tolist())
        ids_presentes = set(self.df_presentes['ID_FILA'].tolist())
        ids_encontrados_hospital = set()
        
        # Para detectar pagos "en contra"
        todas_hc_hospital = []
        
        if callback:
            callback(f"Historias clínicas únicas presentes: {len(hc_presentes)}")
            callback(f"Total de filas/visitas presentes: {len(ids_presentes)}")
        
        for archivo_hospital in self.archivos_hospital:
            try:
                if callback:
                    callback(f"Procesando archivo del hospital: {os.path.basename(archivo_hospital)}")
                
                df_hospital = pd.read_excel(archivo_hospital)
                
                # Encontrar columna HC en archivo del hospital
                col_hc_hospital = self.encontrar_columna_hc(df_hospital)
                
                if not col_hc_hospital:
                    if callback:
                        callback(f"  No se encontró columna HC en {os.path.basename(archivo_hospital)}")
                    continue
                
                if callback:
                    callback(f"  Columna Historia Clínica encontrada: {col_hc_hospital}")
                
                # Procesar cada fila del hospital
                for idx, row in df_hospital.iterrows():
                    hc_norm = self.normalizar_hc(row[col_hc_hospital])
                    if hc_norm is not None:
                        # Agregar a lista completa para análisis "en contra"
                        fila_hospital = row.copy()
                        fila_hospital['HC_NORMALIZADA'] = hc_norm
                        fila_hospital['ARCHIVO_HOSPITAL'] = os.path.basename(archivo_hospital)
                        todas_hc_hospital.append(fila_hospital)
                        
                        # Buscar coincidencias con presentes para marcar como pagadas
                        filas_candidatas = self.df_presentes[
                            (self.df_presentes['HC_NORMALIZADA'] == hc_norm) &
                            (~self.df_presentes['ID_FILA'].isin(ids_encontrados_hospital))
                        ]
                        
                        if len(filas_candidatas) > 0:
                            # Marcar la primera fila encontrada como pagada
                            id_fila = filas_candidatas.iloc[0]['ID_FILA']
                            ids_encontrados_hospital.add(id_fila)
                
                if callback:
                    callback(f"  HC válidas encontradas en este archivo: {len([h for h in todas_hc_hospital if h['ARCHIVO_HOSPITAL'] == os.path.basename(archivo_hospital)])}")
                
            except Exception as e:
                if callback:
                    callback(f"Error procesando {os.path.basename(archivo_hospital)}: {str(e)}")
        
        # Procesar pagos "en contra" (están en hospital pero no en control)
        if todas_hc_hospital:
            df_hospital_completo = pd.DataFrame(todas_hc_hospital)
            hc_solo_hospital = df_hospital_completo[~df_hospital_completo['HC_NORMALIZADA'].isin(hc_presentes)]
            
            if len(hc_solo_hospital) > 0:
                self.df_pagos_en_contra = hc_solo_hospital.copy()
                if callback:
                    callback(f"🔍 Pagos 'en contra' encontrados: {len(hc_solo_hospital)} filas")
            else:
                if callback:
                    callback(f"✅ No se encontraron pagos 'en contra'")
        
        # Eliminar las filas encontradas en el hospital del DataFrame de presentes
        ids_no_pagados = ids_presentes - ids_encontrados_hospital
        
        # Filtrar el DataFrame
        self.df_presentes = self.df_presentes[
            self.df_presentes['ID_FILA'].isin(ids_no_pagados)
        ].copy()
        
        if callback:
            callback(f"Resumen:")
            callback(f"Total de filas/visitas presentes: {len(ids_presentes)}")
            callback(f"Filas encontradas como pagadas en hospital: {len(ids_encontrados_hospital)}")
            callback(f"Filas NO PAGADAS (discrepancias a favor): {len(ids_no_pagados)}")
        
        # Mostrar estadísticas por paciente
        if len(self.df_presentes) > 0:
            pacientes_no_pagados = self.df_presentes['HC_NORMALIZADA'].nunique()
            if callback:
                callback(f"Pacientes únicos con visitas no pagadas: {pacientes_no_pagados}")
        
        return True
    
    def guardar_resultados(self, callback=None):
        """
        Guarda los resultados finales en archivos Excel
        """
        resultados_guardados = 0
        
        # Guardar discrepancias "a favor" (presentes no pagados)
        if self.df_presentes is not None and not self.df_presentes.empty:
            try:
                # Eliminar las columnas auxiliares antes de guardar
                df_salida = self.df_presentes.drop(['HC_NORMALIZADA', 'ID_FILA'], axis=1)
                
                # Ordenar por archivo origen y luego por HC para mejor visualización
                if 'ARCHIVO_ORIGEN' in df_salida.columns:
                    col_hc_original = self.encontrar_columna_hc(df_salida)
                    if col_hc_original:
                        df_salida = df_salida.sort_values(by=['ARCHIVO_ORIGEN', col_hc_original])
                
                df_salida.to_excel(self.archivo_salida, index=False)
                if callback:
                    callback(f"💰 Discrepancias A FAVOR guardadas: {self.archivo_salida}")
                    callback(f"   Total visitas no pagadas: {len(df_salida)}")
                resultados_guardados += 1
                
            except Exception as e:
                if callback:
                    callback(f"Error guardando discrepancias a favor: {str(e)}")
        else:
            if callback:
                callback(f"✅ Sin discrepancias A FAVOR - Todas las visitas fueron pagadas")
        
        # Guardar discrepancias "en contra" (pagos sin correspondencia)
        if self.df_pagos_en_contra is not None and not self.df_pagos_en_contra.empty:
            try:
                # Eliminar la columna auxiliar antes de guardar
                df_contra = self.df_pagos_en_contra.drop(['HC_NORMALIZADA'], axis=1)
                
                # Ordenar por archivo hospital
                df_contra = df_contra.sort_values(by='ARCHIVO_HOSPITAL')
                
                df_contra.to_excel(self.archivo_salida_contra, index=False)
                if callback:
                    callback(f"⚠️  Discrepancias EN CONTRA guardadas: {self.archivo_salida_contra}")
                    callback(f"   Total pagos sin correspondencia: {len(df_contra)}")
                resultados_guardados += 1
                
            except Exception as e:
                if callback:
                    callback(f"Error guardando discrepancias en contra: {str(e)}")
        else:
            if callback:
                callback(f"✅ Sin discrepancias EN CONTRA - Todos los pagos corresponden")
        
        return resultados_guardados > 0
    
    def abrir_resultados(self, callback=None):
        """
        Abre los archivos de resultados
        """
        archivos_a_abrir = []
        
        # Verificar qué archivos existen
        if os.path.exists(self.archivo_salida):
            archivos_a_abrir.append(self.archivo_salida)
        if os.path.exists(self.archivo_salida_contra):
            archivos_a_abrir.append(self.archivo_salida_contra)
        
        if not archivos_a_abrir:
            if callback:
                callback("No hay archivos de resultados para abrir")
            return
        
        for archivo in archivos_a_abrir:
            try:
                if sys.platform == "win32":
                    os.startfile(archivo)
                elif sys.platform == "darwin":
                    subprocess.run(["open", archivo])
                else:
                    subprocess.run(["xdg-open", archivo])
                
                if callback:
                    callback(f"Abriendo: {archivo}")
                
            except Exception as e:
                if callback:
                    callback(f"No se pudo abrir automáticamente {archivo}: {str(e)}")
                    callback(f"Por favor, abra manualmente: {archivo}")

class HistoriaClinicaGUI:
    def __init__(self, root):
        self.root = root
        self.processor = HistoriaClinicaProcessor()
        self.setup_gui()
        
    def setup_gui(self):
        self.root.title("Procesador de Historias Clínicas - v4.3 - by RenzoRossiBrun CC2025")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        
        # Configurar estilo
        style = ttk.Style()
        style.theme_use('clam')
        
        # Frame principal con padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurar grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # Título
        title_label = ttk.Label(main_frame, text="Procesador de Historias Clínicas", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, pady=(0, 20))
        
        # Frame para archivos de control
        control_frame = ttk.LabelFrame(main_frame, text="Archivos de Control (Usuario)", padding="10")
        control_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        control_frame.columnconfigure(0, weight=1)
        
        self.control_files_var = tk.StringVar(value="Ningún archivo seleccionado")
        control_label = ttk.Label(control_frame, textvariable=self.control_files_var)
        control_label.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        control_btn = ttk.Button(control_frame, text="Seleccionar Archivos", 
                                command=self.seleccionar_archivos_control)
        control_btn.grid(row=0, column=1)
        
        # Frame para archivos de hospital
        hospital_frame = ttk.LabelFrame(main_frame, text="Archivos del Hospital", padding="10")
        hospital_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        hospital_frame.columnconfigure(0, weight=1)
        
        self.hospital_files_var = tk.StringVar(value="Ningún archivo seleccionado")
        hospital_label = ttk.Label(hospital_frame, textvariable=self.hospital_files_var)
        hospital_label.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        hospital_btn = ttk.Button(hospital_frame, text="Seleccionar Archivos", 
                                 command=self.seleccionar_archivos_hospital)
        hospital_btn.grid(row=0, column=1)
        
        # Frame para botones de acción
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(row=3, column=0, pady=20)
        
        self.process_btn = ttk.Button(action_frame, text="Procesar Archivos", 
                                     command=self.procesar_archivos, state='disabled',
                                     style='Accent.TButton')
        self.process_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.clear_btn = ttk.Button(action_frame, text="Limpiar Todo", 
                                   command=self.limpiar_todo)
        self.clear_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        self.open_results_btn = ttk.Button(action_frame, text="Abrir Resultados", 
                                          command=self.abrir_resultados, state='disabled')
        self.open_results_btn.pack(side=tk.LEFT)
        
        # Frame para el log de salida
        log_frame = ttk.LabelFrame(main_frame, text="Log de Procesamiento", padding="5")
        log_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # Crear el widget de texto con scroll
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, 
                                                 state=tk.DISABLED, height=15)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Barra de progreso
        self.progress_var = tk.StringVar(value="Listo para procesar")
        progress_label = ttk.Label(main_frame, textvariable=self.progress_var)
        progress_label.grid(row=5, column=0, pady=(10, 0))
        
        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress_bar.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # Log inicial
        self.log_message("=== PROCESADOR DE HISTORIAS CLÍNICAS ===")
        self.log_message("1. Seleccione los archivos de control (varios meses)")
        self.log_message("2. Seleccione los archivos del hospital") 
        self.log_message("3. Haga clic en 'Procesar Archivos'")
        self.log_message("")
        
    def log_message(self, message):
        """Agrega un mensaje al log con timestamp"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.config(state=tk.DISABLED)
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def seleccionar_archivos_control(self):
        """Seleccionar archivos de control"""
        archivos = filedialog.askopenfilenames(
            title="Seleccionar archivos de control (Excel del usuario - varios meses)",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        
        if archivos:
            self.processor.archivos_control = list(archivos)
            count = len(archivos)
            self.control_files_var.set(f"{count} archivo{'s' if count != 1 else ''} seleccionado{'s' if count != 1 else ''}")
            
            self.log_message(f"Archivos de control seleccionados: {count}")
            for archivo in archivos:
                self.log_message(f"  • {os.path.basename(archivo)}")
            
            self.verificar_archivos_completos()
        
    def seleccionar_archivos_hospital(self):
        """Seleccionar archivos del hospital"""
        archivos = filedialog.askopenfilenames(
            title="Seleccionar archivos del hospital (Excel)",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        
        if archivos:
            self.processor.archivos_hospital = list(archivos)
            count = len(archivos)
            self.hospital_files_var.set(f"{count} archivo{'s' if count != 1 else ''} seleccionado{'s' if count != 1 else ''}")
            
            self.log_message(f"Archivos del hospital seleccionados: {count}")
            for archivo in archivos:
                self.log_message(f"  • {os.path.basename(archivo)}")
                
            self.verificar_archivos_completos()
    
    def verificar_archivos_completos(self):
        """Verifica si se han seleccionado todos los archivos necesarios"""
        if self.processor.archivos_control and self.processor.archivos_hospital:
            self.process_btn.config(state='normal')
            self.log_message("✅ Todos los archivos seleccionados. Listo para procesar.")
        else:
            self.process_btn.config(state='disabled')
    
    def procesar_archivos(self):
        """Procesa los archivos en un hilo separado"""
        # Deshabilitar botones durante el procesamiento
        self.process_btn.config(state='disabled')
        self.clear_btn.config(state='disabled')
        self.progress_var.set("Procesando...")
        self.progress_bar.start()
        
        # Limpiar log anterior (mantener instrucciones iniciales)
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete('1.0', tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        self.log_message("=== INICIANDO PROCESAMIENTO ===")
        
        # Ejecutar en hilo separado para no bloquear la UI
        thread = threading.Thread(target=self.ejecutar_procesamiento, daemon=True)
        thread.start()
    
    def ejecutar_procesamiento(self):
        """Ejecuta el procesamiento completo"""
        try:
            # Paso 1: Procesar archivos de control
            self.log_message("Paso 1: Procesando archivos de control...")
            if not self.processor.procesar_archivos_control(callback=self.log_message):
                self.finalizar_procesamiento("Error en el procesamiento de archivos de control", False)
                return
            
            # Paso 2: Procesar archivos del hospital
            self.log_message("\nPaso 2: Procesando archivos del hospital...")
            if not self.processor.procesar_archivos_hospital(callback=self.log_message):
                self.finalizar_procesamiento("Error en el procesamiento de archivos del hospital", False)
                return
            
            # Paso 3: Guardar resultados
            self.log_message("\nPaso 3: Guardando resultados...")
            if not self.processor.guardar_resultados(callback=self.log_message):
                self.finalizar_procesamiento("Error guardando resultados", False)
                return
            
            # Mostrar resumen final
            self.mostrar_resumen_final()
            
            self.finalizar_procesamiento("¡Proceso completado exitosamente!", True)
            
        except Exception as e:
            self.log_message(f"Error inesperado: {str(e)}")
            self.finalizar_procesamiento(f"Error inesperado: {str(e)}", False)
    
    def mostrar_resumen_final(self):
        """Muestra el resumen final del procesamiento"""
        self.log_message("\n=== RESUMEN FINAL ===")
        
        # Resultados A FAVOR
        if self.processor.df_presentes is not None and len(self.processor.df_presentes) > 0:
            self.log_message(f"💰 Discrepancias A FAVOR: {len(self.processor.df_presentes)} visitas")
            self.log_message("   (Pacientes presentes que NO fueron pagados por el hospital)")
        else:
            self.log_message("✅ Sin discrepancias A FAVOR - Todas las visitas fueron pagadas")
        
        # Resultados EN CONTRA
        if self.processor.df_pagos_en_contra is not None and len(self.processor.df_pagos_en_contra) > 0:
            self.log_message(f"⚠️  Discrepancias EN CONTRA: {len(self.processor.df_pagos_en_contra)} pagos")
            self.log_message("   (Pagos del hospital que NO corresponden a tus archivos de control)")
        else:
            self.log_message("✅ Sin discrepancias EN CONTRA - Todos los pagos corresponden")
        
        self.log_message(f"\nArchivos generados:")
        if os.path.exists(self.processor.archivo_salida):
            self.log_message(f"• A favor: {self.processor.archivo_salida}")
        if os.path.exists(self.processor.archivo_salida_contra):
            self.log_message(f"• En contra: {self.processor.archivo_salida_contra}")
        
        self.log_message("\n📋 IMPORTANTE:")
        self.log_message("• A FAVOR = El hospital te debe dinero")  
        self.log_message("• EN CONTRA = El hospital te pagó algo que no corresponde")
        self.log_message("• Cada fila representa una visita/pago independiente")
        self.log_message("• HC=0 y 'sin HC' también se consideran válidos para cobro")
    
    def finalizar_procesamiento(self, mensaje, exito):
        """Finaliza el procesamiento y restaura la UI"""
        self.root.after(0, lambda: self._finalizar_procesamiento_ui(mensaje, exito))
    
    def _finalizar_procesamiento_ui(self, mensaje, exito):
        """Finaliza el procesamiento en el hilo principal"""
        self.progress_bar.stop()
        self.progress_var.set(mensaje)
        self.log_message(f"\n{mensaje}")
        
        # Rehabilitar botones
        self.verificar_archivos_completos()
        self.clear_btn.config(state='normal')
        
        if exito:
            # Habilitar botón de abrir resultados si hay archivos
            if (os.path.exists(self.processor.archivo_salida) or 
                os.path.exists(self.processor.archivo_salida_contra)):
                self.open_results_btn.config(state='normal')
            
            # Mostrar mensaje de éxito
            messagebox.showinfo("Proceso Completado", 
                              "El procesamiento se completó exitosamente.\n\n"
                              "Revise el log para ver los detalles y haga clic en "
                              "'Abrir Resultados' para ver los archivos generados.")
        else:
            # Mostrar mensaje de error
            messagebox.showerror("Error en el Proceso", mensaje)
    
    def abrir_resultados(self):
        """Abre los archivos de resultados"""
        self.processor.abrir_resultados(callback=self.log_message)
    
    def limpiar_todo(self):
        """Limpia todos los datos y reinicia la interfaz"""
        # Confirmar con el usuario
        if messagebox.askyesno("Confirmar", "¿Está seguro de que desea limpiar todo?"):
            # Reiniciar el processor
            self.processor = HistoriaClinicaProcessor()
            
            # Limpiar variables de archivos
            self.control_files_var.set("Ningún archivo seleccionado")
            self.hospital_files_var.set("Ningún archivo seleccionado")
            
            # Deshabilitar botones
            self.process_btn.config(state='disabled')
            self.open_results_btn.config(state='disabled')
            
            # Limpiar log
            self.log_text.config(state=tk.NORMAL)
            self.log_text.delete('1.0', tk.END)
            self.log_text.config(state=tk.DISABLED)
            
            # Reiniciar progreso
            self.progress_var.set("Listo para procesar")
            self.progress_bar.stop()
            
            # Log inicial
            self.log_message("=== PROCESADOR DE HISTORIAS CLÍNICAS ===")
            self.log_message("1. Seleccione los archivos de control (varios meses)")
            self.log_message("2. Seleccione los archivos del hospital") 
            self.log_message("3. Haga clic en 'Procesar Archivos'")
            self.log_message("")
            self.log_message("Todo limpiado. Listo para comenzar nuevo procesamiento.")

def main():
    """
    Función principal para la interfaz gráfica
    """
    try:
        root = tk.Tk()
        app = HistoriaClinicaGUI(root)
        
        # Centrar ventana
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f'{width}x{height}+{x}+{y}')
        
        # Manejar cierre de ventana
        def on_closing():
            if messagebox.askokcancel("Salir", "¿Desea salir del programa?"):
                root.destroy()
        
        root.protocol("WM_DELETE_WINDOW", on_closing)
        
        # Ejecutar la aplicación
        root.mainloop()
        
    except Exception as e:
        messagebox.showerror("Error Fatal", f"Error inesperado: {str(e)}")

if __name__ == "__main__":
    main()