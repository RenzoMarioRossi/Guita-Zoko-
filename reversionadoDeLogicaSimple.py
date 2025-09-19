import pandas as pd
import os
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import sys

class HistoriaClinicaProcessor:
    def __init__(self):
        self.archivo_control = None
        self.archivos_hospital = []
        self.df_presentes = None
        self.archivo_salida = "presentes_no_pagados.xlsx"
        
    def normalizar_hc(self, valor):
        """
        Normaliza los valores de historia clínica para comparación
        Maneja casos especiales como HC=0 y pacientes sin HC
        """
        if pd.isna(valor):
            return None
        
        # Convertir a string y limpiar
        valor_str = str(valor).strip().lower()
        
        # Casos especiales para pacientes sin HC
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
                           'HC_PACIENTE', 'hc_paciente', 'numero_historia', 'Numero_Historia']
        
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
    
    def seleccionar_archivo_control(self):
        """
        Permite al usuario seleccionar el archivo de control
        """
        root = tk.Tk()
        root.withdraw()
        
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo de control (Excel del usuario)",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if archivo:
            self.archivo_control = archivo
            print(f"Archivo de control seleccionado: {os.path.basename(archivo)}")
            return True
        return False
    
    def seleccionar_archivos_hospital(self):
        """
        Permite al usuario seleccionar múltiples archivos del hospital
        """
        root = tk.Tk()
        root.withdraw()
        
        archivos = filedialog.askopenfilenames(
            title="Seleccionar archivos del hospital (Excel)",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if archivos:
            self.archivos_hospital = list(archivos)
            print(f"Archivos del hospital seleccionados: {len(archivos)}")
            for archivo in archivos:
                print(f"  - {os.path.basename(archivo)}")
            return True
        return False
    
    def procesar_archivo_control(self):
        """
        Procesa el archivo de control y extrae solo los presentes
        Cada fila representa una visita independiente que debe ser pagada
        """
        try:
            print(f"\nProcesando archivo de control...")
            df = pd.read_excel(self.archivo_control)
            
            print(f"Columnas encontradas en archivo de control: {list(df.columns)}")
            
            # Encontrar columnas relevantes
            col_hc = self.encontrar_columna_hc(df)
            col_estado = self.encontrar_columna_estado(df)
            
            if not col_hc:
                raise ValueError("No se pudo encontrar la columna de Historia Clínica")
            if not col_estado:
                raise ValueError("No se pudo encontrar la columna de Estado")
            
            print(f"Columna Historia Clínica: {col_hc}")
            print(f"Columna Estado: {col_estado}")
            
            # Filtrar solo los presentes
            df_filtrado = df[df[col_estado].str.upper() == 'P'].copy()
            
            # Normalizar HC para comparación posterior
            df_filtrado['HC_NORMALIZADA'] = df_filtrado[col_hc].apply(self.normalizar_hc)
            
            # Crear identificador único por fila para manejar múltiples visitas
            df_filtrado['ID_FILA'] = df_filtrado.index.astype(str) + '_' + df_filtrado['HC_NORMALIZADA'].astype(str)
            
            # Eliminar filas donde no se pudo normalizar la HC
            antes_filtro = len(df_filtrado)
            df_filtrado = df_filtrado.dropna(subset=['HC_NORMALIZADA'])
            despues_filtro = len(df_filtrado)
            
            if antes_filtro != despues_filtro:
                print(f"  Advertencia: {antes_filtro - despues_filtro} filas eliminadas por HC inválida")
            
            self.df_presentes = df_filtrado
            
            print(f"Total de registros en archivo de control: {len(df)}")
            print(f"Registros con estado 'Presente': {len(df_filtrado)}")
            
            # Mostrar estadísticas de HC especiales
            hc_especiales = df_filtrado[df_filtrado['HC_NORMALIZADA'].astype(str).str.contains('SIN_HC|HC_0|ESPECIAL', na=False)]
            if len(hc_especiales) > 0:
                print(f"  HC especiales encontradas (HC=0, sin HC, etc.): {len(hc_especiales)}")
            
            return True
            
        except Exception as e:
            print(f"Error procesando archivo de control: {str(e)}")
            return False
    
    def procesar_archivos_hospital(self):
        """
        Procesa los archivos del hospital y elimina las HC que coinciden
        Maneja múltiples visitas del mismo paciente como filas independientes
        """
        if self.df_presentes is None:
            print("Error: No hay datos de presentes para procesar")
            return False
        
        hc_presentes = set(self.df_presentes['HC_NORMALIZADA'].tolist())
        ids_presentes = set(self.df_presentes['ID_FILA'].tolist())
        ids_encontrados_hospital = set()
        
        print(f"\nHistorias clínicas únicas presentes: {len(hc_presentes)}")
        print(f"Total de filas/visitas presentes: {len(ids_presentes)}")
        
        for archivo_hospital in self.archivos_hospital:
            try:
                print(f"\nProcesando archivo del hospital: {os.path.basename(archivo_hospital)}")
                df_hospital = pd.read_excel(archivo_hospital)
                
                print(f"Columnas en archivo del hospital: {list(df_hospital.columns)}")
                
                # Encontrar columna HC en archivo del hospital
                col_hc_hospital = self.encontrar_columna_hc(df_hospital)
                
                if not col_hc_hospital:
                    print(f"  No se encontró columna HC en {os.path.basename(archivo_hospital)}")
                    continue
                
                print(f"  Columna Historia Clínica encontrada: {col_hc_hospital}")
                
                # Normalizar HC del hospital
                hc_hospital = []
                for hc in df_hospital[col_hc_hospital]:
                    hc_norm = self.normalizar_hc(hc)
                    if hc_norm is not None:
                        hc_hospital.append(hc_norm)
                
                print(f"  HC válidas encontradas en este archivo: {len(hc_hospital)}")
                
                # Para cada HC en el hospital, buscar coincidencias en presentes
                # y marcar las filas como encontradas (una por una)
                for hc_hosp in hc_hospital:
                    # Buscar filas en presentes que tengan esta HC y aún no estén marcadas como encontradas
                    filas_candidatas = self.df_presentes[
                        (self.df_presentes['HC_NORMALIZADA'] == hc_hosp) &
                        (~self.df_presentes['ID_FILA'].isin(ids_encontrados_hospital))
                    ]
                    
                    if len(filas_candidatas) > 0:
                        # Marcar la primera fila encontrada como pagada
                        id_fila = filas_candidatas.iloc[0]['ID_FILA']
                        ids_encontrados_hospital.add(id_fila)
                
                coincidencias_este_archivo = len([id for id in ids_encontrados_hospital 
                                                if id.split('_')[1:] == [hc for hc in hc_hospital if str(hc) in id]])
                print(f"  Filas marcadas como pagadas en este archivo: {len(ids_encontrados_hospital) - len(ids_encontrados_hospital.intersection(set()))}") 
                
            except Exception as e:
                print(f"Error procesando {os.path.basename(archivo_hospital)}: {str(e)}")
        
        # Eliminar las filas encontradas en el hospital del DataFrame de presentes
        ids_no_pagados = ids_presentes - ids_encontrados_hospital
        
        # Filtrar el DataFrame
        self.df_presentes = self.df_presentes[
            self.df_presentes['ID_FILA'].isin(ids_no_pagados)
        ].copy()
        
        print(f"\nResumen:")
        print(f"Total de filas/visitas presentes: {len(ids_presentes)}")
        print(f"Filas encontradas como pagadas en hospital: {len(ids_encontrados_hospital)}")
        print(f"Filas NO PAGADAS (discrepancias): {len(ids_no_pagados)}")
        
        # Mostrar estadísticas por paciente
        if len(self.df_presentes) > 0:
            pacientes_no_pagados = self.df_presentes['HC_NORMALIZADA'].nunique()
            print(f"Pacientes únicos con visitas no pagadas: {pacientes_no_pagados}")
        
        return True
    
    def guardar_resultados(self):
        """
        Guarda los resultados finales en un archivo Excel
        """
        if self.df_presentes is None or self.df_presentes.empty:
            print("\n✅ ¡Excelente! No hay discrepancias.")
            print("Todas las visitas presentes fueron pagadas correctamente.")
            return True
        
        try:
            # Eliminar las columnas auxiliares antes de guardar
            df_salida = self.df_presentes.drop(['HC_NORMALIZADA', 'ID_FILA'], axis=1)
            
            # Ordenar por HC para mejor visualización
            col_hc_original = self.encontrar_columna_hc(df_salida)
            if col_hc_original:
                df_salida = df_salida.sort_values(by=col_hc_original)
            
            df_salida.to_excel(self.archivo_salida, index=False)
            print(f"\n⚠️  Resultados guardados en: {self.archivo_salida}")
            print(f"📊 Total de visitas no pagadas: {len(df_salida)}")
            
            # Estadísticas adicionales
            if col_hc_original:
                pacientes_unicos = df_salida[col_hc_original].nunique()
                print(f"👥 Pacientes únicos afectados: {pacientes_unicos}")
                
                # Mostrar pacientes con múltiples visitas no pagadas
                visitas_por_paciente = df_salida[col_hc_original].value_counts()
                multiples_visitas = visitas_por_paciente[visitas_por_paciente > 1]
                if len(multiples_visitas) > 0:
                    print(f"🔄 Pacientes con múltiples visitas no pagadas: {len(multiples_visitas)}")
            
            return True
            
        except Exception as e:
            print(f"Error guardando resultados: {str(e)}")
            return False
    
    def abrir_resultados(self):
        """
        Abre el archivo de resultados
        """
        if not os.path.exists(self.archivo_salida):
            print("No existe archivo de resultados para abrir")
            return
        
        try:
            if sys.platform == "win32":
                os.startfile(self.archivo_salida)
            elif sys.platform == "darwin":
                subprocess.run(["open", self.archivo_salida])
            else:
                subprocess.run(["xdg-open", self.archivo_salida])
            
            print(f"Abriendo archivo de resultados: {self.archivo_salida}")
            
        except Exception as e:
            print(f"No se pudo abrir automáticamente el archivo: {str(e)}")
            print(f"Por favor, abra manualmente: {self.archivo_salida}")
    
    def ejecutar(self):
        """
        Ejecuta el proceso completo
        """
        print("=== PROCESADOR DE HISTORIAS CLÍNICAS ===\n")
        
        # Paso 1: Seleccionar archivo de control
        print("Paso 1: Seleccionar archivo de control del usuario")
        if not self.seleccionar_archivo_control():
            print("Proceso cancelado - No se seleccionó archivo de control")
            return
        
        # Paso 2: Seleccionar archivos del hospital
        print("\nPaso 2: Seleccionar archivos del hospital")
        if not self.seleccionar_archivos_hospital():
            print("Proceso cancelado - No se seleccionaron archivos del hospital")
            return
        
        # Paso 3: Procesar archivo de control
        print("\nPaso 3: Procesando archivo de control...")
        if not self.procesar_archivo_control():
            print("Error en el procesamiento del archivo de control")
            return
        
        # Paso 4: Procesar archivos del hospital
        print("\nPaso 4: Procesando archivos del hospital...")
        if not self.procesar_archivos_hospital():
            print("Error en el procesamiento de archivos del hospital")
            return
        
        # Paso 5: Guardar resultados
        print("\nPaso 5: Guardando resultados...")
        if not self.guardar_resultados():
            print("Error guardando resultados")
            return
        
        # Paso 6: Abrir resultados
        print("\nPaso 6: Abriendo resultados...")
        self.abrir_resultados()
        
        print("\n=== PROCESO COMPLETADO ===")
        if self.df_presentes is not None and len(self.df_presentes) > 0:
            print(f"❌ Discrepancias encontradas: {len(self.df_presentes)} visitas")
            print("Estas son las visitas de pacientes presentes que NO fueron pagadas.")
            print("Cada fila representa una visita independiente que requiere pago.")
        else:
            print("✅ ¡Perfecto! No se encontraron discrepancias.")
            print("Todas las visitas presentes han sido correctamente pagadas.")
            
        print(f"\nArchivo de resultados: {self.archivo_salida}")
        print("\n📋 IMPORTANTE:")
        print("• Cada FILA = Una VISITA que debe ser pagada")  
        print("• Si un paciente vino 3 veces, debe aparecer 3 veces en archivos del hospital")
        print("• HC=0 y 'sin HC' también se consideran válidos para cobro")
        print("• Las discrepancias mostradas requieren verificación/pago")

def main():
    """
    Función principal
    """
    try:
        processor = HistoriaClinicaProcessor()
        processor.ejecutar()
    except KeyboardInterrupt:
        print("\nProceso interrumpido por el usuario")
    except Exception as e:
        print(f"Error inesperado: {str(e)}")
        input("Presione Enter para salir...")

if __name__ == "__main__":
    main()