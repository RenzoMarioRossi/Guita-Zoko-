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
        self.archivos_control = []  # Ahora múltiples archivos
        self.archivos_hospital = []
        self.df_presentes = None
        self.df_pagos_en_contra = None  # Nueva: pagos que no corresponden
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
    
    def seleccionar_archivos_control(self):
        """
        Permite al usuario seleccionar múltiples archivos de control (varios meses)
        """
        root = tk.Tk()
        root.withdraw()
        
        archivos = filedialog.askopenfilenames(
            title="Seleccionar archivos de control (Excel del usuario - varios meses)",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if archivos:
            self.archivos_control = list(archivos)
            print(f"Archivos de control seleccionados: {len(archivos)}")
            for archivo in archivos:
                print(f"  - {os.path.basename(archivo)}")
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
    
    def procesar_archivos_control(self):
        """
        Procesa múltiples archivos de control y extrae solo los presentes
        Cada fila representa una visita independiente que debe ser pagada
        """
        try:
            print(f"\nProcesando {len(self.archivos_control)} archivos de control...")
            df_todos_presentes = []
            
            for archivo in self.archivos_control:
                print(f"\n  Procesando: {os.path.basename(archivo)}")
                df = pd.read_excel(archivo)
                
                # Encontrar columnas relevantes
                col_hc = self.encontrar_columna_hc(df)
                col_estado = self.encontrar_columna_estado(df)
                
                if not col_hc:
                    print(f"    ❌ No se encontró columna HC en {os.path.basename(archivo)}")
                    continue
                if not col_estado:
                    print(f"    ❌ No se encontró columna Estado en {os.path.basename(archivo)}")
                    continue
                
                # Filtrar solo los presentes
                df_filtrado = df[df[col_estado].str.upper() == 'P'].copy()
                
                # Normalizar HC para comparación posterior
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
                    print(f"    ⚠️  {antes_filtro - despues_filtro} filas eliminadas por HC inválida")
                
                df_todos_presentes.append(df_filtrado)
                print(f"    ✅ {len(df_filtrado)} registros presentes")
            
            # Combinar todos los archivos
            if df_todos_presentes:
                self.df_presentes = pd.concat(df_todos_presentes, ignore_index=True)
                
                total_presentes = len(self.df_presentes)
                archivos_procesados = len([df for df in df_todos_presentes if len(df) > 0])
                
                print(f"\n📊 Resumen archivos de control:")
                print(f"  Archivos procesados exitosamente: {archivos_procesados}")
                print(f"  Total registros presentes: {total_presentes}")
                
                # Mostrar estadísticas de HC especiales
                hc_especiales = self.df_presentes[self.df_presentes['HC_NORMALIZADA'].astype(str).str.contains('SIN_HC|HC_0|ESPECIAL', na=False)]
                if len(hc_especiales) > 0:
                    print(f"  HC especiales (HC=0, sin HC, etc.): {len(hc_especiales)}")
                
                return True
            else:
                print("❌ No se pudieron procesar archivos de control")
                return False
                
        except Exception as e:
            print(f"Error procesando archivos de control: {str(e)}")
            return False
    
    def procesar_archivos_hospital(self):
        """
        Procesa los archivos del hospital y elimina las HC que coinciden
        También identifica pagos "en contra" (en hospital pero no en control)
        """
        if self.df_presentes is None:
            print("Error: No hay datos de presentes para procesar")
            return False
        
        hc_presentes = set(self.df_presentes['HC_NORMALIZADA'].tolist())
        ids_presentes = set(self.df_presentes['ID_FILA'].tolist())
        ids_encontrados_hospital = set()
        
        # Para detectar pagos "en contra"
        todas_hc_hospital = []
        
        print(f"\nHistorias clínicas únicas presentes: {len(hc_presentes)}")
        print(f"Total de filas/visitas presentes: {len(ids_presentes)}")
        
        for archivo_hospital in self.archivos_hospital:
            try:
                print(f"\nProcesando archivo del hospital: {os.path.basename(archivo_hospital)}")
                df_hospital = pd.read_excel(archivo_hospital)
                
                # Encontrar columna HC en archivo del hospital
                col_hc_hospital = self.encontrar_columna_hc(df_hospital)
                
                if not col_hc_hospital:
                    print(f"  No se encontró columna HC en {os.path.basename(archivo_hospital)}")
                    continue
                
                print(f"  Columna Historia Clínica encontrada: {col_hc_hospital}")
                
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
                
                print(f"  HC válidas encontradas en este archivo: {len([h for h in todas_hc_hospital if h['ARCHIVO_HOSPITAL'] == os.path.basename(archivo_hospital)])}")
                
            except Exception as e:
                print(f"Error procesando {os.path.basename(archivo_hospital)}: {str(e)}")
        
        # Procesar pagos "en contra" (están en hospital pero no en control)
        if todas_hc_hospital:
            df_hospital_completo = pd.DataFrame(todas_hc_hospital)
            hc_solo_hospital = df_hospital_completo[~df_hospital_completo['HC_NORMALIZADA'].isin(hc_presentes)]
            
            if len(hc_solo_hospital) > 0:
                self.df_pagos_en_contra = hc_solo_hospital.copy()
                print(f"\n🔍 Pagos 'en contra' encontrados: {len(hc_solo_hospital)} filas")
            else:
                print(f"\n✅ No se encontraron pagos 'en contra'")
        
        # Eliminar las filas encontradas en el hospital del DataFrame de presentes
        ids_no_pagados = ids_presentes - ids_encontrados_hospital
        
        # Filtrar el DataFrame
        self.df_presentes = self.df_presentes[
            self.df_presentes['ID_FILA'].isin(ids_no_pagados)
        ].copy()
        
        print(f"\nResumen:")
        print(f"Total de filas/visitas presentes: {len(ids_presentes)}")
        print(f"Filas encontradas como pagadas en hospital: {len(ids_encontrados_hospital)}")
        print(f"Filas NO PAGADAS (discrepancias a favor): {len(ids_no_pagados)}")
        
        # Mostrar estadísticas por paciente
        if len(self.df_presentes) > 0:
            pacientes_no_pagados = self.df_presentes['HC_NORMALIZADA'].nunique()
            print(f"Pacientes únicos con visitas no pagadas: {pacientes_no_pagados}")
        
        return True
    
    def guardar_resultados(self):
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
                print(f"\n💰 Discrepancias A FAVOR guardadas: {self.archivo_salida}")
                print(f"   Total visitas no pagadas: {len(df_salida)}")
                resultados_guardados += 1
                
            except Exception as e:
                print(f"Error guardando discrepancias a favor: {str(e)}")
        else:
            print(f"\n✅ Sin discrepancias A FAVOR - Todas las visitas fueron pagadas")
        
        # Guardar discrepancias "en contra" (pagos sin correspondencia)
        if self.df_pagos_en_contra is not None and not self.df_pagos_en_contra.empty:
            try:
                # Eliminar la columna auxiliar antes de guardar
                df_contra = self.df_pagos_en_contra.drop(['HC_NORMALIZADA'], axis=1)
                
                # Ordenar por archivo hospital
                df_contra = df_contra.sort_values(by='ARCHIVO_HOSPITAL')
                
                df_contra.to_excel(self.archivo_salida_contra, index=False)
                print(f"\n⚠️  Discrepancias EN CONTRA guardadas: {self.archivo_salida_contra}")
                print(f"   Total pagos sin correspondencia: {len(df_contra)}")
                resultados_guardados += 1
                
            except Exception as e:
                print(f"Error guardando discrepancias en contra: {str(e)}")
        else:
            print(f"\n✅ Sin discrepancias EN CONTRA - Todos los pagos corresponden")
        
        return resultados_guardados > 0
    
    def abrir_resultados(self):
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
            print("No hay archivos de resultados para abrir")
            return
        
        for archivo in archivos_a_abrir:
            try:
                if sys.platform == "win32":
                    os.startfile(archivo)
                elif sys.platform == "darwin":
                    subprocess.run(["open", archivo])
                else:
                    subprocess.run(["xdg-open", archivo])
                
                print(f"Abriendo: {archivo}")
                
            except Exception as e:
                print(f"No se pudo abrir automáticamente {archivo}: {str(e)}")
                print(f"Por favor, abra manualmente: {archivo}")
    
    def ejecutar(self):
        """
        Ejecuta el proceso completo
        """
        print("=== PROCESADOR DE HISTORIAS CLÍNICAS ===\n")
        
        # Paso 1: Seleccionar archivos de control
        print("Paso 1: Seleccionar archivos de control del usuario (varios meses)")
        if not self.seleccionar_archivos_control():
            print("Proceso cancelado - No se seleccionaron archivos de control")
            return
        
        # Paso 2: Seleccionar archivos del hospital
        print("\nPaso 2: Seleccionar archivos del hospital")
        if not self.seleccionar_archivos_hospital():
            print("Proceso cancelado - No se seleccionaron archivos del hospital")
            return
        
        # Paso 3: Procesar archivos de control
        print("\nPaso 3: Procesando archivos de control...")
        if not self.procesar_archivos_control():
            print("Error en el procesamiento de archivos de control")
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
        
        # Resultados A FAVOR
        if self.df_presentes is not None and len(self.df_presentes) > 0:
            print(f"💰 Discrepancias A FAVOR: {len(self.df_presentes)} visitas")
            print("   (Pacientes presentes que NO fueron pagados por el hospital)")
        else:
            print("✅ Sin discrepancias A FAVOR - Todas las visitas fueron pagadas")
        
        # Resultados EN CONTRA
        if self.df_pagos_en_contra is not None and len(self.df_pagos_en_contra) > 0:
            print(f"⚠️  Discrepancias EN CONTRA: {len(self.df_pagos_en_contra)} pagos")
            print("   (Pagos del hospital que NO corresponden a tus archivos de control)")
        else:
            print("✅ Sin discrepancias EN CONTRA - Todos los pagos corresponden")
            
        print(f"\nArchivos generados:")
        print(f"• A favor: {self.archivo_salida}")
        print(f"• En contra: {self.archivo_salida_contra}")
        
        print("\n📋 IMPORTANTE:")
        print("• A FAVOR = El hospital te debe dinero")  
        print("• EN CONTRA = El hospital te pagó algo que no corresponde a estos períodos")
        print("• Cada fila representa una visita/pago independiente")
        print("• Revisa ambos archivos para cuadrar las cuentas completamente")
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