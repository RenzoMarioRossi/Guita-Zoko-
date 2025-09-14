import pandas as pd
import numpy as np

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

                                                                        def compare_records(user_file, hospital_files):
                                                                            """Compara tu registro con los del hospital y genera un Excel con discrepancias."""
                                                                                # Leer tu registro
        def compare_records(user_file, hospital_files):
                """Compara tu registro con los del hospital y genera un Excel con discrepancias."""
                # Leer tu registro
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
                                'apellido_nombre': 'Doctor'  # Ignoramos si es el nombre del doctor
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

                # Extra en tu registro (a favor)
                extra_user_mask = merged['_merge'] == 'left_only'
                extra_user = merged[extra_user_mask]
                if not extra_user.empty:
                        if 'Nombre_user' in extra_user.columns:
                                extra_user['Nombre'] = extra_user['Nombre_user'].combine_first(extra_user['Nombre_hospital'])
                        else:
                                extra_user['Nombre'] = extra_user.get('Nombre', np.nan)
                        if 'Monto_user' in extra_user.columns:
                                extra_user['Monto'] = extra_user['Monto_user'].combine_first(extra_user['Monto_hospital'])
                        else:
                                extra_user['Monto'] = extra_user.get('Monto', 0.0)
                        if 'Hora_user' in extra_user.columns:
                                extra_user['Hora'] = extra_user['Hora_user'].combine_first(extra_user['Hora_hospital'])
                        else:
                                extra_user['Hora'] = extra_user.get('Hora', np.nan)
                        extra_user = extra_user[key_cols + ['Nombre', 'Hora', 'Monto']]

                # Extra en hospital (en contra)
                extra_hospital_mask = merged['_merge'] == 'right_only'
                extra_hospital = merged[extra_hospital_mask]
                if not extra_hospital.empty:
                        if 'Nombre_hospital' in extra_hospital.columns:
                                extra_hospital['Nombre'] = extra_hospital['Nombre_hospital'].combine_first(extra_hospital['Nombre_user'])
                        else:
                                extra_hospital['Nombre'] = extra_hospital.get('Nombre', np.nan)
                        if 'Monto_hospital' in extra_hospital.columns:
                                extra_hospital['Monto'] = extra_hospital['Monto_hospital'].combine_first(extra_hospital['Monto_user'])
                        else:
                                extra_hospital['Monto'] = extra_hospital.get('Monto', 0.0)
                        if 'Hora_hospital' in extra_hospital.columns:
                                extra_hospital['Hora'] = extra_hospital['Hora_hospital'].combine_first(extra_hospital['Hora_user'])
                        else:
                                extra_hospital['Hora'] = extra_hospital.get('Hora', np.nan)
                        extra_hospital = extra_hospital[key_cols + ['Nombre', 'Hora', 'Monto']]

                # Ordenar
                cols = ['Nombre', 'Fecha', 'Hora', 'HC', 'Monto']
                extra_user = extra_user[cols].sort_values(by=['Nombre', 'Fecha'])
                extra_hospital = extra_hospital[cols].sort_values(by=['Nombre', 'Fecha'])

                # Resumen
                summary = pd.DataFrame({
                        'Tipo': ['Extra en mi registro (a favor)', 'Extra en hospital (en contra)'],
                        'Cantidad': [len(extra_user), len(extra_hospital)],
                        'Monto Total': [extra_user['Monto'].sum(), extra_hospital['Monto'].sum()]
                })

                # Exportar
                output_file = 'discrepancias_pacientes.xlsx'
                with pd.ExcelWriter(output_file) as writer:
                        extra_user.to_excel(writer, sheet_name='Extra_Mi_Registro', index=False)
                        extra_hospital.to_excel(writer, sheet_name='Extra_Hospital', index=False)
                        summary.to_excel(writer, sheet_name='Resumen', index=False)

                return output_file

        # Ejemplo de uso
        # archivo_generado = compare_records('tu_registro.xlsx', ['hospital_planes.xlsx', 'hospital_pami.xlsx', 'hospital_ooss.xlsx'])
        # print(f'Archivo generado: {archivo_generado}')
        print('Código listo para usar. Ajusta las rutas de archivos en la llamada a compare_records.')
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                return output_file

                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                # Ejemplo de uso
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                # archivo_generado = compare_records('tu_registro.xlsx', ['hospital_planes.xlsx', 'hospital_pami.xlsx', 'hospital_ooss.xlsx'])
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                # print(f'Archivo generado: {archivo_generado}')

                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                print('Código listo para usar. Ajusta las rutas de archivos en la llamada a compare_records.')</parameter
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                </xai:function_call