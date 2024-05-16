# process.py
import os
import pandas as pd
import openpyxl
from pathlib import Path
import re

# Definir la función para procesar un archivo Excel
def procesar_excel(file_path, sheet_name, data_certificados):
    try:
        # Leer el archivo Excel, especificando que los nombres de columnas están en la fila 14 (índice 13)
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=14)

        # Eliminar las primeras dos filas no deseadas
        df = df.drop([0, 1])

        # Restablecer el índice
        df.reset_index(drop=True, inplace=True)

        # Renombrar las columnas correctamente
        df.columns = ["Método de Análisis", "Parámetro", "CM", "Unidad", "LD", "LQ"] + df.columns[6:].tolist()

        # Transformar al formato largo
        id_vars = ["Método de Análisis", "Parámetro", "CM", "Unidad", "LD", "LQ"]
        melted_data = pd.melt(df, id_vars=id_vars, var_name='Sample_ID', value_name='Valor')
        melted_data['Ruta'] = file_path

        # Extraer el segmento de la ruta Proyecto
        path = Path(file_path)
        path_parts = path.parts
        start_segment = "Certificados"
        try:
            start_index = path_parts.index(start_segment) + 1
        except ValueError:
            return "Segmento no encontrado en la ruta."

        extracted_path = path_parts[start_index]
        cleaned_path = re.sub(r'^\d+\.\s*', '', extracted_path)
        melted_data['Proyecto'] = cleaned_path

        # Obtener Campaña
        try:
            start_ = path_parts.index(start_segment) + 2
        except ValueError:
            return "Segmento no encontrado en la ruta."

        extracted_campaña = path_parts[start_]
        campaña = re.sub(r'^\d+\.\s*', ' ', extracted_campaña)
        melted_data['Campaña'] = campaña

        # Cargamos nuevamente los datos originales de la hoja de resultados
        original_data = pd.read_excel(file_path, sheet_name=sheet_name)
        identification_data = original_data.iloc[9:14, 1:].transpose()
        identification_data.columns = identification_data.iloc[0]
        identification_data = identification_data.drop(identification_data.index[0])
        identification_data.reset_index(drop=True, inplace=True)
        identification_data.columns = ["N° Informe_LB", "Fecha de Muestreo", "Hora de Muestreo", "Tipo de Muestra", "FLOW"]

        # Unimos los datos de identificación con los datos fundidos basándonos en el índice
        combined_data = pd.merge(melted_data, identification_data, left_on='Sample_ID', right_on='FLOW')
        combined_data.drop('Sample_ID', axis=1, inplace=True)
        combined_data.rename(columns={'FLOW': 'Sample_ID'}, inplace=True)

        # Reorganizando las columnas para que 'Sample_ID', 'Fecha de Muestreo', y 'Tipo de Muestra' aparezcan primero
        column_order = ['Sample_ID', 'Fecha de Muestreo', "Hora de Muestreo", 'Tipo de Muestra', 'Método de Análisis', 'Parámetro','Proyecto','Campaña','Ruta', 
                        'CM', 'Unidad', 'LD', 'LQ', 'N° Informe_LB', 'Valor']
        combined_data = combined_data[column_order]

        # LIMPIEZA DE DATA
        combined_data_cleaned = combined_data[combined_data['Parámetro'] != 'Fecha de Análisis']
        try:
            combined_data_cleaned['Fecha de Muestreo'] = pd.to_datetime(combined_data_cleaned['Fecha de Muestreo'], dayfirst=True)
        except Exception as e:
            print(f"Error ajustando el formato de la fecha: {e}")

        # Añadir los datos procesados al DataFrame global
        data_certificados = pd.concat([data_certificados, combined_data_cleaned], ignore_index=True)

        return data_certificados

    except Exception as e:
        print(f"Error procesando el archivo: {e}")
        return data_certificados

def process_excel_files(directory_path):
    data_certificados = pd.DataFrame()
    excel_files = [file for file in os.listdir(directory_path) if file.endswith('.xlsx')]

    for file in excel_files:
        file_path = os.path.join(directory_path, file)
        xls = pd.ExcelFile(file_path)

        for sheet_name in xls.sheet_names:
            data_certificados = procesar_excel(file_path, sheet_name, data_certificados)

    return data_certificados
