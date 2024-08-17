import pandas as pd
from difflib import SequenceMatcher
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
from io import BytesIO


def preparar_dataframes(contactos_data, egresados_data):
    contactos_df = pd.read_excel(contactos_data)
    egresados_df = pd.read_excel(egresados_data)
    contactos_df['Nombre'] = contactos_df[['First Name', 'Middle Name', 'Last Name']].fillna('').agg(' '.join, axis=1).str.strip()
    return contactos_df, egresados_df

def remove_accents(input_str):
    replacements = (
        ("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u"),
        ("Á", "A"), ("É", "E"), ("Í", "I"), ("Ó", "O"), ("Ú", "U")
    )
    for a, b in replacements:
        input_str = input_str.replace(a, b)
    return input_str

def normalize_name(name):
    name = str(name).strip()
    name = remove_accents(name)
    name = name.upper()
    return ' '.join(name.split())

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

def calcular_coincidencias(nombre1, nombre2):
    palabras1 = set(nombre1.split())
    palabras2 = set(nombre2.split())
    coincidencias = palabras1.intersection(palabras2)
    return len(coincidencias)

def ajustar_filas_y_columnas(worksheet):
    for column in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except Exception as e:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column_letter].width = adjusted_width

    for row in worksheet.iter_rows():
        max_height = 0
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical='top')
            if cell.value:
                cell_text = str(cell.value)
                cell_height = cell_text.count('\n') + 1
                if cell_height > max_height:
                    max_height = cell_height
        worksheet.row_dimensions[cell.row].height = max_height * 15  # Approximate height per line

def generar_archivo_combinado(contactos_data, egresados_data, output_stream, progress_bar):
    contactos_df, egresados_df = preparar_dataframes(contactos_data, egresados_data)
    contactos_df['Nombre'] = contactos_df['Nombre'].apply(normalize_name)
    egresados_df['Nombres'] = egresados_df['Nombres'].apply(normalize_name)

    # Verificar las columnas de teléfono
    telefono1_col = 'Phone 1 - Value' if 'Phone 1 - Value' in contactos_df.columns else 'Mobile Phone'
    telefono2_col = 'Phone 2 - Value' if 'Phone 2 - Value' in contactos_df.columns else ('Primary Phone' if 'Other Phone' in contactos_df.columns else 'Primary Phone')

    keyword_index = defaultdict(list)
    for index, egresado in egresados_df.iterrows():
        words = set(egresado['Nombres'].split())
        for word in words:
            keyword_index[word].append((index, egresado['Nombres']))

    resultados = []

    total_rows = len(contactos_df)
    for idx, contacto in enumerate(contactos_df.iterrows()):
        index, contacto = contacto
        nombre_contacto = contacto['Nombre']
        palabras_contacto = set(nombre_contacto.split())
        posibles_candidatos = set()
        for palabra in palabras_contacto:
            if palabra in keyword_index:
                posibles_candidatos.update(keyword_index[palabra])
        mejor_match = None
        mejor_similitud = 0
        mejor_coincidencias = 0
        for candidato in posibles_candidatos:
            index2, nombre_egresado = candidato
            similitud = similar(nombre_contacto, nombre_egresado)
            coincidencias = calcular_coincidencias(nombre_contacto, nombre_egresado)
            if coincidencias > mejor_coincidencias or (coincidencias == mejor_coincidencias and similitud > mejor_similitud):
                mejor_coincidencias = coincidencias
                mejor_similitud = similitud
                mejor_match = egresados_df.loc[index2]
        if mejor_match is not None:
            resultados.append({
                'Cedula': mejor_match['Cedula'],
                'Nombre': nombre_contacto,
                'Telefono1': contacto[telefono1_col],
                'Telefono2': contacto[telefono2_col],
                'Nombre Egresado': mejor_match['Nombres'],
                'Certeza': mejor_similitud * 100,
                'Coincidencias': mejor_coincidencias
            })

        # Update progress
        progress_bar.progress(int((idx + 1) / total_rows * 100))

    # Convert the results to a DataFrame
    resultados_df = pd.DataFrame(resultados)

    # Sort by Coincidencias and then by Certeza, both descending
    resultados_df.sort_values(by=['Coincidencias', 'Certeza'], ascending=[False, False], inplace=True)

    # Create a new workbook
    workbook = Workbook()
    worksheet = workbook.active

    # Write DataFrame to worksheet
    for r_idx, row in enumerate(dataframe_to_rows(resultados_df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            worksheet.cell(row=r_idx, column=c_idx, value=value)

    # Adjust rows and columns
    ajustar_filas_y_columnas(worksheet)

    # Save to the output stream
    workbook.save(output_stream)
    output_stream.seek(0)  # Reset stream position