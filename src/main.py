import streamlit as st
import pandas as pd
from io import BytesIO
from processing import generar_archivo_combinado

def csv_to_xlsx(csv_file):
    # Leer el archivo CSV con codificación UTF-8
    df = pd.read_csv(csv_file, encoding='utf-8')
    
    # Guardar el DataFrame en un archivo Excel en memoria
    output = BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)  # Mover el cursor al principio del stream
    return output

# Título y descripción de la aplicación
st.title("📊 Data Processing App")

# Sección 1: Convertidor de CSV a Excel
st.header("🔄 CSV 2 Excel Converter")

# Cargador de archivos CSV
uploaded_csv = st.file_uploader("📁 Choose a CSV file", type="csv", key="csv_upload")

# Entrada de texto para el nombre del archivo de salida
output_filename_csv = st.text_input("💾 Enter Output File Name (without extension)", value="contacts")

if uploaded_csv is not None:
    # Convertir CSV a Excel en memoria
    excel_data = csv_to_xlsx(uploaded_csv)
    
    st.download_button(
        label="⬇️ Download Excel file",
        data=excel_data,
        file_name=f"{output_filename_csv}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Sección 2: Procesar Archivos Excel
st.header("🔍 Compare Excel Files (Contactos and Egresados)")

# Título y descripción
st.write("📑 This application processes Excel files by matching and sorting data.")

# Cargador de archivos para Contactos y Egresados
contactos_file = st.file_uploader("📂 Upload Contactos File", type=["xlsx"], key="contactos_upload")
egresados_file = st.file_uploader("📂 Upload Egresados File", type=["xlsx"], key="egresados_upload")

# Entrada de texto para el nombre del archivo de salida
output_filename = st.text_input("💾 Enter Output File Name (without extension)", value="result")

# Barra de progreso
progress_bar = st.progress(0)

# Botón para procesar archivos
if st.button("⚙️ Match and Sort Files"):
    if contactos_file and egresados_file:
        # Crear flujos binarios en memoria para los archivos subidos
        contactos_data = BytesIO(contactos_file.getvalue())
        egresados_data = BytesIO(egresados_file.getvalue())

        # Crear un flujo de salida en memoria para el resultado
        output_data = BytesIO()

        # Ejecutar la función de procesamiento
        try:
            generar_archivo_combinado(contactos_data, egresados_data, output_data, progress_bar)
            output_data.seek(0)  # Restablecer el cursor al principio del stream

            st.download_button(
                label="⬇️ Download Processed File",
                data=output_data,
                file_name=f"{output_filename}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"⚠️ An error occurred: {str(e)}")
    else:
        st.error("⚠️ Please upload both Contactos and Egresados files.")
