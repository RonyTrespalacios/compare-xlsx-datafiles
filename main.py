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
st.title("📊 Compara tus contactos con la base de datos de habilitados a votar!!")

# Sección: Cargar archivo de contactos
st.header("⬇️ Subir tu lista de contactos")
contactos_file = st.file_uploader("Por favor, sube tu archivo de contactos (puede ser en formato CSV o Excel)", type=["csv", "xlsx"], key="contactos_upload")

# Botones de selección única para elegir la base de datos a comparar
st.header("Base de datos para comparar ⚙️")
db_option = st.radio(
    "Selecciona la base de datos con la que deseas comparar tus contactos:",
    options=[
        ("consulta_egresados.xlsx", "Censo para elección representante de egresados"),
        ("consulta_rector.xlsx", "Censo de egresados para la consulta rectoral")
    ],
    index=0,  # Establece "Censo de egresados para la consulta rectoral" como predeterminada
    format_func=lambda x: x[1]  # Muestra solo el label en la interfaz
)

# Extrae el nombre del archivo seleccionado
db_file = db_option[0]

# Entrada de texto para el nombre del archivo de salida
output_filename = st.text_input("💾 Ingrese el nombre del archivo de salida (sin extensión)", value="resultado")

# Barra de progreso
progress_bar = st.progress(0)

# Botón para procesar archivos
if st.button("Click para comparar!! 👇"):
    if contactos_file:
        # Mostrar indicador de carga mientras se procesa el archivo CSV
        with st.spinner('Procesando tu archivo de contactos...'):
            # Si el archivo de contactos es CSV, convertirlo a Excel
            if contactos_file.name.endswith('.csv'):
                contactos_data = csv_to_xlsx(contactos_file)
            else:
                contactos_data = BytesIO(contactos_file.getvalue())

        # Mostrar indicador de carga mientras se carga la base de datos local
        with st.spinner('Cargando la base de datos seleccionada...'):
            try:
                with open(db_file, 'rb') as f:
                    egresados_data = BytesIO(f.read())
            except Exception as e:
                st.error(f"⚠️ Ocurrió un error al cargar la base de datos: {str(e)}")
                egresados_data = None

        if egresados_data:
            # Crear un flujo de salida en memoria para el resultado
            output_data = BytesIO()

            # Ejecutar la función de procesamiento
            try:
                with st.spinner('Comparando y generando el archivo final...'):
                    generar_archivo_combinado(contactos_data, egresados_data, output_data, progress_bar)
                    output_data.seek(0)  # Restablecer el cursor al principio del stream

                st.success("✅ Comparación completada con éxito!")
                st.download_button(
                    label="⬇️ Descargar archivo procesado",
                    data=output_data,
                    file_name=f"{output_filename}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"⚠️ Ocurrió un error: {str(e)}")
    else:
        st.error("⚠️ Por favor, sube un archivo de contactos.")