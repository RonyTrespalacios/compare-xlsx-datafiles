import streamlit as st
import pandas as pd
from io import BytesIO
from processing import generar_archivo_combinado
import sqlite3

def init_db():
    # Conectarse a la base de datos (o crearla si no existe)
    conn = sqlite3.connect('contador.db')
    c = conn.cursor()
    
    # Crear una tabla para el contador si no existe
    c.execute('''CREATE TABLE IF NOT EXISTS contador (
                 id INTEGER PRIMARY KEY,
                 count INTEGER)''')
    
    # Insertar un registro de contador inicial si no existe
    c.execute('''INSERT OR IGNORE INTO contador (id, count) VALUES (1, 0)''')
    
    # Guardar los cambios y cerrar la conexión
    conn.commit()
    conn.close()

def leer_contador():
    conn = sqlite3.connect('contador.db')
    c = conn.cursor()
    
    # Leer el valor del contador
    c.execute('''SELECT count FROM contador WHERE id=1''')
    count = c.fetchone()[0]
    
    conn.close()
    return count

def incrementar_contador():
    conn = sqlite3.connect('contador.db')
    c = conn.cursor()
    
    # Incrementar el valor del contador
    c.execute('''UPDATE contador SET count = count + 1 WHERE id=1''')
    
    # Guardar cambios
    conn.commit()
    
    # Leer el nuevo valor del contador
    c.execute('''SELECT count FROM contador WHERE id=1''')
    count = c.fetchone()[0]
    
    conn.close()
    return count

def csv_to_xlsx(csv_file):
    # Leer el archivo CSV con codificación UTF-8
    df = pd.read_csv(csv_file, encoding='utf-8')
    
    # Guardar el DataFrame en un archivo Excel en memoria
    output = BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)  # Mover el cursor al principio del stream
    return output

# Ruta del archivo local de la base de datos
db_file_path = "consulta_rector.xlsx"

# Título y descripción de la aplicación
st.title("📊 Compara tus contactos con la base de datos de habilitados a votar!!")

init_db()  # Inicializa la base de datos

# Sección: Cargar archivo de contactos
st.header("⬇️ Subir tu lista de contactos")
contactos_file = st.file_uploader("Por favor, sube tu archivo de contactos (puede ser en formato CSV o Excel)", type=["csv", "xlsx"], key="contactos_upload")

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

        # Leer el archivo de la base de datos local
        try:
            with st.spinner('Cargando la base de datos seleccionada...'):
                egresados_data = BytesIO(open(db_file_path, 'rb').read())
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
                contador = incrementar_contador()
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

# Mostrar el contador al final de la página
contador = leer_contador()
st.write(f"**Comparaciones realizadas hasta ahora: {contador}**")