# main.py
import streamlit as st
import pandas as pd
import vobject
import quopri
from io import BytesIO
import sqlite3
from processing import generar_archivo_combinado, generar_archivo_filtro_unillanos

def init_db():
    conn = sqlite3.connect('contador.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS contador (
                 id INTEGER PRIMARY KEY,
                 count INTEGER)''')
    c.execute('''INSERT OR IGNORE INTO contador (id, count) VALUES (1, 0)''')
    conn.commit()
    conn.close()

def leer_contador():
    conn = sqlite3.connect('contador.db')
    c = conn.cursor()
    c.execute('''SELECT count FROM contador WHERE id=1''')
    count = c.fetchone()[0]
    conn.close()
    return count

def incrementar_contador():
    conn = sqlite3.connect('contador.db')
    c = conn.cursor()
    c.execute('''UPDATE contador SET count = count + 1 WHERE id=1''')
    conn.commit()
    c.execute('''SELECT count FROM contador WHERE id=1''')
    count = c.fetchone()[0]
    conn.close()
    return count

def csv_to_xlsx(csv_file):
    df = pd.read_csv(csv_file, encoding='utf-8')
    output = BytesIO()
    df.to_excel(output, index=False, engine='openpyxl')
    output.seek(0)
    return output

def buscar_por_cedula(cedula, egresados_df):
    resultado = egresados_df[egresados_df['Cedula'] == cedula]
    if not resultado.empty:
        nombre = resultado['Nombres'].values[0]
        tipo = resultado['Tipo'].values[0]
        return nombre, tipo
    else:
        return None, None

def convertir_vcf_a_csv(vcf_file):
    contactos = []

    try:
        vcf_content = vcf_file.read().decode('utf-8')
    except UnicodeDecodeError:
        vcf_file.seek(0)
        vcf_content = vcf_file.read().decode('ISO-8859-1')

    vcf_content = quopri.decodestring(vcf_content).decode('ISO-8859-1')
    
    filtered_lines = []
    skip_line = False
    for line in vcf_content.splitlines():
        if line.startswith("PHOTO;ENCODING=") or line.startswith("PHOTO;"):
            skip_line = True
        elif skip_line and not line.startswith(" "):
            skip_line = False
        if not skip_line:
            filtered_lines.append(line)
    
    vcf_content_filtered = "\n".join(filtered_lines)

    for vcard in vobject.readComponents(vcf_content_filtered):
        contacto = {}
        contacto['Nombre'] = vcard.fn.value if hasattr(vcard, 'fn') else ''
        contacto['Teléfono'] = vcard.tel.value if hasattr(vcard, 'tel') else ''
        contacto['Email'] = vcard.email.value if hasattr(vcard, 'email') else ''
        contactos.append(contacto)

    df = pd.DataFrame(contactos)
    
    output = BytesIO()
    df.to_csv(output, index=False, encoding='utf-8')
    output.seek(0)
    
    return output

db_file_path = "consulta_rector.xlsx"

st.title("📊 Compara tus contactos y más!")

init_db()

# Agregando la navegación entre pestañas
tab1, tab2, tab3, tab4 = st.tabs(["Comparar Contactos", "Filtro Unillanos", "Buscar por Cédula", "Convertir VCF a CSV"])

with tab1:
    st.header("⬇️ Subir tu lista de contactos")
    contactos_file = st.file_uploader("Por favor, sube tu archivo de contactos (puede ser en formato CSV o Excel)", type=["csv", "xlsx"], key="contactos_upload")
    output_filename = st.text_input("💾 Ingrese el nombre del archivo de salida (sin extensión)", value="resultado")
    progress_bar = st.progress(0)

    if st.button("Click para comparar!! 👇"):
        if contactos_file:
            with st.spinner('Procesando tu archivo de contactos...'):
                if contactos_file.name.endswith('.csv'):
                    contactos_data = csv_to_xlsx(contactos_file)
                else:
                    contactos_data = BytesIO(contactos_file.getvalue())
            try:
                with st.spinner('Cargando la base de datos seleccionada...'):
                    egresados_data = BytesIO(open(db_file_path, 'rb').read())
            except Exception as e:
                st.error(f"⚠️ Ocurrió un error al cargar la base de datos: {str(e)}")
                egresados_data = None

            if egresados_data:
                output_data = BytesIO()
                try:
                    with st.spinner('Comparando y generando el archivo final...'):
                        generar_archivo_combinado(contactos_data, egresados_data, output_data, progress_bar)
                        output_data.seek(0)

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

    contador = leer_contador()
    st.write(f"**Comparaciones realizadas hasta ahora: {contador}**")

with tab2:
    st.header("🎓 Filtro Unillanos")
    contactos_file = st.file_uploader("Sube tu archivo de contactos para filtrar por 'U' o 'Unillanos'", type=["csv", "xlsx"], key="filtro_unillanos_upload")
    output_filename = st.text_input("💾 Ingrese el nombre del archivo de salida (sin extensión)", value="resultado_unillanos")
    progress_bar = st.progress(0)

    if st.button("Aplicar filtro y comparar!"):
        if contactos_file:
            with st.spinner('Procesando tu archivo de contactos...'):
                if contactos_file.name.endswith('.csv'):
                    contactos_data = csv_to_xlsx(contactos_file)
                else:
                    contactos_data = BytesIO(contactos_file.getvalue())
            try:
                with st.spinner('Cargando la base de datos seleccionada...'):
                    egresados_data = BytesIO(open(db_file_path, 'rb').read())
            except Exception as e:
                st.error(f"⚠️ Ocurrió un error al cargar la base de datos: {str(e)}")
                egresados_data = None

            if egresados_data:
                output_data = BytesIO()
                try:
                    with st.spinner('Aplicando filtro y generando el archivo final...'):
                        generar_archivo_filtro_unillanos(contactos_data, egresados_data, output_data, progress_bar)
                        output_data.seek(0)

                    st.success("✅ Filtro aplicado y comparación completada con éxito!")
                    contador = incrementar_contador()
                    st.download_button(
                        label="⬇️ Descargar archivo filtrado y procesado",
                        data=output_data,
                        file_name=f"{output_filename}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"⚠️ Ocurrió un error: {str(e)}")
        else:
            st.error("⚠️ Por favor, sube un archivo de contactos.")

    contador = leer_contador()
    st.write(f"**Comparaciones realizadas hasta ahora: {contador}**")

with tab3:
    st.header("🔍 Buscar por número de cédula")
    cedula = st.number_input("Ingrese el número de cédula:", min_value=0, step=1, format="%d")

    if st.button("Buscar"):
        with st.spinner('Buscando en la base de datos...'):
            try:
                egresados_df = pd.read_excel(db_file_path)
                nombre, tipo = buscar_por_cedula(cedula, egresados_df)
                if nombre:
                    st.success(f"Nombre: {nombre}")
                    st.info(f"Tipo: {tipo}")
                else:
                    st.warning("⚠️ No se encontró ningún registro con esa cédula.")
            except Exception as e:
                st.error(f"⚠️ Ocurrió un error al buscar la cédula: {str(e)}")

with tab4:
    st.header("📤 Convertir VCF a CSV UTF-8")
    vcf_file = st.file_uploader("Sube tu archivo VCF", type=["vcf"])

    if st.button("Convertir"):
        if vcf_file:
            with st.spinner('Convirtiendo VCF a CSV...'):
                try:
                    csv_output = convertir_vcf_a_csv(vcf_file)
                    st.success("✅ Conversión completada con éxito!")
                    st.download_button(
                        label="⬇️ Descargar CSV",
                        data=csv_output,
                        file_name="contactos.csv",
                        mime="text/csv"
                    )
                except Exception as e:
                    st.error(f"⚠️ Ocurrió un error al convertir el archivo: {str(e)}")
        else:
            st.error("⚠️ Por favor, sube un archivo VCF.")