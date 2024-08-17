import streamlit as st
import pandas as pd
from processing import generar_archivo_combinado
import tempfile  # Import the tempfile module

# Title and description
st.title("Excel File Processor")
st.write("This application processes Excel files by matching and sorting data.")

# File uploaders for Contactos and Egresados
contactos_file = st.file_uploader("Upload Contactos File", type=["xlsx"])
egresados_file = st.file_uploader("Upload Egresados File", type=["xlsx"])

# Text input for Output file name
output_filename = st.text_input("Enter Output File Name (without extension)", value="output")

# Progress bar
progress_bar = st.progress(0)

# Process button
if st.button("Process Files"):
    if contactos_file and egresados_file:
        # Create temporary files for uploaded files
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as contactos_temp:
            contactos_temp.write(contactos_file.getbuffer())
            contactos_path = contactos_temp.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as egresados_temp:
            egresados_temp.write(egresados_file.getbuffer())
            egresados_path = egresados_temp.name
        
        # Create temporary file for output
        output_path = f"{output_filename}.xlsx"
        
        # Run the processing function
        try:
            generar_archivo_combinado(contactos_path, egresados_path, output_path, progress_bar)
            st.success("Processing completed successfully!")

            # Provide download link for the output file
            with open(output_path, "rb") as file:
                btn = st.download_button(
                    label="Download Processed File",
                    data=file,
                    file_name=f"{output_filename}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
    else:
        st.error("Please upload both Contactos and Egresados files.")