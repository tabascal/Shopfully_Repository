import streamlit as st
import pandas as pd
import pptx
import os
import time

# ========================== FUNCIONES AUXILIARES ==========================

def clean_path(path):
    """Limpia la ruta asegurando que sea válida."""
    path = os.path.normpath(path)  # Normaliza la ruta según el sistema operativo
    return path

def ensure_directory_exists(path):
    """Crea la carpeta si no existe."""
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)

def list_files_in_directory(path):
    """Lista los archivos en el directorio para verificar que se han guardado."""
    if os.path.exists(path):
        files = os.listdir(path)
        if files:
            st.write("📂 Archivos en la carpeta después de guardar:")
            for file in files:
                st.write(f"📄 {file}")
        else:
            st.warning("⚠️ No hay archivos en la carpeta después de guardar.")
    else:
        st.error(f"🚨 La carpeta `{path}` no existe.")

# ========================== INTERFAZ STREAMLIT ==========================

st.title("Shopfully Dashboard Generator")

# Sección para que el usuario elija la ruta de guardado
save_path = st.text_input(
    "Selecciona la carpeta donde se guardarán los PPTX:",
    value=os.getcwd()  # Directorio actual por defecto
)

# Corregir la ruta antes de usarla
absolute_save_path = clean_path(save_path)

# Crear la carpeta si no existe
ensure_directory_exists(absolute_save_path)

st.write(f"📂 Guardando archivos en: `{absolute_save_path}`")

# ========================== FUNCIONES DE PROCESAMIENTO ==========================

def process_files(ppt_file, excel_file, search_option, start_row, end_row, store_ids, save_path):
    """Procesa los archivos y genera los PPTX en la ruta especificada."""
    
    if ppt_file is None or excel_file is None:
        st.error("⚠️ Error: Debes subir ambos archivos antes de procesar.")
        return
    
    # Verificar si los números ingresados son válidos
    if search_option == "rows" and (start_row is None or end_row is None or start_row > end_row):
        st.error("⚠️ Error: Debes ingresar filas de inicio y fin válidas.")
        return

    # Asegurar que la ruta de guardado es válida
    save_path = clean_path(save_path)
    ensure_directory_exists(save_path)

    # Guardar archivos en la carpeta especificada por el usuario
    ppt_template_path = os.path.join(save_path, ppt_file.name)
    excel_file_path = os.path.join(save_path, excel_file.name)

    try:
        with open(ppt_template_path, "wb") as f:
            f.write(ppt_file.getbuffer())
        
        with open(excel_file_path, "wb") as f:
            f.write(excel_file.getbuffer())

    except Exception as e:
        st.error(f"❌ Error al guardar los archivos: {e}")
        return

    st.success(f"📁 Archivos guardados correctamente en `{save_path}`")

    # Verificar archivos guardados
    list_files_in_directory(save_path)

def process_row(presentation_path, row, save_path):
    """Procesa una fila del dataset y genera un PPTX."""
    presentation = pptx.Presentation(presentation_path)

    # Asegurar que la ruta de guardado es válida
    save_path = clean_path(save_path)
    ensure_directory_exists(save_path)

    output_path = os.path.join(save_path, f"presentation_{row.name}.pptx")

    try:
        presentation.save(output_path)
        st.success(f"✅ Presentación guardada en: `{output_path}`")
    except Exception as e:
        st.error(f"❌ Error al guardar la presentación: {e}")

# ========================== INTERFAZ PARA SUBIR ARCHIVOS ==========================

st.title("PPTX Processor with Streamlit")

ppt_template = st.file_uploader("Sube tu plantilla PPTX", type=["pptx"])
data_file = st.file_uploader("Sube tu dataset (xlsx)", type=["xlsx"])

search_option = st.radio("Filtrar por:", ["rows", "store_id"])

start_row, end_row, store_ids = None, None, None
if search_option == "rows":
    start_row = st.number_input("Fila de inicio", min_value=0, step=1, value=1)
    end_row = st.number_input("Fila de fin", min_value=0, step=1, value=1)
elif search_option == "store_id":
    store_ids = st.text_input("Introduce los Store IDs (separados por comas)")

if st.button("Procesar"):
    process_files(ppt_template, data_file, search_option, start_row, end_row, store_ids, absolute_save_path)
