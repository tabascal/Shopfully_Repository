import streamlit as st
import pandas as pd
import pptx
import os
import time


# ========================== FUNCIONES AUXILIARES ==========================

def clean_path(path):
    """Limpia y normaliza la ruta asegurando que sea válida."""
    return os.path.normpath(path)  # Normaliza la ruta según el sistema operativo

def ensure_directory_exists(path):
    """Verifica si la ruta existe. Si no, intenta crearla."""
    try:
        if not os.path.exists(path):
            os.makedirs(path, exist_ok=True)
        return True
    except Exception as e:
        st.error(f"❌ Error al crear la carpeta: {e}")
        return False

# ========================== INTERFAZ STREAMLIT ==========================

st.title("Shopfully Dashboard Generator")

# Input para que el usuario seleccione la ruta de guardado
save_path = st.text_input(
    "📂 Ingresa la ruta donde se guardarán los PPTX:",
    value=os.getcwd()  # Usa el directorio actual como valor por defecto
)

# Normalizar la ruta
absolute_save_path = clean_path(save_path)

# Verificar y crear la ruta si es necesario
if ensure_directory_exists(absolute_save_path):
    st.success(f"✅ Archivos se guardarán en: `{absolute_save_path}`")
else:
    st.error("⚠️ No se pudo usar la ruta especificada.")


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
