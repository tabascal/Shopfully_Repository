import streamlit as st
import pandas as pd
import pptx
import os
import time


# ========================== FUNCIONES AUXILIARES ==========================

def clean_path(path):
    """Limpia y normaliza la ruta asegurando que sea válida."""
    return os.path.normpath(path)

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

def process_files(ppt_file, save_path):
    """Procesa el archivo y guarda el PPTX en la ruta especificada."""
    
    if ppt_file is None:
        st.error("⚠️ Error: No se ha subido ningún archivo PPTX.")
        return

    save_path = clean_path(save_path)
    ensure_directory_exists(save_path)

    # Ruta final donde se guardará
    ppt_output_path = os.path.join(save_path, "output_presentation.pptx")

    try:
        # Guardar archivo subido
        with open(ppt_output_path, "wb") as f:
            f.write(ppt_file.getbuffer())

        # Verificar que el archivo realmente se guardó
        if os.path.exists(ppt_output_path):
            st.success(f"✅ Presentación guardada correctamente en: `{ppt_output_path}`")
        else:
            st.error("❌ No se encontró el archivo después de guardarlo.")

        # Mostrar archivos en la carpeta
        st.write("📂 Archivos en la carpeta después de guardar:")
        st.write(os.listdir(save_path))

    except Exception as e:
        st.error(f"❌ Error al guardar el archivo PPTX: {e}")


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
    process_files(ppt_template, absolute_save_path)
