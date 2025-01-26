import streamlit as st
import pandas as pd
import pptx
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
import os
import time
import zipfile
import io
import shutil
from datetime import datetime
import re


def create_zip_of_presentations(folder_path):
    """Crea un archivo ZIP con todos los PPTX generados en la carpeta."""
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file)
            if file.endswith(".pptx"):  # Evitamos incluir plantilla y Excel
                zipf.write(file_path, arcname=file)
    
    zip_buffer.seek(0)
    return zip_buffer


def get_filename_from_selection(row, selected_columns):
    """Genera el nombre del archivo según las columnas seleccionadas."""
    file_name_parts = [str(row[col]) for col in selected_columns if col in row]
    return "_".join(file_name_parts)


def update_text_of_textbox(presentation, column_letter, new_text):
    """Busca y reemplaza texto dentro de las cajas de texto que tengan el formato {A}, {B}, etc."""
    pattern = rf"\{{{column_letter}\}}"  # Expresión regular para encontrar "{A}", "{B}", etc.

    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text:
                if re.search(pattern, shape.text):  # Buscar patrón en el texto
                    text_frame = shape.text_frame
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.text = re.sub(pattern, str(new_text), run.text)  # Reemplazo


def process_files(ppt_file, excel_file, search_option, start_row, end_row, store_ids, selected_columns):
    """Procesa los archivos y genera las presentaciones."""
    
    # Crear un identificador único basado en la fecha y hora actual
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Nombre único para la carpeta y el ZIP
    folder_name = f"Presentations_{timestamp}"

    # Crear carpeta de salida
    os.makedirs(folder_name, exist_ok=True)

    # Guardar archivos en una carpeta temporal fuera de la de presentaciones
    temp_folder = "temp_files"
    os.makedirs(temp_folder, exist_ok=True)

    ppt_template_path = os.path.join(temp_folder, ppt_file.name)
    excel_file_path = os.path.join(temp_folder, excel_file.name)

    with open(ppt_template_path, "wb") as f:
        f.write(ppt_file.getbuffer())
    with open(excel_file_path, "wb") as f:
        f.write(excel_file.getbuffer())

    # Leer datos del archivo Excel
    try:
        with pd.ExcelFile(excel_file_path) as xls:
            df1 = pd.read_excel(xls, sheet_name=0)  # Primera hoja
    except PermissionError as e:
        st.error(f"Error reading Excel file: {e}")
        return

    # Definir correctamente el número total de archivos a generar
    if search_option == 'rows':
        total_files = len(df1.iloc[start_row:end_row + 1])  # Solo las filas seleccionadas
    elif search_option == 'store_id':
        store_id_list = [store_id.strip() for store_id in store_ids.split(',')]
        total_files = sum(df1.iloc[:, 0].astype(str).isin(store_id_list))  # Solo los Store ID seleccionados
    else:
        total_files = 0

    if total_files == 0:
        st.error("⚠️ No hay archivos para generar. Verifica los filtros.")
        return

    # Crear una barra de progreso
    progress_bar = st.progress(0)
    progress_text = st.empty()

    current_file = 0  # Contador de archivos generados

    for index, row in df1.iterrows():
        if search_option == 'rows' and (index < start_row or index > end_row):
            continue
        elif search_option == 'store_id' and str(row.iloc[0]) not in store_ids.split(','):
            continue

        process_row(ppt_template_path, row, df1, index, selected_columns, folder_name)

        current_file += 1
        progress = current_file / total_files
        progress_bar.progress(progress)  # Actualiza la barra de progreso
        progress_text.write(f"📄 Generating presentation {current_file}/{total_files}")  # ✅ Se muestra el número correcto

    # Crear un ZIP único sin la plantilla ni el Excel
    zip_path = f"{folder_name}.zip"
    shutil.make_archive(zip_path.replace(".zip", ""), 'zip', folder_name)

    # Mostrar el botón de descarga
    with open(zip_path, "rb") as zip_file:
        st.download_button(
            label=f"📥 Download {total_files} presentations",
            data=zip_file,
            file_name=f"{folder_name}.zip",
            mime="application/zip"
        )

    # Indicar que la generación ha finalizado
    progress_text.write("✅ All presentations have been generated!")

def process_row(presentation_path, row, df1, index, selected_columns, output_folder):
    """Procesa una fila del dataset y genera un PPTX en la carpeta de salida."""
    presentation = pptx.Presentation(presentation_path)

    for col_idx, col_name in enumerate(row.index):
        column_letter = chr(65 + col_idx)
        update_text_of_textbox(presentation, column_letter, row[col_name])

    file_name = get_filename_from_selection(row, selected_columns)
    output_path = os.path.join(output_folder, f"{file_name}.pptx")
    presentation.save(output_path)


# Interfaz de Streamlit
st.title("Shopfully Dashboard Generator")

ppt_template = st.file_uploader("Upload PPTX Template (Text Box format to edit {X})", type=["pptx"])
data_file = st.file_uploader("Upload Excel File", type=["xlsx"])

search_option = st.radio("Search by:", ["rows", "store_id"])

start_row, end_row, store_ids = None, None, None
if search_option == "rows":
    start_row = st.number_input("Start Row", min_value=0, step=1)
    end_row = st.number_input("End Row", min_value=0, step=1)
elif search_option == "store_id":
    store_ids = st.text_input("Enter Store IDs (comma-separated)")

if data_file is not None:
    df = pd.read_excel(data_file, sheet_name=0)  # Leer la primera hoja del Excel
    column_names = df.columns.tolist()

    selected_columns = st.multiselect(
        "📂 Select and order the columns for the file name:",
        column_names,
        default=column_names[:3]
    )

    st.write("🔹 Example file name:", get_filename_from_selection(df.iloc[0], selected_columns))

if st.button("Process"):
    if ppt_template and data_file:
        process_files(ppt_template, data_file, search_option, start_row, end_row, store_ids, selected_columns)
    else:
        st.error("Please upload both files before processing.")
