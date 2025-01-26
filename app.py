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
            if file.endswith(".pptx"):
                zipf.write(os.path.join(folder_path, file), arcname=file)

    zip_buffer.seek(0)
    return zip_buffer


# Crear la carpeta de subidas si no existe
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Estado global para progreso
progress = 0

def get_filename_from_selection(row, selected_columns):
    """Genera el nombre del archivo según las columnas seleccionadas."""
    file_name_parts = [str(row[col]) for col in selected_columns if col in row]
    return "_".join(file_name_parts)


def update_text_of_textbox(presentation, column_letter, new_text):
    """Busca y reemplaza texto dentro de las cajas de texto que tengan el formato {A}, {B}, etc."""
    pattern = rf"\{{{
        column_letter}\}}"  # Expresión regular para encontrar "{A}", "{B}", etc.

    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text:
                if re.search(pattern, shape.text):  # Buscar patrón en el texto
                    text_frame = shape.text_frame
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.text = re.sub(pattern, str(
                                new_text), run.text)  # Reemplazo


def process_files(ppt_file, excel_file, search_option, start_row, end_row, store_ids, file_name_order_1, file_name_order_2, file_name_order_3):
    # Crear un identificador único basado en la fecha y hora actual
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Crear una carpeta de salida única
    unique_output_folder = os.path.join(
        UPLOAD_FOLDER, f"pptx_files_{timestamp}")
    os.makedirs(unique_output_folder, exist_ok=True)

    # Guardar archivos en la carpeta temporal
    ppt_template_path = os.path.join(unique_output_folder, ppt_file.name)
    excel_file_path = os.path.join(unique_output_folder, excel_file.name)

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

    # Definir el número total de archivos a generar
    total_files = 0
    if search_option == 'rows':
        total_files = end_row - start_row + 1
    elif search_option == 'store_id':
        total_files = len(store_ids.split(','))

    # Crear una barra de progreso
    progress_bar = st.progress(0)
    progress_text = st.empty()

    current_file = 0  # Contador de archivos generados

    if search_option == 'rows':
        for index, row in df1.iterrows():
            if index < start_row or index > end_row:
                continue

            process_row(ppt_template_path, row, df1, index, file_name_order_1,
                        file_name_order_2, file_name_order_3, unique_output_folder)

            current_file += 1
            progress = current_file / total_files
            progress_bar.progress(progress)  # Actualiza la barra de progreso
            progress_text.write(f"📄 Generando presentación {
                                current_file}/{total_files}")

    elif search_option == 'store_id':
        store_id_list = [store_id.strip() for store_id in store_ids.split(',')]

        for store_id in store_id_list:
            matching_rows = df1[df1.iloc[:, 0].astype(str) == store_id]
            if matching_rows.empty:
                st.warning(f"No matching rows found for Store ID: {store_id}")
                continue

            row = matching_rows.iloc[0]
            index = row.name

            process_row(ppt_template_path, row, df1, index, file_name_order_1,
                        file_name_order_2, file_name_order_3, unique_output_folder)

            current_file += 1
            progress = current_file / total_files
            progress_bar.progress(progress)  # Actualiza la barra de progreso
            progress_text.write(f"📄 Generando presentación {
                                current_file}/{total_files}")

    # Crear un ZIP único con la carpeta generada
    unique_zip_path = os.path.join(
        UPLOAD_FOLDER, f"presentaciones_{timestamp}.zip")
    shutil.make_archive(unique_zip_path.replace(
        ".zip", ""), 'zip', unique_output_folder)

    # Mostrar el botón de descarga
    with open(unique_zip_path, "rb") as zip_file:
        st.download_button(
            label=f"📥 Descargar {total_files} presentaciones",
            data=zip_file,
            file_name=f"presentaciones_{timestamp}.zip",
            mime="application/zip"
        )

    # Indicar que la generación ha finalizado
    progress_text.write("✅ ¡Todas las presentaciones han sido generadas!")


def process_row(presentation_path, row, df1, index, file_name_order_1, file_name_order_2, file_name_order_3, output_folder):
    """Procesa una fila del dataset y genera un PPTX en la carpeta de salida."""
    presentation = pptx.Presentation(presentation_path)

    for col_idx, col_name in enumerate(row.index):
        # Convertimos índice numérico en letra A-Z
        column_letter = chr(65 + col_idx)
        update_text_of_textbox(presentation, column_letter,
                               row[col_name])  # Pasamos letra sin {}

    file_name_parts = []
    for order in [file_name_order_1, file_name_order_2, file_name_order_3]:
        if order:
            try:
                idx = int(order)
                if idx < len(row):
                    file_name_parts.append(str(row.iloc[idx]))
            except ValueError:
                continue

    file_name = '_'.join(
        file_name_parts) if file_name_parts else f"presentation_{index}"
    output_path = os.path.join(output_folder, f"{file_name}.pptx")
    presentation.save(output_path)


# Interfaz de Streamlit
st.title("PPTX Processor with Streamlit")

ppt_template = st.file_uploader("Upload PPTX Template", type=["pptx"])
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
        "📂 Selecciona y ordena las columnas para el nombre del archivo:",
        column_names,
        default=column_names[:3]
    )

    st.write("🔹 Ejemplo de nombre de archivo:", get_filename_from_selection(df.iloc[0], selected_columns))

if st.button("Process"):
    if ppt_template and data_file:
        process_files(ppt_template, data_file, search_option, start_row, end_row, store_ids, selected_columns)
    else:
        st.error("Please upload both files before processing.")