import streamlit as st
import pandas as pd
import pptx
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
import os
import time

# Crear la carpeta de subidas si no existe
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Estado global para progreso
progress = 0

# Función para actualizar el texto de un cuadro de texto en una presentación


def update_text_of_textbox(presentation, column_letter, new_text):
    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text:
                if shape.text.strip() == column_letter:
                    text_frame = shape.text_frame
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.text = new_text


# Streamlit app
st.title("Shopfully Dashboard Generator")

# Sección para que el usuario elija la ruta de guardado
save_path = st.text_input(
    "Select the path where the PPTX will be stored:",
    value=os.getcwd()  # Ruta predeterminada: directorio actual
)

# Verificar si la ruta es válida
if not os.path.exists(save_path):
    st.warning(
        "La ruta de guardado especificada no existe. Se usará el directorio predeterminado.")

# Función para procesar archivos


def process_files(ppt_file, excel_file, search_option, start_row, end_row, store_ids, file_name_order_1, file_name_order_2, file_name_order_3, save_path):
    global progress

    # Guardar archivos en la carpeta temporal
    ppt_template_path = os.path.join(UPLOAD_FOLDER, ppt_file.name)
    excel_file_path = os.path.join(UPLOAD_FOLDER, excel_file.name)

    with open(ppt_template_path, "wb") as f:
        f.write(ppt_file.getbuffer())
    with open(excel_file_path, "wb") as f:
        f.write(excel_file.getbuffer())

    # Leer datos del archivo Excel
    try:
        with pd.ExcelFile(excel_file_path) as xls:
            df1 = pd.read_excel(xls, sheet_name=0)  # Primera hoja
            df2 = pd.read_excel(xls, sheet_name=1)  # Segunda hoja
            sheet_names = xls.sheet_names
    except PermissionError as e:
        st.error(f"Error reading Excel file: {e}")
        return

    if search_option == 'rows':
        total_rows = end_row - start_row + 1
        current_row = 0

        for index, row in df1.iterrows():
            if index < start_row or index > end_row:
                continue

            process_row(ppt_template_path, row, sheet_names, df1, df2, index,
                        file_name_order_1, file_name_order_2, file_name_order_3, save_path)

            current_row += 1
            progress = int((current_row / total_rows) * 100)
            time.sleep(1)
            st.progress(progress / 100)

    elif search_option == 'store_id':
        store_id_list = [store_id.strip() for store_id in store_ids.split(',')]
        total_ids = len(store_id_list)
        current_id = 0

        for store_id in store_id_list:
            matching_rows = df1[df1.iloc[:, 0].astype(str) == store_id]
            if matching_rows.empty:
                st.warning(f"No matching rows found for Store ID: {store_id}")
                continue

            row = matching_rows.iloc[0]
            index = row.name

            process_row(ppt_template_path, row, sheet_names, df1, df2, index,
                        file_name_order_1, file_name_order_2, file_name_order_3, save_path)

            current_id += 1
            progress = int((current_id / total_ids) * 100)
            time.sleep(1)
            st.progress(progress / 100)

# Función para procesar una fila y generar un archivo PPTX


def process_row(presentation_path, row, sheet_names, df1, df2, index, file_name_order_1, file_name_order_2, file_name_order_3, save_path):
    presentation = pptx.Presentation(presentation_path)

    for col_idx, col_name in enumerate(row.index):
        column_letter = chr(65 + col_idx)
        update_text_of_textbox(presentation, column_letter, row[col_name])

    file_name_parts = []
    for order in [file_name_order_1, file_name_order_2, file_name_order_3]:
        if order:
            try:
                idx = int(order)
                if idx < len(row):
                    file_name_parts.append(str(row.iloc[idx]))
            except ValueError:
                continue

    file_name = '_'.join(file_name_parts)

    # Usar la ruta proporcionada por el usuario
    output_path = os.path.join(save_path, f"{file_name}.pptx")

    presentation.save(output_path)

    st.success(f"Presentación guardada correctamente: {output_path}")
    with open(output_path, "rb") as f:
        st.download_button(label="Descargar PPTX", data=f,
                           file_name=f"{file_name}.pptx")


# Interfaz de Streamlit
st.title("PPTX Processor with Streamlit")

ppt_template = st.file_uploader("Sube la plantilla PPTX", type=["pptx"])
data_file = st.file_uploader("Sube el archivo Excel", type=["xlsx"])

search_option = st.radio("Buscar por:", ["rows", "store_id"])

start_row, end_row, store_ids = None, None, None
if search_option == "rows":
    start_row = st.number_input("Fila de inicio", min_value=0, step=1)
    end_row = st.number_input("Fila de fin", min_value=0, step=1)
elif search_option == "store_id":
    store_ids = st.text_input("Introduce los Store IDs (separados por comas)")

file_name_order_1 = st.text_input(
    "Orden de nombre de archivo 1 (Índice de columna)")
file_name_order_2 = st.text_input(
    "Orden de nombre de archivo 2 (Índice de columna)")
file_name_order_3 = st.text_input(
    "Orden de nombre de archivo 3 (Índice de columna)")

if st.button("Procesar"):
    if ppt_template and data_file:
        process_files(ppt_template, data_file, search_option, start_row, end_row,
                      store_ids, file_name_order_1, file_name_order_2, file_name_order_3, save_path)
    else:
        st.error("Por favor, sube ambos archivos antes de procesar.")
