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
st.title("Aplicación para modificar presentaciones PPTX")

# Sección para que el usuario elija la ruta de guardado
save_path = st.text_input(
    "Introduce la ruta donde deseas guardar los archivos PPTX:", value=os.getcwd())

# Botón para guardar la presentación
if st.button("Guardar presentación"):
    # Aquí puedes agregar el código para guardar la presentación en la ruta especificada
    presentation = pptx.Presentation()
    # ... código para modificar la presentación ...
    save_file_path = os.path.join(save_path, "presentacion_modificada.pptx")
    presentation.save(save_file_path)
    st.success(f"Presentación guardada en: {save_file_path}")


def process_files(ppt_file, excel_file, search_option, start_row, end_row, store_ids, file_name_order_1, file_name_order_2, file_name_order_3):
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
                        file_name_order_1, file_name_order_2, file_name_order_3)

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
                        file_name_order_1, file_name_order_2, file_name_order_3)

            current_id += 1
            progress = int((current_id / total_ids) * 100)
            time.sleep(1)
            st.progress(progress / 100)


def process_row(presentation_path, row, sheet_names, df1, df2, index, file_name_order_1, file_name_order_2, file_name_order_3):
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
    output_path = os.path.join(UPLOAD_FOLDER, f"{file_name}.pptx")
    presentation.save(output_path)

    st.success(f"Presentation saved successfully: {file_name}.pptx")
    with open(output_path, "rb") as f:
        st.download_button(label="Download PPTX", data=f,
                           file_name=f"{file_name}.pptx")


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

file_name_order_1 = st.text_input("File Name Order 1 (Column Index)")
file_name_order_2 = st.text_input("File Name Order 2 (Column Index)")
file_name_order_3 = st.text_input("File Name Order 3 (Column Index)")

if st.button("Process"):
    if ppt_template and data_file:
        process_files(ppt_template, data_file, search_option, start_row, end_row,
                      store_ids, file_name_order_1, file_name_order_2, file_name_order_3)
    else:
        st.error("Please upload both files before processing.")
