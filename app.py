import streamlit as st
import pandas as pd
import pptx
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
import os
import time

# Función para asegurar que la ruta es válida en el sistema operativo
def clean_path(path):
    """Limpia la ruta asegurando que no tenga prefijos incorrectos en entornos Linux/Windows."""
    path = os.path.normpath(path)  # Normaliza la ruta
    if os.name == "nt":  # Solo en Windows
        if path.startswith("/"):
            path = path.lstrip("/mnt/src/shopfully_repository/")  # Quita prefijo innecesario
    return os.path.abspath(path)

# Streamlit app
st.title("Shopfully Dashboard Generator")

# Sección para seleccionar la ruta de guardado
save_path = st.text_input(
    "Selecciona la carpeta donde se guardarán los PPTX:",
    value=os.getcwd()  # Directorio actual por defecto
)

# Verificar y corregir la ruta
absolute_save_path = clean_path(save_path)

if not os.path.exists(absolute_save_path):
    os.makedirs(absolute_save_path, exist_ok=True)  # Crea la carpeta si no existe

st.write(f"📂 Guardando archivos en: `{absolute_save_path}`")

# Función para actualizar los cuadros de texto en una presentación
def update_text_of_textbox(presentation, column_letter, new_text):
    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text:
                if shape.text.strip() == column_letter:
                    text_frame = shape.text_frame
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.text = str(new_text)

# Botón para probar el guardado
if st.button("Guardar presentación de prueba"):
    presentation = pptx.Presentation()
    test_file_path = os.path.join(absolute_save_path, "prueba_guardado.pptx")
    presentation.save(test_file_path)
    st.success(f"✅ Presentación guardada en: `{test_file_path}`")

# Función para procesar archivos
def process_files(ppt_file, excel_file, search_option, start_row, end_row, store_ids, file_name_order_1, file_name_order_2, file_name_order_3, save_path):
    ppt_template_path = os.path.join("uploads", ppt_file.name)
    excel_file_path = os.path.join("uploads", excel_file.name)

    # Guardar los archivos temporales
    with open(ppt_template_path, "wb") as f:
        f.write(ppt_file.getbuffer())
    with open(excel_file_path, "wb") as f:
        f.write(excel_file.getbuffer())

    # Leer datos del Excel
    try:
        with pd.ExcelFile(excel_file_path) as xls:
            df1 = pd.read_excel(xls, sheet_name=0)  # Primera hoja
            df2 = pd.read_excel(xls, sheet_name=1)  # Segunda hoja
    except PermissionError as e:
        st.error(f"Error al leer el Excel: {e}")
        return

    if search_option == 'rows':
        for index, row in df1.iterrows():
            if index < start_row or index > end_row:
                continue
            process_row(ppt_template_path, row, df1, df2, index, file_name_order_1, file_name_order_2, file_name_order_3, save_path)

    elif search_option == 'store_id':
        store_id_list = [store_id.strip() for store_id in store_ids.split(',')]
        for store_id in store_id_list:
            matching_rows = df1[df1.iloc[:, 0].astype(str) == store_id]
            if not matching_rows.empty:
                row = matching_rows.iloc[0]
                index = row.name
                process_row(ppt_template_path, row, df1, df2, index, file_name_order_1, file_name_order_2, file_name_order_3, save_path)

# Función para procesar una fila y generar un PPTX
def process_row(presentation_path, row, df1, df2, index, file_name_order_1, file_name_order_2, file_name_order_3, save_path):
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

    file_name = '_'.join(file_name_parts) if file_name_parts else f"presentation_{index}"

    # Asegurar que la ruta de guardado es válida
    absolute_save_path = clean_path(save_path)

    # Crear la carpeta si no existe
    if not os.path.exists(absolute_save_path):
        os.makedirs(absolute_save_path, exist_ok=True)

    output_path = os.path.join(absolute_save_path, f"{file_name}.pptx")
    output_path = clean_path(output_path)  # Asegurar que es una ruta válida

    try:
        presentation.save(output_path)
        os.sync()  # Forzar la escritura en el sistema de archivos (en Linux/macOS)
        st.success(f"✅ Presentación guardada en: `{output_path}`")
    except Exception as e:
        st.error(f"❌ Error al guardar la presentación: {e}")

# Interfaz de Streamlit para subir archivos
st.title("PPTX Processor with Streamlit")

ppt_template = st.file_uploader("Sube tu plantilla PPTX", type=["pptx"])
data_file = st.file_uploader("Sube tu dataset (xlsx)", type=["xlsx"])

search_option = st.radio("Filtrar por:", ["rows", "store_id"])

start_row, end_row, store_ids = None, None, None
if search_option == "rows":
    start_row = st.number_input("Fila de inicio", min_value=0, step=1)
    end_row = st.number_input("Fila de fin", min_value=0, step=1)
elif search_option == "store_id":
    store_ids = st.text_input("Introduce los Store IDs (separados por comas)")

file_name_order_1 = st.text_input("Orden de nombre de archivo 1 (Índice de columna)")
file_name_order_2 = st.text_input("Orden de nombre de archivo 2 (Índice de columna)")
file_name_order_3 = st.text_input("Orden de nombre de archivo 3 (Índice de columna)")

if st.button("Procesar"):
    if ppt_template and data_file:
        process_files(ppt_template, data_file, search_option, start_row, end_row, store_ids, file_name_order_1, file_name_order_2, file_name_order_3, absolute_save_path)
    else:
        st.error("⚠️ Debes subir ambos archivos antes de procesar.")
