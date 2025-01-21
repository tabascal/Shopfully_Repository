from flask import Flask, request, render_template, redirect, url_for, jsonify
import pandas as pd
import pptx
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
import os
import time

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

progress = 0


def update_text_of_textbox(presentation, column_letter, new_text):
    for slide in presentation.slides:
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text:
                if shape.text.strip() == column_letter:
                    text_frame = shape.text_frame
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            run.text = str(new_text)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_files():
    global progress
    progress = 0
    if 'ppt_template' not in request.files or 'data_file' not in request.files:
        return 'No file part'
    ppt_template = request.files['ppt_template']
    data_file = request.files['data_file']
    search_option = request.form['search_option']
    start_row = request.form.get('start_row', type=int)
    end_row = request.form.get('end_row', type=int)
    store_ids = request.form.get('store_id')
    file_name_order_1 = request.form.get('file_name_order_1')
    file_name_order_2 = request.form.get('file_name_order_2')
    file_name_order_3 = request.form.get('file_name_order_3')
    if ppt_template.filename == '' or data_file.filename == '':
        return 'No selected file'
    ppt_template_path = os.path.join(
        app.config['UPLOAD_FOLDER'], ppt_template.filename)
    data_file_path = os.path.join(
        app.config['UPLOAD_FOLDER'], data_file.filename)
    ppt_template.save(ppt_template_path)
    data_file.save(data_file_path)
    process_files(ppt_template_path, data_file_path,
                  search_option, start_row, end_row, store_ids, file_name_order_1, file_name_order_2, file_name_order_3)
    os.remove(ppt_template_path)  # Eliminar la plantilla después de usarla
    os.remove(data_file_path)  # Eliminar el archivo de datos después de usarlo
    return redirect(url_for('index'))


@app.route('/progress')
def progress_status():
    global progress
    return jsonify({'progress': progress})


def process_files(presentation_path, excel_path, search_option, start_row, end_row, store_ids, file_name_order_1, file_name_order_2, file_name_order_3):
    global progress
    # Leer datos del archivo Excel
    try:
        with pd.ExcelFile(excel_path) as xls:
            df1 = pd.read_excel(xls, sheet_name=0)  # Primera hoja
            df2 = pd.read_excel(xls, sheet_name=1)  # Segunda hoja
            sheet_names = xls.sheet_names
    except PermissionError as e:
        print(f"Error reading Excel file: {e}")
        return

    if search_option == 'rows':
        total_rows = end_row - start_row + 1
        current_row = 0

    # Iterar sobre las filas de la primera hoja del DataFrame
        for index, row in df1.iterrows():
            if index < start_row or index > end_row:
                continue

            # Procesar la fila
            process_row(presentation_path, row, sheet_names,
                        df1, df2, index, file_name_order_1, file_name_order_2, file_name_order_3)

            # Actualizar el progreso
            current_row += 1
            progress = int((current_row / total_rows) * 100)
            time.sleep(1)  # Simular tiempo de procesamiento

    elif search_option == 'store_id':
        store_id_list = [store_id.strip() for store_id in store_ids.split(',')]
        total_ids = len(store_id_list)
        current_id = 0

        for store_id in store_id_list:
            # Buscar la fila correspondiente al Store ID
            print(f"Searching for Store ID: {store_id}")
            matching_rows = df1[df1.iloc[:, 0].astype(str) == store_id]
            if matching_rows.empty:
                print(f"No matching rows found for Store ID: {store_id}")
                continue

        row = matching_rows.iloc[0]
        index = row.name
        print(f"Found matching row for Store ID: {store_id} at index {index}")

        # Procesar la fila
        process_row(presentation_path, row, sheet_names,
                    df1, df2, index, file_name_order_1, file_name_order_2, file_name_order_3)

        # Actualizar el progreso
        current_id += 1
        progress = int((current_id / total_ids) * 100)
        time.sleep(1)  # Simular tiempo de procesamiento


def process_row(presentation_path, row, sheet_names, df1, df2, index, file_name_order_1, file_name_order_2, file_name_order_3):
    # Cargar una copia de la presentación base
    presentation = pptx.Presentation(presentation_path)

    # Iterar sobre las columnas del DataFrame y actualizar las text boxes correspondientes
    for col_idx, col_name in enumerate(row.index):
        # Convertir índice de columna a letra (A, B, C, ...)
        column_letter = chr(65 + col_idx)
        update_text_of_textbox(presentation, column_letter, row[col_name])


# Generar el nombre del archivo según la elección del usuario
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

    output_path = os.path.join(
        app.config['UPLOAD_FOLDER'], f"{file_name}.pptx")
    presentation.save(output_path)
    print(f"Presentation saved successfully to {output_path}")


# Inicia el servidor Flask solo si este archivo se ejecuta directamente
if __name__ == '__main__':
    app.run(debug=True)


# TO DO:
# Modify text in template text boxes {A}
# Add "place to save" in the interface
