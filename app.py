from flask import Flask, render_template, request, send_file, redirect, url_for
import os
from pdf2docx import Converter
import tabula
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted'
ALLOWED_EXTENSIONS = {'pdf'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['CONVERTED_FOLDER'] = CONVERTED_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/pdf_to_word', methods=['GET', 'POST'])
def pdf_to_word():
    if request.method == 'POST':
        if 'file' not in request.files:
            return render_template('pdf_to_word.html', error='No file part')

        file = request.files['file']
        if file.filename == '':
            return render_template('pdf_to_word.html', error='No selected file')

        if file and allowed_file(file.filename):
            # Save the uploaded file
            filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filename)

            # Convert PDF to Word
            word_filename = os.path.join(app.config['CONVERTED_FOLDER'], file.filename.replace('.pdf', '.docx'))
            cv = Converter(filename)
            cv.convert(word_filename, start=0, end=None)
            cv.close()

            # Return the converted Word file for download
            return send_file(word_filename, as_attachment=True)

    return render_template('pdf_to_word.html')

@app.route('/download/<path:filename>')
def download_file(filename):
    return send_file(filename, as_attachment=True)

def format_excel(ws, df):
    for r_idx, row in enumerate(df.iterrows(), start=1):
        for c_idx, value in enumerate(row[1], start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.font = Font(name='Arial', size=10)
            if r_idx == 1:  # Header row
                cell.fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")  # Yellow background for header
                cell.font = Font(bold=True)
            else:
                if c_idx % 2 == 0:  # Alternate row color
                    cell.fill = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")

@app.route('/pdf_to_excel', methods=['GET', 'POST'])
def pdf_to_excel():
    if request.method == 'POST':
        if 'file' not in request.files:
            return render_template('pdf_to_excel.html', error='No file part')

        file = request.files['file']
        if file.filename == '':
            return render_template('pdf_to_excel.html', error='No selected file')

        if file and allowed_file(file.filename):
            # Save the uploaded file
            filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filename)

            # Convert PDF to CSV
            csv_filename = os.path.join(app.config['CONVERTED_FOLDER'], file.filename.replace('.pdf', '.csv'))
            try:
                tabula.convert_into(filename, csv_filename, output_format="csv", pages='all')
            except Exception as e:
                return render_template('pdf_to_excel.html', error=f'Error converting PDF to CSV: {str(e)}')

            # Read the CSV data into a pandas DataFrame
            df = pd.read_csv(csv_filename)

            # Create a new Excel workbook and sheet
            excel_filename = os.path.join(app.config['CONVERTED_FOLDER'], file.filename.replace('.pdf', '.xlsx'))
            wb = Workbook()
            ws = wb.active
            ws.title = 'Data'

            # Write DataFrame to Excel
            for r_idx, row in df.iterrows():
                for c_idx, value in enumerate(row, start=1):
                    cell = ws.cell(row=r_idx+2, column=c_idx, value=value)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.font = Font(name='Arial', size=10)
                    if r_idx == 0:  # Header row
                        cell.fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
                        cell.font = Font(bold=True)
                    else:
                        if r_idx % 2 == 0:
                            cell.fill = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")

            # Save the Excel file
            wb.save(excel_filename)

            # Return the converted Excel file for download
            return send_file(excel_filename, as_attachment=True)

    return render_template('pdf_to_excel.html')

if __name__ == '__main__':
    app.run(debug=True)
