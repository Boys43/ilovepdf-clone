
from flask import Flask, render_template, request, send_file, redirect, url_for
from werkzeug.utils import secure_filename
import os
from pdf2docx import Converter
from PyPDF2 import PdfMerger

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/pdf-to-word')
def pdf_to_word_form():
    return render_template('pdf_to_word.html')

@app.route('/merge-pdf')
def merge_pdf_form():
    return render_template('merge_pdf.html')

@app.route('/convert-pdf-to-word', methods=['POST'])
def convert_pdf_to_word():
    file = request.files['pdf']
    filename = secure_filename(file.filename)
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    output_filename = filename.replace('.pdf', '.docx')
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

    file.save(input_path)

    try:
        cv = Converter(input_path)
        cv.convert(output_path, start=0, end=None)
        cv.close()
    except Exception as e:
        return f'Error: {e}', 500

    return send_file(output_path, as_attachment=True)

@app.route('/merge-pdfs', methods=['POST'])
def merge_pdfs():
    files = request.files.getlist('pdfs')
    merger = PdfMerger()

    input_paths = []
    for file in files:
        filename = secure_filename(file.filename)
        path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(path)
        input_paths.append(path)
        merger.append(path)

    output_path = os.path.join(app.config['OUTPUT_FOLDER'], 'merged.pdf')
    merger.write(output_path)
    merger.close()

    # Cleanup
    for p in input_paths:
        os.remove(p)

    return send_file(output_path, as_attachment=True)

import os

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
