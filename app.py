import os
from flask import Flask, render_template, request, send_file, redirect
from werkzeug.utils import secure_filename
from docx import Document
from pdf2docx import Converter

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'

# Ensure folders exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'pdf_file' not in request.files:
        return redirect('/')

    file = request.files['pdf_file']
    if file.filename == '':
        return redirect('/')

    filename = secure_filename(file.filename)
    input_path = os.path.join(UPLOAD_FOLDER, filename)
    output_path = os.path.join(OUTPUT_FOLDER, filename.replace('.pdf', '.docx'))

    file.save(input_path)

    # Convert PDF to DOCX using pdf2docx
    cv = Converter(input_path)
    cv.convert(output_path, start=0, end=None)
    cv.close()

    return send_file(output_path, as_attachment=True)

# ✅ Proper host + port for Railway, Render, etc.
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
