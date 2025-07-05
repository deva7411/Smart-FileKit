from flask import Flask, render_template, request, send_file
import os
import zipfile
from PyPDF2 import PdfMerger
from PIL import Image
from werkzeug.utils import secure_filename
from PIL import Image
from docx2pdf import convert as docx_convert    # For converting DOCX to PDF
import uuid
from docx import Document
from zipfile import ZipFile, ZIP_DEFLATED
from docx.shared import Inches
from PIL import Image
import os, uuid, shutil
import pythoncom
import win32com.client
import comtypes.client
import pythoncom


app = Flask(__name__)

app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20 MB

UPLOAD_FOLDER = 'uploads'
MERGED_FOLDER = 'merged'
MAX_FILE_SIZE_MB = 1
MAX_CONTENT_LENGTH_MB = 5
COMPRESS_FOLDER = 'compressed'
CONVERTED_FOLDER = 'converted' 
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(MERGED_FOLDER, exist_ok=True)
os.makedirs(COMPRESS_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)


app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH_MB * 1024 * 1024  # Total request size

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(MERGED_FOLDER, exist_ok=True)


def docx_convert(input_path, output_path):
    pythoncom.CoInitialize()  # FIX: this is required
    try:
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(input_path)
        doc.SaveAs(output_path, FileFormat=17)  # 17 = PDF
        doc.Close()
        word.Quit()
    finally:
        pythoncom.CoUninitialize() 
# Home page showing tools
@app.route('/')
def index():
    return render_template('index.html')

# Form page for merging PDFs (GET)
@app.route('/merge', methods=['GET'])
def merge_page():
    return render_template('merge.html')  # Make sure merge.html exists in templates/

# Logic to merge PDFs (POST)
@app.route('/merge', methods=['POST'])
def merge_files():
    files = request.files.getlist('pdfs')

    if len(files) < 2:
        return "❌ Please upload at least two PDF files.", 400

    for file in files:
        if not file.filename.endswith('.pdf'):
            return f"❌ Invalid file type: {file.filename}", 400
        if file.content_length and file.content_length > MAX_FILE_SIZE_MB * 1024 * 1024:
            return f"❌ File {file.filename} exceeds {MAX_FILE_SIZE_MB}MB limit.", 400

    merger = PdfMerger()

    for file in files:
        filename = secure_filename(file.filename)
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(file_path)
        merger.append(file_path)

    merged_pdf_path = os.path.join(MERGED_FOLDER, 'merged_output.pdf')
    merger.write(merged_pdf_path)
    merger.close()

    return send_file(merged_pdf_path, as_attachment=True)


# Form page for compressing files (POST)
@app.route('/compress', methods=['GET'])
def compress_page():
    return render_template('compress.html')

@app.route('/compress_file', methods=['POST'])
def compress_file():
    file = request.files['file']
    target_kb = request.form.get('target_size')

    if file.filename == '':
        return "No file selected", 400

    filename = secure_filename(file.filename)
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(file_path)

    file_ext = filename.rsplit('.', 1)[-1].lower()
    base_name = filename.rsplit('.', 1)[0]
    compressed_path = os.path.join(COMPRESS_FOLDER, base_name + '_compressed')

    # Convert target_kb to bytes
    if target_kb:
        try:
            target_bytes = int(target_kb) * 1024
        except:
            target_bytes = None
    else:
        target_bytes = None

    # Handle image compression
    if file_ext in ['jpg', 'jpeg', 'png']:
        compressed_img_path = compressed_path + '.jpg'
        image = Image.open(file_path)

        quality = 85
        if target_bytes:
            # Try reducing quality until under target
            for q in range(85, 10, -5):
                image.save(compressed_img_path, format='JPEG', quality=q)
                if os.path.getsize(compressed_img_path) <= target_bytes:
                    break
        else:
            image.save(compressed_img_path, format='JPEG', quality=85)

        return send_file(compressed_img_path, as_attachment=True)

    # If not image — just zip it
    zip_path = compressed_path + '.zip'
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(file_path, arcname=filename)

    return send_file(zip_path, as_attachment=True)




# Form page for converting files (GET)
@app.route('/convert', methods=['GET'])
def convert_page():
    return render_template('convert.html')

@app.route('/convert', methods=['POST'])
def to_pdf():
    try:
        file = request.files['file']
        file_type = request.form['type']

        if file.filename == '':
            return "❌ No file selected", 400

        filename = secure_filename(file.filename)
        input_path = os.path.abspath(os.path.join(UPLOAD_FOLDER, filename))
        file.save(input_path)

        output_filename = str(uuid.uuid4()) + ".pdf"
        output_path = os.path.abspath(os.path.join(CONVERTED_FOLDER, output_filename))

        if file_type == 'image':
            image = Image.open(input_path)
            image = image.convert('RGB')
            image.save(output_path, 'PDF')

        elif file_type == 'word':
            docx_convert(input_path, output_path)

        else:
            return "❌ Invalid file type. Only image or Word files are supported.", 400

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return f"🔥 Error during conversion: {e}", 500




if __name__ == '__main__':
    app.run(debug=True)
    
