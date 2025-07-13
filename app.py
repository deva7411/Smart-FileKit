from flask import Flask, render_template, request, send_file
import os
import zipfile
from PyPDF2 import PdfMerger
from PIL import Image
from werkzeug.utils import secure_filename
import uuid
from docx import Document
from zipfile import ZipFile, ZIP_DEFLATED
from docx.shared import Inches
import shutil

app = Flask(__name__)

app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20 MB

UPLOAD_FOLDER = 'uploads'
MERGED_FOLDER = 'merged'
COMPRESS_FOLDER = 'compressed'
CONVERTED_FOLDER = 'converted'
MAX_FILE_SIZE_MB = 1
MAX_CONTENT_LENGTH_MB = 5

for folder in [UPLOAD_FOLDER, MERGED_FOLDER, COMPRESS_FOLDER, CONVERTED_FOLDER]:
    os.makedirs(folder, exist_ok=True)

app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH_MB * 1024 * 1024  # Total request size

# Home Page
@app.route('/')
def index():
    return render_template('index.html')
   #return render_template('home.html')

# --- MERGE PDFs ---
@app.route('/merge', methods=['GET'])
def merge_page():
    return render_template('merge.html')

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

# --- COMPRESS FILES ---
@app.route('/compress', methods=['GET'])
def compress_page():
    return render_template('compress.html')
@app.route('/compress_file', methods=['POST'])
def compress_file():
    file = request.files['file']
    target_kb = request.form.get('target_size')

    if not file or file.filename == '':
        return "No file selected", 400

    filename = secure_filename(file.filename)
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(file_path)

    file_ext = filename.rsplit('.', 1)[-1].lower()
    base_name = filename.rsplit('.', 1)[0]
    compressed_path = os.path.join(COMPRESS_FOLDER, base_name + '_compressed')

    # Convert target size to bytes
    try:
        target_bytes = int(target_kb) * 1024 if target_kb else None
    except:
        target_bytes = None

    # Compress image
    if file_ext in ['jpg', 'jpeg', 'png']:
        compressed_img_path = compressed_path + '.jpg'
        image = Image.open(file_path)

        for q in range(85, 10, -5):
            image.save(compressed_img_path, format='JPEG', quality=q)
            if not target_bytes or os.path.getsize(compressed_img_path) <= target_bytes:
                break

        return send_file(compressed_img_path, as_attachment=True)

    # Other file types: return original without zip
    return send_file(file_path, as_attachment=True)

    return send_file(zip_path, as_attachment=True)

# --- CONVERT TO PDF ---
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
            return "❌ DOCX to PDF is currently only supported on Windows.", 400

        else:
            return "❌ Invalid file type. Only image or Word files are supported.", 400

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return f"🔥 Error during conversion: {e}", 500



@app.route('/sitemap.xml')
def sitemap():
    return send_file('static/sitemap.xml', mimetype='application/xml')  
@app.route('/robots.txt')
def robots():
    return send_file('static/robots.txt', mimetype='text/plain')

@app.errorhandler(413)
def request_entity_too_large(error):
    return "❌ File too large. Maximum size is 20MB.", 413


if __name__ == '__main__':
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)

