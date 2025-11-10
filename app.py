import os
import zipfile
import tempfile
import shutil
from flask import Flask, render_template, request, send_file, jsonify, send_from_directory
from werkzeug.utils import secure_filename
import pandas as pd
from extract_shipping_bills import process_shipping_bills

app = Flask(__name__, static_folder="static")
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'

# Ensure folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf', 'zip'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Only PDF and ZIP files are allowed'}), 400
    
    try:
        # Create a temporary directory for processing
        temp_dir = tempfile.mkdtemp()
        pdf_dir = os.path.join(temp_dir, 'pdfs')
        os.makedirs(pdf_dir, exist_ok=True)
        
        filename = secure_filename(file.filename)
        file_path = os.path.join(temp_dir, filename)
        file.save(file_path)
        
        # Handle ZIP files
        if filename.lower().endswith('.zip'):
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                # Extract all PDF files
                for zip_info in zip_ref.filelist:
                    if zip_info.filename.lower().endswith('.pdf'):
                        # Extract to pdf_dir
                        zip_info.filename = os.path.basename(zip_info.filename)
                        zip_ref.extract(zip_info, pdf_dir)
        else:
            # Single PDF file
            shutil.copy(file_path, os.path.join(pdf_dir, filename))
        
        # Count PDF files
        pdf_files = [f for f in os.listdir(pdf_dir) if f.lower().endswith('.pdf')]
        num_files = len(pdf_files)
        
        if num_files == 0:
            shutil.rmtree(temp_dir)
            return jsonify({'error': 'No PDF files found in the uploaded file'}), 400
        
        # Process the PDFs
        df = process_shipping_bills(pdf_dir)
        
        # Generate unique output filenames
        import time
        timestamp = str(int(time.time()))
        output_csv = os.path.join(app.config['OUTPUT_FOLDER'], f'shipping_bill_{timestamp}.csv')
        output_excel = os.path.join(app.config['OUTPUT_FOLDER'], f'shipping_bill_{timestamp}.xlsx')
        
        # Save outputs
        df.to_csv(output_csv, index=False)
        df.to_excel(output_excel, index=False)
        
        # Clean up temp directory
        shutil.rmtree(temp_dir)
        
        return jsonify({
            'success': True,
            'files_processed': num_files,
            'rows_extracted': len(df),
            'csv_file': os.path.basename(output_csv),
            'excel_file': os.path.basename(output_excel)
        })
    
    except Exception as e:
        # Clean up on error
        if 'temp_dir' in locals():
            shutil.rmtree(temp_dir, ignore_errors=True)
        return jsonify({'error': f'Processing error: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        if not os.path.exists(file_path):
            return jsonify({'error': 'File not found'}), 404
        
        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run()
