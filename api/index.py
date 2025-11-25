#!/usr/bin/env python3
"""
Vercel serverless function for PDF to Excel conversion
"""

from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify
import os
import sys
from werkzeug.utils import secure_filename
import tempfile
import io

# Add parent directory to path to import our extractor
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from pdf_to_excel_extractor import PDFDataExtractor

app = Flask(__name__, template_folder='../templates')
app.secret_key = os.environ.get('SECRET_KEY', 'your-secret-key-here')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if file and allowed_file(file.filename):
        try:
            # Save uploaded file to temp location
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_input:
                file.save(tmp_input.name)
                input_path = tmp_input.name
            
            # Extract data
            extractor = PDFDataExtractor(input_path)
            
            # Save to temp Excel file
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_output:
                output_path = tmp_output.name
            
            extractor.save_to_excel(output_path)
            
            # Read the Excel file into memory
            with open(output_path, 'rb') as f:
                excel_data = io.BytesIO(f.read())
            
            # Clean up temp files
            os.unlink(input_path)
            os.unlink(output_path)
            
            excel_data.seek(0)
            return send_file(
                excel_data,
                as_attachment=True,
                download_name='extracted_data.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        except Exception as e:
            return jsonify({'error': f'Error processing file: {str(e)}'}), 500
    
    return jsonify({'error': 'Invalid file type. Please upload a PDF file.'}), 400

@app.route('/demo')
def demo():
    """Demo with the provided sample file"""
    try:
        # Use the sample file from project root
        sample_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'Data Input.pdf')
        
        if not os.path.exists(sample_path):
            return jsonify({'error': 'Demo file not found'}), 404
        
        extractor = PDFDataExtractor(sample_path)
        
        # Save to temp Excel file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_output:
            output_path = tmp_output.name
        
        extractor.save_to_excel(output_path)
        
        # Read the Excel file into memory
        with open(output_path, 'rb') as f:
            excel_data = io.BytesIO(f.read())
        
        # Clean up temp file
        os.unlink(output_path)
        
        excel_data.seek(0)
        return send_file(
            excel_data,
            as_attachment=True,
            download_name='demo_output.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': f'Error in demo: {str(e)}'}), 500

# Vercel serverless function handler
def handler(request):
    with app.request_context(request.environ):
        return app.full_dispatch_request()
