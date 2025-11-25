"""
Flask web application for PDF to Excel conversion
"""

from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import os
from werkzeug.utils import secure_filename
from pdf_to_excel_extractor import PDFDataExtractor
import tempfile

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('No file part')
        return redirect(request.url)
    
    file = request.files['file']
    
    if file.filename == '':
        flash('No selected file')
        return redirect(request.url)
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        try:
            # Extract data
            extractor = PDFDataExtractor(filepath)
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'output.xlsx')
            extractor.save_to_excel(output_path)
            
            # Clean up input file
            os.remove(filepath)
            
            return send_file(
                output_path,
                as_attachment=True,
                download_name='extracted_data.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        except Exception as e:
            flash(f'Error processing file: {str(e)}')
            return redirect(url_for('index'))
    
    flash('Invalid file type. Please upload a PDF file.')
    return redirect(url_for('index'))

@app.route('/demo')
def demo():
    """Demo with the provided sample file"""
    try:
        extractor = PDFDataExtractor('Data Input.pdf')
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'demo_output.xlsx')
        extractor.save_to_excel(output_path)
        
        return send_file(
            output_path,
            as_attachment=True,
            download_name='demo_output.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        flash(f'Error in demo: {str(e)}')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001)
