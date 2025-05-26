import os
import logging
from flask import Flask, render_template, request, flash, redirect, url_for, send_file
from werkzeug.utils import secure_filename
import tempfile
import shutil
from utils.excel_processor import ExcelProcessor
from utils.word_processor import WordProcessor
from utils.calculations import CalculationEngine

# Configure logging
logging.basicConfig(level=logging.DEBUG)

app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "fallback-secret-key")

# Configuration
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXCEL_EXTENSIONS = {'xlsx'}
ALLOWED_WORD_EXTENSIONS = {'docx'}
MAX_FILE_SIZE = 16 * 1024 * 1024  # 16MB

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE

# Create upload folder if it doesn't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename, allowed_extensions):
    """Check if file has allowed extension"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in allowed_extensions

@app.route('/')
def index():
    """Main page for file upload and processing"""
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_files():
    """Process uploaded Excel and Word files"""
    try:
        # Check if files were uploaded
        if 'excel_file' not in request.files or 'word_file' not in request.files:
            flash('Ambos os arquivos (Excel e Word) são obrigatórios', 'error')
            return redirect(url_for('index'))
        
        excel_file = request.files['excel_file']
        word_file = request.files['word_file']
        
        # Check if files are selected
        if excel_file.filename == '' or word_file.filename == '':
            flash('Por favor, selecione ambos os arquivos', 'error')
            return redirect(url_for('index'))
        
        # Validate file extensions
        if not allowed_file(excel_file.filename, ALLOWED_EXCEL_EXTENSIONS):
            flash('O arquivo Excel deve ter extensão .xlsx', 'error')
            return redirect(url_for('index'))
        
        if not allowed_file(word_file.filename, ALLOWED_WORD_EXTENSIONS):
            flash('O arquivo Word deve ter extensão .docx', 'error')
            return redirect(url_for('index'))
        
        # Create temporary directory for processing
        with tempfile.TemporaryDirectory() as temp_dir:
            # Save uploaded files
            excel_filename = secure_filename(excel_file.filename)
            word_filename = secure_filename(word_file.filename)
            
            excel_path = os.path.join(temp_dir, excel_filename)
            word_path = os.path.join(temp_dir, word_filename)
            
            excel_file.save(excel_path)
            word_file.save(word_path)
            
            # Process Excel file
            excel_processor = ExcelProcessor()
            excel_data = excel_processor.extract_data(excel_path)
            
            if not excel_data:
                flash('Erro ao processar arquivo Excel. Verifique se os dados estão na linha 4.', 'error')
                return redirect(url_for('index'))
            
            # Perform calculations
            calc_engine = CalculationEngine()
            calculated_data = calc_engine.process_row(excel_data)
            
            # Process Word file
            word_processor = WordProcessor()
            output_path = os.path.join(temp_dir, 'output_' + word_filename)
            
            success = word_processor.fill_template(word_path, calculated_data, output_path)
            
            if not success:
                flash('Erro ao processar arquivo Word. Verifique se o template possui uma tabela.', 'error')
                return redirect(url_for('index'))
            
            # Copy output file to uploads folder for download
            final_output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'resultado_' + word_filename)
            shutil.copy2(output_path, final_output_path)
            
            flash('Arquivo processado com sucesso!', 'success')
            return send_file(final_output_path, 
                           as_attachment=True, 
                           download_name=f'resultado_{word_filename}')
    
    except Exception as e:
        app.logger.error(f"Erro durante processamento: {str(e)}")
        flash(f'Erro durante o processamento: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.errorhandler(413)
def too_large(e):
    """Handle file too large error"""
    flash('Arquivo muito grande. Tamanho máximo permitido: 16MB', 'error')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
