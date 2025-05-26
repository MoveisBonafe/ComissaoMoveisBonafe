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
    """Process uploaded Excel file using fixed Word template"""
    try:
        # Check if Excel file was uploaded
        if 'excel_file' not in request.files:
            flash('Arquivo Excel é obrigatório', 'error')
            return redirect(url_for('index'))
        
        excel_file = request.files['excel_file']
        
        # Check if file is selected
        if excel_file.filename == '':
            flash('Por favor, selecione o arquivo Excel', 'error')
            return redirect(url_for('index'))
        
        # Validate file extension
        if not allowed_file(excel_file.filename, ALLOWED_EXCEL_EXTENSIONS):
            flash('O arquivo Excel deve ter extensão .xlsx', 'error')
            return redirect(url_for('index'))
        
        # Create temporary directory for processing
        with tempfile.TemporaryDirectory() as temp_dir:
            # Save uploaded Excel file
            excel_filename = secure_filename(excel_file.filename)
            excel_path = os.path.join(temp_dir, excel_filename)
            excel_file.save(excel_path)
            
            # Use fixed Word template from project
            word_template_path = os.path.join('templates_word', 'modelo_padrao.docx')
            
            # Process Excel file - extract all rows
            excel_processor = ExcelProcessor()
            excel_result = excel_processor.extract_data(excel_path)
            
            if not excel_result:
                flash('Erro ao processar arquivo Excel. Verifique se existem dados válidos a partir da linha 4.', 'error')
                return redirect(url_for('index'))
            
            # Extract worksheet name and data
            worksheet_name = excel_result.get('worksheet_name', 'Planilha')
            excel_data_list = excel_result.get('data', [])
            
            # Perform calculations for each row
            calc_engine = CalculationEngine()
            calculated_data_list = []
            
            for row_data in excel_data_list:
                calculated_row = calc_engine.process_row(row_data)
                calculated_data_list.append(calculated_row)
            
            # Process Word file with all calculated data and worksheet name
            word_processor = WordProcessor()
            output_filename = f'resultado_{excel_filename.replace(".xlsx", ".docx")}'
            output_path = os.path.join(temp_dir, output_filename)
            
            success = word_processor.fill_template(word_template_path, calculated_data_list, output_path, worksheet_name)
            
            if not success:
                flash('Erro ao processar arquivo Word. Verifique se o template possui uma tabela.', 'error')
                return redirect(url_for('index'))
            
            # Copy output file to uploads folder for download
            final_output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
            shutil.copy2(output_path, final_output_path)
            
            flash('Arquivo processado com sucesso!', 'success')
            return send_file(final_output_path, 
                           as_attachment=True, 
                           download_name=output_filename)
    
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
