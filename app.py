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
    """Process uploaded Excel files using fixed Word template"""
    try:
        # Check if Excel files were uploaded
        if 'excel_files' not in request.files:
            flash('Pelo menos um arquivo Excel é obrigatório', 'error')
            return redirect(url_for('index'))
        
        excel_files = request.files.getlist('excel_files')
        
        # Check if files are selected
        if not excel_files or all(file.filename == '' for file in excel_files):
            flash('Por favor, selecione pelo menos um arquivo Excel', 'error')
            return redirect(url_for('index'))
        
        # Validate file extensions
        for excel_file in excel_files:
            if excel_file.filename and not allowed_file(excel_file.filename, ALLOWED_EXCEL_EXTENSIONS):
                flash(f'O arquivo {excel_file.filename} deve ter extensão .xlsx', 'error')
                return redirect(url_for('index'))
        
        # Create temporary directory for processing
        with tempfile.TemporaryDirectory() as temp_dir:
            # Use fixed Word template from project
            word_template_path = os.path.join('templates_word', 'modelo_padrao.docx')
            
            # Initialize processors
            excel_processor = ExcelProcessor()
            calc_engine = CalculationEngine()
            word_processor = WordProcessor()
            
            processed_files = []
            
            # Process each Excel file
            for excel_file in excel_files:
                if not excel_file.filename:
                    continue
                    
                # Save uploaded Excel file
                excel_filename = secure_filename(excel_file.filename)
                excel_path = os.path.join(temp_dir, excel_filename)
                excel_file.save(excel_path)
                
                # Process Excel file - extract all rows
                excel_result = excel_processor.extract_data(excel_path)
                
                if not excel_result:
                    flash(f'Erro ao processar arquivo {excel_filename}. Verifique se o arquivo possui dados nas colunas A, B, D, E, F, G a partir da linha 4.', 'warning')
                    continue
                
                # Get worksheet name and data
                excel_data_list = excel_result.get('data', [])
                worksheet_name = excel_result.get('worksheet_name', 'Planilha')
                
                # Process all rows with calculations
                calculated_data_list = []
                for row_data in excel_data_list:
                    calculated_row = calc_engine.process_row(row_data)
                    calculated_data_list.append(calculated_row)
                
                # Process Word file with all calculated data and worksheet name
                output_filename = f'resultado_{excel_filename.replace(".xlsx", ".docx")}'
                output_path = os.path.join(temp_dir, output_filename)
                
                success = word_processor.fill_template(word_template_path, calculated_data_list, output_path, worksheet_name)
                
                if success:
                    # Copy output file to uploads folder
                    final_output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
                    shutil.copy2(output_path, final_output_path)
                    processed_files.append((output_filename, final_output_path))
                else:
                    flash(f'Erro ao processar arquivo Word para {excel_filename}.', 'warning')
            
            # Check if any files were processed successfully
            if not processed_files:
                flash('Nenhum arquivo foi processado com sucesso.', 'error')
                return redirect(url_for('index'))
            
            # If only one file was processed, download it directly
            if len(processed_files) == 1:
                flash('Arquivo processado com sucesso!', 'success')
                return send_file(processed_files[0][1], 
                               as_attachment=True, 
                               download_name=processed_files[0][0])
            
            # If multiple files were processed, create a ZIP file
            import zipfile
            zip_filename = f'resultados_{len(processed_files)}_arquivos.zip'
            zip_path = os.path.join(temp_dir, zip_filename)
            
            with zipfile.ZipFile(zip_path, 'w') as zip_file:
                for filename, filepath in processed_files:
                    zip_file.write(filepath, filename)
            
            # Copy ZIP to uploads folder
            final_zip_path = os.path.join(app.config['UPLOAD_FOLDER'], zip_filename)
            shutil.copy2(zip_path, final_zip_path)
            
            flash(f'{len(processed_files)} arquivos processados com sucesso!', 'success')
            return send_file(final_zip_path, 
                           as_attachment=True, 
                           download_name=zip_filename)
    
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
