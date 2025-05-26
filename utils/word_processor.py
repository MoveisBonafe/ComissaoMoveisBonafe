"""
Word document processing utilities for filling templates with data
"""

import logging
from docx import Document
from typing import Dict, Any, List

class WordProcessor:
    """Class to handle Word document template processing"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def fill_template(self, template_path: str, data: Dict[str, Any], output_path: str) -> bool:
        """
        Fill Word template with provided data
        
        Args:
            template_path: Path to the Word template file
            data: Dictionary containing the data to fill
            output_path: Path where the filled document will be saved
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Load the document
            doc = Document(template_path)
            
            # Find the first table in the document
            if not doc.tables:
                self.logger.error("No tables found in Word document")
                return False
            
            table = doc.tables[0]  # Use the first table
            
            # Ensure table has at least 2 rows
            if len(table.rows) < 2:
                self.logger.error("Table must have at least 2 rows")
                return False
            
            # Ensure table has at least 7 columns
            if len(table.columns) < 7:
                self.logger.error("Table must have at least 7 columns")
                return False
            
            # Fill row 2 (index 1) with data
            row = table.rows[1]  # Row 2 (0-indexed)
            
            # Map data to table columns
            self._fill_cell(row.cells[0], data.get('data'))           # Column 1 - Data
            self._fill_cell(row.cells[1], data.get('numero_pedido'))  # Column 2 - Número do Pedido
            self._fill_cell(row.cells[2], data.get('nome_cliente'))   # Column 3 - Nome do Cliente
            self._fill_cell(row.cells[3], data.get('prazo'))          # Column 4 - Prazo
            self._fill_cell(row.cells[4], data.get('valor_pedido'))   # Column 5 - Valor do Pedido
            self._fill_cell(row.cells[5], data.get('porcentagem'))    # Column 6 - Porcentagem
            self._fill_cell(row.cells[6], data.get('valor_comissao')) # Column 7 - Valor da Comissão (calculated)
            
            # Save the filled document
            doc.save(output_path)
            
            self.logger.info(f"Successfully filled template and saved to {output_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error filling Word template: {str(e)}")
            return False
    
    def _fill_cell(self, cell, value: Any) -> None:
        """
        Fill a table cell with the provided value
        
        Args:
            cell: The table cell object
            value: The value to insert
        """
        try:
            # Clear existing content
            cell.text = ""
            
            # Add new content
            if value is not None:
                # Format numeric values
                if isinstance(value, (int, float)):
                    # Format currency values
                    if abs(value) >= 0.01:  # Avoid very small numbers
                        formatted_value = f"R$ {value:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                    else:
                        formatted_value = f"{value:.4f}"
                    cell.text = formatted_value
                else:
                    cell.text = str(value)
            else:
                cell.text = ""
                
        except Exception as e:
            self.logger.warning(f"Error filling cell: {str(e)}")
            cell.text = str(value) if value is not None else ""
    
    def get_table_info(self, template_path: str) -> Dict[str, Any]:
        """
        Get information about tables in the Word document
        
        Args:
            template_path: Path to the Word template file
            
        Returns:
            Dictionary with table information
        """
        try:
            doc = Document(template_path)
            
            tables_info = []
            for i, table in enumerate(doc.tables):
                table_info = {
                    'index': i,
                    'rows': len(table.rows),
                    'columns': len(table.columns),
                    'headers': []
                }
                
                # Try to extract headers from first row
                if table.rows:
                    first_row = table.rows[0]
                    for cell in first_row.cells:
                        table_info['headers'].append(cell.text.strip())
                
                tables_info.append(table_info)
            
            return {
                'total_tables': len(doc.tables),
                'tables': tables_info,
                'has_suitable_table': any(
                    t['rows'] >= 2 and t['columns'] >= 7 
                    for t in tables_info
                )
            }
            
        except Exception as e:
            self.logger.error(f"Error getting table info: {str(e)}")
            return {'total_tables': 0, 'tables': [], 'has_suitable_table': False}
    
    def validate_template(self, template_path: str) -> bool:
        """
        Validate that the Word template has the required structure
        
        Args:
            template_path: Path to the Word template file
            
        Returns:
            True if template is valid, False otherwise
        """
        try:
            doc = Document(template_path)
            
            # Check if document has tables
            if not doc.tables:
                self.logger.error("Template must contain at least one table")
                return False
            
            # Check first table structure
            table = doc.tables[0]
            
            if len(table.rows) < 2:
                self.logger.error("Table must have at least 2 rows (header + data)")
                return False
            
            if len(table.columns) < 7:
                self.logger.error("Table must have at least 7 columns")
                return False
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error validating template: {str(e)}")
            return False
    
    def create_sample_template(self, output_path: str) -> bool:
        """
        Create a sample Word template for testing purposes
        
        Args:
            output_path: Path where the sample template will be saved
            
        Returns:
            True if successful, False otherwise
        """
        try:
            doc = Document()
            
            # Add title
            title = doc.add_heading('Pedidos - Comissão 5% - mês', 0)
            
            # Add table
            table = doc.add_table(rows=2, cols=7)
            table.style = 'Table Grid'
            
            # Fill header row
            headers = [
                'Data', 'Nº Pedido', 'Nome do Cliente', 'Prazo p/ pagam.',
                'Valor Pedido', 'Porcentagem', 'Valor da Comissão'
            ]
            
            for i, header in enumerate(headers):
                table.cell(0, i).text = header
            
            # Add empty data row
            for i in range(7):
                table.cell(1, i).text = ""
            
            # Save document
            doc.save(output_path)
            
            self.logger.info(f"Sample template created at {output_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error creating sample template: {str(e)}")
            return False
