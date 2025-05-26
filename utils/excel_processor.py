"""
Excel file processing utilities for extracting data from specific cells
"""

import logging
from openpyxl import load_workbook
from typing import Dict, Any, Optional

class ExcelProcessor:
    """Class to handle Excel file processing and data extraction"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def extract_data(self, file_path: str) -> Optional[Dict[str, Any]]:
        """
        Extract data from Excel file row 4, columns A, B, D, E, F, G
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            Dictionary with extracted data or None if error
        """
        try:
            # Load workbook
            workbook = load_workbook(file_path, data_only=True)
            
            # Get the first worksheet
            worksheet = workbook.active
            
            # Extract data from row 4, specific columns
            data = {
                'data': self._get_cell_value(worksheet, 'A4'),        # Column A - Data
                'numero_pedido': self._get_cell_value(worksheet, 'B4'), # Column B - Número do Pedido
                'nome_cliente': self._get_cell_value(worksheet, 'D4'),  # Column D - Nome do Cliente
                'prazo': self._get_cell_value(worksheet, 'E4'),         # Column E - Prazo
                'valor_pedido': self._get_cell_value(worksheet, 'F4'),  # Column F - Valor do Pedido
                'porcentagem': self._get_cell_value(worksheet, 'G4')    # Column G - Porcentagem
            }
            
            # Validate that we have at least some data
            non_empty_values = [v for v in data.values() if v is not None and str(v).strip()]
            if len(non_empty_values) == 0:
                self.logger.warning("No data found in row 4 of Excel file")
                return None
            
            # Log extracted data for debugging
            self.logger.info(f"Extracted data: {data}")
            
            return data
            
        except Exception as e:
            self.logger.error(f"Error extracting data from Excel file: {str(e)}")
            return None
    
    def _get_cell_value(self, worksheet, cell_address: str) -> Any:
        """
        Get value from a specific cell, handling different data types
        
        Args:
            worksheet: The worksheet object
            cell_address: Cell address (e.g., 'A4')
            
        Returns:
            Cell value or None if empty
        """
        try:
            cell = worksheet[cell_address]
            value = cell.value
            
            # Handle None values
            if value is None:
                return None
            
            # Handle datetime objects - convert to string
            if hasattr(value, 'strftime'):
                return value.strftime('%d/%m/%Y')
            
            # Handle numeric values
            if isinstance(value, (int, float)):
                return value
            
            # Handle string values - strip whitespace
            if isinstance(value, str):
                stripped = value.strip()
                return stripped if stripped else None
            
            # Return as-is for other types
            return value
            
        except Exception as e:
            self.logger.warning(f"Error getting value from cell {cell_address}: {str(e)}")
            return None
    
    def validate_excel_structure(self, file_path: str) -> bool:
        """
        Validate that the Excel file has the expected structure
        
        Args:
            file_path: Path to the Excel file
            
        Returns:
            True if structure is valid, False otherwise
        """
        try:
            workbook = load_workbook(file_path, data_only=True)
            worksheet = workbook.active
            
            # Check if worksheet has at least 4 rows
            if worksheet.max_row < 4:
                self.logger.error("Excel file must have at least 4 rows")
                return False
            
            # Check if worksheet has at least 7 columns (A-G)
            if worksheet.max_column < 7:
                self.logger.error("Excel file must have at least 7 columns (A-G)")
                return False
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error validating Excel structure: {str(e)}")
            return False
    
    def get_preview_data(self, file_path: str, max_rows: int = 10) -> Optional[Dict[str, Any]]:
        """
        Get preview data from Excel file for display purposes
        
        Args:
            file_path: Path to the Excel file
            max_rows: Maximum number of rows to preview
            
        Returns:
            Dictionary with preview data or None if error
        """
        try:
            workbook = load_workbook(file_path, data_only=True)
            worksheet = workbook.active
            
            # Get headers (assuming row 1 contains headers)
            headers = []
            for col in range(1, min(8, worksheet.max_column + 1)):  # Columns A-G
                cell_value = worksheet.cell(row=1, column=col).value
                headers.append(str(cell_value) if cell_value else f"Col {col}")
            
            # Get data rows
            rows = []
            start_row = 2  # Start from row 2 (after headers)
            end_row = min(start_row + max_rows, worksheet.max_row + 1)
            
            for row in range(start_row, end_row):
                row_data = []
                for col in range(1, len(headers) + 1):
                    cell_value = worksheet.cell(row=row, column=col).value
                    row_data.append(str(cell_value) if cell_value else "")
                rows.append(row_data)
            
            return {
                'headers': headers,
                'rows': rows,
                'total_rows': worksheet.max_row,
                'target_row': 4,
                'target_data': self.extract_data(file_path)
            }
            
        except Exception as e:
            self.logger.error(f"Error getting preview data: {str(e)}")
            return None
