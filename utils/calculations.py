"""
Calculation engine for processing Excel data and computing commission values
"""

import logging
import re
from typing import Dict, Any, Optional

class CalculationEngine:
    """Class to handle calculations for commission processing"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def process_row(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Process a row of data and calculate the commission value
        
        Args:
            data: Dictionary containing the raw Excel data
            
        Returns:
            Dictionary with processed data including calculated commission
        """
        try:
            # Extract and validate input values
            valor_pedido = self._to_float(data.get('valor_pedido', 0))
            porcentagem = self._to_float(data.get('porcentagem', 0))
            prazo_raw = str(data.get('prazo', '')).strip()
            
            # Process prazo value according to rules
            prazo_value = self._process_prazo(prazo_raw)
            
            # Calculate commission
            valor_comissao = self._calculate_commission(valor_pedido, porcentagem, prazo_value)
            
            # Prepare processed data
            processed_data = {
                'data': self._format_date(data.get('data')),
                'numero_pedido': self._format_string(data.get('numero_pedido')),
                'nome_cliente': self._format_string(data.get('nome_cliente')),
                'prazo': self._format_string(data.get('prazo')),
                'valor_pedido': valor_pedido,
                'porcentagem': porcentagem,
                'valor_comissao': valor_comissao,
                'prazo_processed_value': prazo_value  # For debugging
            }
            
            self.logger.info(f"Processed data: {processed_data}")
            return processed_data
            
        except Exception as e:
            self.logger.error(f"Error processing row data: {str(e)}")
            # Return original data with zero commission on error
            return {
                'data': data.get('data', ''),
                'numero_pedido': data.get('numero_pedido', ''),
                'nome_cliente': data.get('nome_cliente', ''),
                'prazo': data.get('prazo', ''),
                'valor_pedido': self._to_float(data.get('valor_pedido', 0)),
                'porcentagem': self._to_float(data.get('porcentagem', 0)),
                'valor_comissao': 0,
                'prazo_processed_value': 0
            }
    
    def _process_prazo(self, prazo_str: str) -> float:
        """
        Process prazo string according to business rules
        
        Rules:
        - No "/" character: value = 0
        - Has "/" character: extract last number after "/"
          - >30 and <60: value = 2
          - >=60 and <90: value = 4
          - >=90 and <120: value = 5
          - >=120: value = 7
        
        Args:
            prazo_str: The prazo string from Excel
            
        Returns:
            Processed value according to rules
        """
        try:
            if not prazo_str or '/' not in prazo_str:
                self.logger.debug(f"No '/' found in prazo '{prazo_str}', returning 0")
                return 0
            
            # Split by '/' and get the last part
            parts = prazo_str.split('/')
            last_part = parts[-1].strip()
            
            # Extract numeric value from last part
            # Use regex to find numbers
            numbers = re.findall(r'\d+', last_part)
            if not numbers:
                self.logger.warning(f"No numbers found in last part '{last_part}' of prazo '{prazo_str}'")
                return 0
            
            # Take the first (or only) number found
            last_number = int(numbers[0])
            
            # Apply business rules
            if last_number > 30 and last_number < 60:
                value = 2
            elif last_number >= 60 and last_number < 90:
                value = 4
            elif last_number >= 90 and last_number < 120:
                value = 5
            elif last_number >= 120:
                value = 7
            else:
                value = 0
            
            self.logger.debug(f"Prazo '{prazo_str}' -> last_number: {last_number} -> value: {value}")
            return value
            
        except Exception as e:
            self.logger.error(f"Error processing prazo '{prazo_str}': {str(e)}")
            return 0
    
    def _calculate_commission(self, valor_pedido: float, porcentagem: float, prazo_value: float) -> float:
        """
        Calculate commission value using the formula:
        valor_pedido - percentage of (porcentagem + prazo_value)
        
        Args:
            valor_pedido: Order value
            porcentagem: Percentage value
            prazo_value: Processed prazo value
            
        Returns:
            Calculated commission value
        """
        try:
            if valor_pedido <= 0:
                return 0
            
            # Sum of porcentagem and prazo_value
            sum_percentages = porcentagem + prazo_value
            
            # Calculate percentage amount to subtract
            percentage_amount = valor_pedido * (sum_percentages / 100)
            
            # Calculate final commission
            commission = valor_pedido - percentage_amount
            
            self.logger.debug(f"Commission calculation: {valor_pedido} - ({sum_percentages}% of {valor_pedido}) = {commission}")
            
            return max(0, commission)  # Ensure non-negative result
            
        except Exception as e:
            self.logger.error(f"Error calculating commission: {str(e)}")
            return 0
    
    def _to_float(self, value: Any) -> float:
        """
        Convert value to float, handling various input types
        
        Args:
            value: Value to convert
            
        Returns:
            Float value or 0 if conversion fails
        """
        try:
            if value is None:
                return 0.0
            
            if isinstance(value, (int, float)):
                return float(value)
            
            if isinstance(value, str):
                # Remove common currency symbols and formatting
                cleaned = value.strip()
                cleaned = cleaned.replace('R$', '').replace('$', '')
                cleaned = cleaned.replace('.', '').replace(',', '.')  # Handle Brazilian format
                cleaned = re.sub(r'[^\d.-]', '', cleaned)  # Keep only digits, dots, and minus
                
                if cleaned:
                    return float(cleaned)
            
            return 0.0
            
        except Exception as e:
            self.logger.warning(f"Error converting '{value}' to float: {str(e)}")
            return 0.0
    
    def _format_date(self, date_value: Any) -> str:
        """
        Format date value for display
        
        Args:
            date_value: Date value to format
            
        Returns:
            Formatted date string
        """
        try:
            if date_value is None:
                return ""
            
            # If already a string, return as-is
            if isinstance(date_value, str):
                return date_value.strip()
            
            # If datetime object, format it
            if hasattr(date_value, 'strftime'):
                return date_value.strftime('%d/%m/%Y')
            
            return str(date_value)
            
        except Exception as e:
            self.logger.warning(f"Error formatting date '{date_value}': {str(e)}")
            return str(date_value) if date_value is not None else ""
    
    def _format_string(self, value: Any) -> str:
        """
        Format string value for display
        
        Args:
            value: Value to format
            
        Returns:
            Formatted string
        """
        try:
            if value is None:
                return ""
            return str(value).strip()
        except Exception as e:
            self.logger.warning(f"Error formatting string '{value}': {str(e)}")
            return ""
    
    def validate_data(self, data: Dict[str, Any]) -> bool:
        """
        Validate that the data contains required fields
        
        Args:
            data: Data dictionary to validate
            
        Returns:
            True if data is valid, False otherwise
        """
        required_fields = ['valor_pedido', 'porcentagem']
        
        for field in required_fields:
            if field not in data:
                self.logger.error(f"Missing required field: {field}")
                return False
            
            value = self._to_float(data[field])
            if value < 0:
                self.logger.error(f"Field {field} cannot be negative: {value}")
                return False
        
        return True
