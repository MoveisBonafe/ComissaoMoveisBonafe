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
            
            # Process frete value (column I from Excel, becomes column 8 in Word)
            frete_raw = self._to_float(data.get('frete', 0))
            # Convert from decimal to percentage (0.05 -> 5)
            frete_value = abs(frete_raw) * 100
            
            # Calculate referencia_comissao (column 9 in Word)
            # This is frete_value% of valor_comissao
            # Example: if valor_comissao=1000 and frete_value=5, then 5% of 1000 = 50
            referencia_comissao = valor_comissao * (frete_value / 100)
            
            # Format prazo for display (handle multiple slashes)
            prazo_formatted = self._format_prazo_display(str(data.get('prazo', '')).strip())
            
            # Prepare processed data
            processed_data = {
                'data': self._format_date(data.get('data')),
                'numero_pedido': self._format_string(data.get('numero_pedido')),
                'nome_cliente': self._format_string(data.get('nome_cliente')),
                'prazo': prazo_formatted,
                'valor_pedido': valor_pedido,
                'porcentagem': porcentagem,
                'valor_comissao': valor_comissao,
                'frete': frete_value,  # Column 8 - without % symbol
                'referencia_comissao': referencia_comissao,  # Column 9 - calculated
                'pagamento': 'BOLETOS',  # Column 10 - fixed text
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
          - >30 and <60: value = 4  (updated rule)
          - >=60 and <90: value = 4
          - >=90 and <120: value = 5
          - >=120: value = 7
        
        Args:
            prazo_str: The prazo string from Excel
            
        Returns:
            Processed value according to rules (negative for calculation)
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
            
            # Apply business rules - return negative values for calculation
            if last_number > 30 and last_number < 60:
                value = -4  # Example: 30/60 = -4
            elif last_number >= 60 and last_number < 90:
                value = -4
            elif last_number >= 90 and last_number < 120:
                value = -5
            elif last_number >= 120:
                value = -7
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
        
        Example: valor_pedido=1000, porcentagem=-7, prazo_value=-4 (from 30/60)
        sum = -7 + (-4) = -11
        commission = 1000 - (-11% of 1000) = 1000 - (-110) = 1000 + 110 = 1110
        But since we want subtraction: 1000 - 11% = 890
        
        Args:
            valor_pedido: Order value
            porcentagem: Percentage value (negative)
            prazo_value: Processed prazo value (negative)
            
        Returns:
            Calculated commission value
        """
        try:
            if valor_pedido <= 0:
                return 0
            
            # Sum of absolute values for percentage calculation
            # Since both values are negative, we need their absolute sum for the discount
            abs_porcentagem = abs(porcentagem)
            abs_prazo = abs(prazo_value)
            total_discount_percentage = abs_porcentagem + abs_prazo
            
            # Calculate percentage amount to subtract
            percentage_amount = valor_pedido * (total_discount_percentage / 100)
            
            # Calculate final commission (subtract the discount)
            commission = valor_pedido - percentage_amount
            
            self.logger.debug(f"Commission calculation: {valor_pedido} - ({total_discount_percentage}% of {valor_pedido}) = {commission}")
            
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
        Format date value for display in dd/mm format
        
        Args:
            date_value: Date value to format
            
        Returns:
            Formatted date string in dd/mm format
        """
        try:
            if date_value is None:
                return ""
            
            # If already a string, try to convert to dd/mm format
            if isinstance(date_value, str):
                date_str = date_value.strip()
                # If it's in dd/mm/yyyy format, extract dd/mm
                if '/' in date_str:
                    parts = date_str.split('/')
                    if len(parts) >= 2:
                        return f"{parts[0]}/{parts[1]}"
                return date_str
            
            # If datetime object, format it to dd/mm
            if hasattr(date_value, 'strftime'):
                return date_value.strftime('%d/%m')
            
            return str(date_value)
            
        except Exception as e:
            self.logger.warning(f"Error formatting date '{date_value}': {str(e)}")
            return str(date_value) if date_value is not None else ""
    
    def _format_string(self, value: Any) -> str:
        """
        Format string value for display with proper capitalization and character limit
        
        Args:
            value: Value to format
            
        Returns:
            Formatted string with first letter capitalized after each space, limited to 37 characters
        """
        try:
            if value is None:
                return ""
            
            text = str(value).strip()
            
            # Convert to title case (first letter of each word capitalized)
            # This will convert "PEDRO HENRIQUE" to "Pedro Henrique"
            formatted_text = text.title()
            
            # Limit to 37 characters for column 3 (nome do cliente)
            if len(formatted_text) > 37:
                formatted_text = formatted_text[:37]
            
            return formatted_text
        except Exception as e:
            self.logger.warning(f"Error formatting string '{value}': {str(e)}")
            return ""
    
    def _format_prazo_display(self, prazo_str: str) -> str:
        """
        Format prazo for display handling multiple slashes
        
        Rules:
        - If exactly 3 slashes (4 parts): show "first a last" (e.g., "30/60/90/120" -> "30 a 120")
        - If 1-2 slashes: show as-is (e.g., "30/60/90" stays "30/60/90")
        - If no slash: show as-is
        
        Args:
            prazo_str: The prazo string from Excel
            
        Returns:
            Formatted prazo string for display
        """
        try:
            if not prazo_str or '/' not in prazo_str:
                return prazo_str
            
            # Count the number of slashes
            slash_count = prazo_str.count('/')
            
            # Only format if there are exactly 3 slashes (4 parts)
            if slash_count == 3:
                # Split by '/' to get parts
                parts = prazo_str.split('/')
                
                # Take first part (index 0) and last part (index -1)
                first_part = parts[0].strip()
                last_part = parts[-1].strip()
                
                # Extract numbers from first and last parts
                first_numbers = re.findall(r'\d+', first_part)
                last_numbers = re.findall(r'\d+', last_part)
                
                if first_numbers and last_numbers:
                    first_num = first_numbers[0]
                    last_num = last_numbers[0]
                    formatted = f"{first_num} a {last_num}"
                    self.logger.debug(f"Formatted prazo '{prazo_str}' -> parts: {parts} -> first: '{first_part}' last: '{last_part}' -> '{formatted}'")
                    return formatted
            
            # For other cases (1-2 slashes), return as-is
            self.logger.debug(f"Prazo '{prazo_str}' ({slash_count} slashes) -> keeping as-is")
            return prazo_str
            
        except Exception as e:
            self.logger.warning(f"Error formatting prazo display '{prazo_str}': {str(e)}")
            return prazo_str
    
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
