"""
Document Parser Service
Extracts variables and data from Excel files with ##variable## format.
"""
import pandas as pd
import re
from typing import List, Dict, Tuple
from pathlib import Path


class DocumentParser:
    """Parses Excel files to extract variables and data."""
    
    VARIABLE_PATTERN = re.compile(r'##([^#]+)##')
    
    def __init__(self):
        """Initialize DocumentParser."""
        pass
    
    def extract_variables(self, text: str) -> List[str]:
        """
        Extract variable names from text with ##variable## format.
        
        Args:
            text: Text containing variables
            
        Returns:
            List of variable names (without ## markers)
        """
        if not isinstance(text, str):
            return []
        return self.VARIABLE_PATTERN.findall(text)
    
    def parse_excel_data(self, file_path: str, sheet_name: str = None) -> Dict:
        """
        Parse Excel file to extract variables and data.
        Variables are expected in the first row with ##variable## format.
        
        Args:
            file_path: Path to Excel file
            sheet_name: Specific sheet name (None for first sheet)
            
        Returns:
            Dictionary with structure:
            {
                'variables': List of variable names,
                'data': List of dictionaries (one per row),
                'raw_headers': Original first row values,
                'sheet_name': Sheet name used
            }
            
        Raises:
            FileNotFoundError: If file doesn't exist
            ValueError: If file format is invalid
        """
        if not Path(file_path).exists():
            raise FileNotFoundError(f"Excel file not found: {file_path}")
        
        try:
            import openpyxl
            from openpyxl.utils import get_column_letter
            
            # Load workbook with data_only=False to preserve formulas and formatting
            wb = openpyxl.load_workbook(file_path, data_only=True)
            
            # Get the specified sheet or the first one
            if sheet_name:
                if sheet_name not in wb.sheetnames:
                    raise ValueError(f"Sheet '{sheet_name}' not found in workbook")
                ws = wb[sheet_name]
            else:
                ws = wb.active
                sheet_name = ws.title
            
            # Get all rows
            rows = list(ws.iter_rows(values_only=False))
            if len(rows) < 2:
                raise ValueError("Excel file must have at least 2 rows (header + data)")
            
            # Extract first row (headers with ##variable##)
            header_row = rows[0]
            raw_headers = []
            variables = []
            
            for cell in header_row:
                header_value = cell.value if cell.value else ""
                raw_headers.append(str(header_value))
                
                if header_value:
                    extracted = self.extract_variables(str(header_value))
                    if extracted:
                        variables.append(extracted[0])
                    else:
                        # If no ##variable## format, use the raw value
                        variables.append(str(header_value))
                else:
                    variables.append(f"column_{len(variables)}")
            
            # Parse data rows (skip first row which is header)
            data_rows = []
            for row in rows[1:]:
                row_data = {}
                for col_idx, cell in enumerate(row):
                    if col_idx >= len(variables):
                        break
                    
                    var_name = variables[col_idx]
                    
                    # Handle empty cells
                    if cell.value is None:
                        row_data[var_name] = ""
                        continue
                    
                    # Format cell value based on its number format
                    formatted_value = self._format_cell_value(cell)
                    row_data[var_name] = formatted_value
                
                data_rows.append(row_data)
            
            wb.close()
            
            return {
                'variables': variables,
                'data': data_rows,
                'raw_headers': raw_headers,
                'sheet_name': sheet_name,
                'total_rows': len(data_rows)
            }
            
        except Exception as e:
            raise ValueError(f"Error parsing Excel file: {str(e)}")
    
    def _format_cell_value(self, cell):
        """
        Format cell value based on its number format to preserve percentages, currency, dates, etc.
        
        Args:
            cell: openpyxl cell object
            
        Returns:
            Formatted string value
        """
        from datetime import datetime
        
        value = cell.value
        number_format = cell.number_format
        
        # Handle dates
        if isinstance(value, datetime):
            return value.strftime('%Y-%m-%d %H:%M:%S') if value.hour or value.minute or value.second else value.strftime('%Y-%m-%d')
        
        # Handle percentages
        if number_format and '%' in number_format:
            if isinstance(value, (int, float)):
                return f"{value * 100:.2f}%"
        
        # Handle currency
        if number_format and any(sym in number_format for sym in ['$', '€', '£', '¥']):
            if isinstance(value, (int, float)):
                # Extract currency symbol
                currency_symbol = '$' if '$' in number_format else '€' if '€' in number_format else '£' if '£' in number_format else '¥'
                return f"{currency_symbol}{value:,.2f}"
        
        # Handle numbers with decimals
        if isinstance(value, float):
            # Check if it's a whole number
            if value.is_integer():
                return str(int(value))
            else:
                return f"{value:.2f}"
        
        # Default: convert to string
        return str(value)
    
    def parse_multiple_sheets(self, file_path: str) -> Dict[str, Dict]:
        """
        Parse all sheets in an Excel file.
        
        Args:
            file_path: Path to Excel file
            
        Returns:
            Dictionary mapping sheet names to parsed data
        """
        if not Path(file_path).exists():
            raise FileNotFoundError(f"Excel file not found: {file_path}")
        
        try:
            excel_file = pd.ExcelFile(file_path)
            result = {}
            
            for sheet_name in excel_file.sheet_names:
                try:
                    result[sheet_name] = self.parse_excel_data(file_path, sheet_name)
                except Exception as e:
                    result[sheet_name] = {'error': str(e)}
            
            return result
            
        except Exception as e:
            raise ValueError(f"Error reading Excel file: {str(e)}")
    
    def validate_data_completeness(self, data: Dict) -> Dict:
        """
        Validate that all variables have values in all data rows.
        
        Args:
            data: Parsed data from parse_excel_data
            
        Returns:
            Dictionary with validation results:
            {
                'is_valid': bool,
                'missing_data': List of issues,
                'empty_cells': int
            }
        """
        missing_data = []
        empty_cells = 0
        
        for row_idx, row in enumerate(data['data'], start=2):  # Start at 2 (row 1 is header)
            for var_name, value in row.items():
                if not value or value.strip() == "":
                    missing_data.append({
                        'row': row_idx,
                        'variable': var_name,
                        'message': f"Empty value for '{var_name}' in row {row_idx}"
                    })
                    empty_cells += 1
        
        return {
            'is_valid': len(missing_data) == 0,
            'missing_data': missing_data,
            'empty_cells': empty_cells,
            'total_cells': len(data['variables']) * len(data['data'])
        }
    
    def get_sample_data(self, data: Dict, num_rows: int = 3) -> List[Dict]:
        """
        Get sample data rows for preview.
        
        Args:
            data: Parsed data from parse_excel_data
            num_rows: Number of rows to return
            
        Returns:
            List of sample data rows
        """
        return data['data'][:num_rows]
