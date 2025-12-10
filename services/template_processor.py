"""
Template Processor Service
Processes templates in Word, Excel, and .msg formats with placeholder substitution.
"""
import os
import re
from pathlib import Path
from typing import Dict, List
import shutil

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
except ImportError:
    Document = None

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
except ImportError:
    openpyxl = None

try:
    import win32com.client
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    pythoncom = None


class TemplateProcessor:
    """Processes templates with placeholder substitution."""
    
    VARIABLE_PATTERN = re.compile(r'##([^#]+)##')
    
    def __init__(self):
        """Initialize TemplateProcessor."""
        self.supported_formats = ['.docx', '.xlsx', '.msg']
    
    def is_supported_format(self, file_path: str) -> bool:
        """Check if file format is supported."""
        ext = Path(file_path).suffix.lower()
        return ext in self.supported_formats
    
    def extract_template_variables(self, template_path: str) -> List[str]:
        """
        Extract all ##variable## placeholders from a template.
        
        Args:
            template_path: Path to template file
            
        Returns:
            List of unique variable names found in template
        """
        ext = Path(template_path).suffix.lower()
        
        if ext == '.docx':
            return self._extract_variables_from_docx(template_path)
        elif ext == '.xlsx':
            return self._extract_variables_from_xlsx(template_path)
        elif ext == '.msg':
            return self._extract_variables_from_msg(template_path)
        else:
            raise ValueError(f"Unsupported template format: {ext}")
    
    def _extract_variables_from_docx(self, file_path: str) -> List[str]:
        """Extract variables from Word document."""
        if Document is None:
            raise ImportError("python-docx is required for Word templates")
        
        doc = Document(file_path)
        variables = set()
        
        # Check paragraphs
        for para in doc.paragraphs:
            variables.update(self.VARIABLE_PATTERN.findall(para.text))
        
        # Check tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    variables.update(self.VARIABLE_PATTERN.findall(cell.text))
        
        # Check headers and footers
        for section in doc.sections:
            for para in section.header.paragraphs:
                variables.update(self.VARIABLE_PATTERN.findall(para.text))
            for para in section.footer.paragraphs:
                variables.update(self.VARIABLE_PATTERN.findall(para.text))
        
        return sorted(list(variables))
    
    def _extract_variables_from_xlsx(self, file_path: str) -> List[str]:
        """Extract variables from Excel workbook."""
        if openpyxl is None:
            raise ImportError("openpyxl is required for Excel templates")
        
        wb = openpyxl.load_workbook(file_path)
        variables = set()
        
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        variables.update(self.VARIABLE_PATTERN.findall(cell.value))
        
        return sorted(list(variables))
    
    def _extract_variables_from_msg(self, file_path: str) -> List[str]:
        """Extract variables from .msg file."""
        if not WIN32_AVAILABLE:
            raise ImportError("pywin32 is required for .msg templates")
        
        variables = set()
        
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            msg = outlook.CreateItemFromTemplate(file_path)
            
            # Check subject
            if msg.Subject:
                variables.update(self.VARIABLE_PATTERN.findall(msg.Subject))
            
            # Check body
            if msg.Body:
                variables.update(self.VARIABLE_PATTERN.findall(msg.Body))
            
            # Check HTML body
            if msg.HTMLBody:
                variables.update(self.VARIABLE_PATTERN.findall(msg.HTMLBody))
            
        except Exception as e:
            raise ValueError(f"Error reading .msg template: {str(e)}")
        finally:
            pythoncom.CoUninitialize()
        
        return sorted(list(variables))
    
    def process_template(self, template_path: str, data: Dict, output_path: str, sheet_name: str = None) -> str:
        """
        Process a template with data substitution.
        
        Args:
            template_path: Path to template file
            data: Dictionary mapping variable names to values
            output_path: Path for output file
            sheet_name: Optional sheet name for Excel templates
            
        Returns:
            Path to the generated file
        """
        ext = Path(template_path).suffix.lower()
        
        if ext == '.docx':
            return self._process_docx_template(template_path, data, output_path)
        elif ext == '.xlsx':
            return self._process_xlsx_template(template_path, data, output_path, sheet_name)
        elif ext == '.msg':
            return self._process_msg_template(template_path, data, output_path)
        else:
            raise ValueError(f"Unsupported template format: {ext}")
    
    def _replace_text_in_paragraph(self, paragraph, data: Dict):
        """Replace variables in a paragraph while preserving formatting."""
        for var_name, value in data.items():
            placeholder = f"##{var_name}##"
            if placeholder in paragraph.text:
                # Replace inline while preserving runs
                for run in paragraph.runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, str(value))
    
    def _process_docx_template(self, template_path: str, data: Dict, output_path: str) -> str:
        """Process Word template."""
        if Document is None:
            raise ImportError("python-docx is required for Word templates")
        
        doc = Document(template_path)
        
        # Replace in paragraphs
        for para in doc.paragraphs:
            self._replace_text_in_paragraph(para, data)
        
        # Replace in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        self._replace_text_in_paragraph(para, data)
        
        # Replace in headers and footers
        for section in doc.sections:
            for para in section.header.paragraphs:
                self._replace_text_in_paragraph(para, data)
            for para in section.footer.paragraphs:
                self._replace_text_in_paragraph(para, data)
        
        # Save the document
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        doc.save(output_path)
        return output_path
    
    def _process_xlsx_template(self, template_path: str, data: Dict, output_path: str, sheet_name: str = None) -> str:
        """Process Excel template."""
        if openpyxl is None:
            raise ImportError("openpyxl is required for Excel templates")
        
        wb = openpyxl.load_workbook(template_path)
        
        # Determine which sheets to process
        sheets_to_process = []
        if sheet_name:
            # Process only the specified sheet
            if sheet_name in wb.sheetnames:
                sheets_to_process = [wb[sheet_name]]
            else:
                raise ValueError(f"Sheet '{sheet_name}' not found in template")
        else:
            # Process all sheets
            sheets_to_process = wb.worksheets
        
        for sheet in sheets_to_process:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        new_value = cell.value
                        for var_name, value in data.items():
                            placeholder = f"##{var_name}##"
                            new_value = new_value.replace(placeholder, str(value))
                        if new_value != cell.value:
                            cell.value = new_value
        
        # Save the workbook
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        wb.save(output_path)
        return output_path
    
    def _process_msg_template(self, template_path: str, data: Dict, output_path: str) -> str:
        """Process .msg template."""
        if not WIN32_AVAILABLE:
            raise ImportError("pywin32 is required for .msg templates")
        
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            msg = outlook.CreateItemFromTemplate(template_path)
            
            # Replace in subject
            if msg.Subject:
                for var_name, value in data.items():
                    placeholder = f"##{var_name}##"
                    msg.Subject = msg.Subject.replace(placeholder, str(value))
            
            # Replace in body
            if msg.Body:
                for var_name, value in data.items():
                    placeholder = f"##{var_name}##"
                    msg.Body = msg.Body.replace(placeholder, str(value))
            
            # Replace in HTML body
            if msg.HTMLBody:
                for var_name, value in data.items():
                    placeholder = f"#{var_name}##"
                    msg.HTMLBody = msg.HTMLBody.replace(placeholder, str(value))
            
            # Save the message
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            msg.SaveAs(output_path)
            
        except Exception as e:
            raise ValueError(f"Error processing .msg template: {str(e)}")
        finally:
            pythoncom.CoUninitialize()
        
        return output_path
    
    def validate_template_data(self, template_path: str, data: Dict) -> Dict:
        """
        Validate that all template variables have corresponding data.
        
        Args:
            template_path: Path to template file
            data: Data dictionary
            
        Returns:
            Validation result with missing/extra variables
        """
        template_vars = set(self.extract_template_variables(template_path))
        data_vars = set(data.keys())
        
        missing_vars = template_vars - data_vars
        extra_vars = data_vars - template_vars
        
        return {
            'is_valid': len(missing_vars) == 0,
            'missing_variables': sorted(list(missing_vars)),
            'extra_variables': sorted(list(extra_vars)),
            'template_variables': sorted(list(template_vars)),
            'data_variables': sorted(list(data_vars))
        }
