"""
Format Converter Service
Converts documents between different formats (PDF, Word, Excel, .msg).
"""
import os
import sys
import shutil
import subprocess
import time
from pathlib import Path
from typing import List, Dict, Optional

try:
    from docx import Document
    from docx2pdf import convert as docx2pdf_convert
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False
    Document = None

try:
    import openpyxl
    from openpyxl import Workbook
except ImportError:
    openpyxl = None
    Workbook = None

try:
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

try:
    import win32com.client
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    pythoncom = None

# Check for LibreOffice installation (portable or system)
def _check_libreoffice():
    """Check if LibreOffice is available (portable or system)."""
    # First, try portable version - just check if file exists
    portable_path = _get_portable_soffice_path()
    if portable_path and os.path.exists(portable_path):
        print(f"[LibreOffice] Found portable version at: {portable_path}")
        return True
    
    # Fall back to system installation
    try:
        result = subprocess.run(['soffice', '--version'], 
                              capture_output=True, timeout=10)
        if result.returncode == 0:
            print(f"[LibreOffice] Found system installation")
            return True
    except:
        pass
    
    return False

def _get_portable_soffice_path():
    """Get path to portable LibreOffice executable."""
    # Get the base directory (handles both frozen and development)
    if getattr(sys, 'frozen', False):
        # Running in PyInstaller bundle
        base_dir = sys._MEIPASS
    else:
        # Running in development
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    portable_path = os.path.join(base_dir, 'portable', 'libreoffice', 'program', 'soffice.exe')
    return portable_path

LIBREOFFICE_AVAILABLE = _check_libreoffice()


class FormatConverter:
    """Converts documents between various formats."""
    
    def __init__(self):
        """Initialize FormatConverter."""
        self.supported_inputs = ['.docx', '.xlsx', '.msg']
        self.supported_outputs = ['pdf', 'word', 'excel', 'excel_workbook', 'msg']
        
        # Log available conversion methods
        methods = []
        if WIN32_AVAILABLE:
            methods.append("MS Office COM (reliable)")
        if REPORTLAB_AVAILABLE:
            methods.append("ReportLab (fast, basic)")
        if LIBREOFFICE_AVAILABLE:
            methods.append("LibreOffice (optional)")
        print(f"[FormatConverter] Available methods: {', '.join(methods) if methods else 'None'}")
    
    def convert(self, input_path: str, output_format: str, output_dir: str, print_settings: Optional[Dict] = None) -> str:
        """
        Convert a document to specified format.
        
        Args:
            input_path: Path to input file
            output_format: Target format ('pdf', 'word', 'excel', 'excel_workbook', 'msg')
            output_dir: Directory for output file
            
        Returns:
            Path to converted file
        """
        input_ext = Path(input_path).suffix.lower()
        base_name = Path(input_path).stem
        
        os.makedirs(output_dir, exist_ok=True)
        
        # Route to appropriate conversion method
        if output_format == 'pdf':
            return self._convert_to_pdf(input_path, output_dir, print_settings)
        elif output_format == 'word':
            return self._convert_to_word(input_path, output_dir)
        elif output_format == 'excel':
            return self._convert_to_excel_single(input_path, output_dir)
        elif output_format == 'excel_workbook':
            return self._convert_to_excel_workbook(input_path, output_dir)
        elif output_format == 'msg':
            return self._convert_to_msg(input_path, output_dir)
        else:
            raise ValueError(f"Unsupported output format: {output_format}")
    
    def _convert_to_pdf(self, input_path: str, output_dir: str, print_settings: Optional[Dict] = None) -> str:
        """Convert document to PDF."""
        input_ext = Path(input_path).suffix.lower()
        base_name = Path(input_path).stem
        output_path = os.path.join(output_dir, f"{base_name}.pdf")
        
        print(f"Converting {input_ext} to PDF: {input_path} -> {output_path}")
        
        try:
            if input_ext == '.docx':
                # Convert Word to PDF - Try LibreOffice first (faster), fall back to COM
                if LIBREOFFICE_AVAILABLE:
                    try:
                        # LibreOffice: Fast (1-2s), portable, good quality
                        print(f"[FormatConverter] Using LibreOffice for Word→PDF")
                        self._libreoffice_to_pdf(input_path, output_path)
                    except Exception as e:
                        print(f"[FormatConverter] LibreOffice failed: {e}")
                        if WIN32_AVAILABLE:
                            print(f"[FormatConverter] Falling back to MS Word COM")
                            self._docx_to_pdf_com(input_path, output_path)
                        else:
                            raise
                elif WIN32_AVAILABLE:
                    # COM: Excellent quality, reliable, requires Office
                    print(f"[FormatConverter] Using MS Word COM for Word→PDF")
                    self._docx_to_pdf_com(input_path, output_path)
                else:
                    raise ImportError("PDF conversion requires MS Office or LibreOffice")
            
            elif input_ext == '.xlsx':
                # Convert Excel to PDF - check if complex print settings needed
                use_reportlab = not print_settings or not self._has_complex_print_settings(print_settings)
                
                if use_reportlab and REPORTLAB_AVAILABLE:
                    # ReportLab: Very fast (0.5-1s), basic formatting
                    print(f"[FormatConverter] Using ReportLab for Excel→PDF")
                    self._xlsx_to_pdf_reportlab(input_path, output_path)
                elif LIBREOFFICE_AVAILABLE:
                    try:
                        # LibreOffice: Fast, portable
                        print(f"[FormatConverter] Using LibreOffice for Excel→PDF")
                        self._libreoffice_to_pdf(input_path, output_path)
                    except Exception as e:
                        print(f"[FormatConverter] LibreOffice failed: {e}")
                        if WIN32_AVAILABLE:
                            print(f"[FormatConverter] Falling back to MS Excel COM")
                            self._xlsx_to_pdf_com(input_path, output_path, print_settings)
                        else:
                            raise
                elif WIN32_AVAILABLE:
                    # COM: Supports complex print settings, reliable
                    print(f"[FormatConverter] Using MS Excel COM for Excel→PDF")
                    self._xlsx_to_pdf_com(input_path, output_path, print_settings)
                    # LibreOffice: Fast (2-3s), good formatting
                    print(f"[FormatConverter] Using LibreOffice for Excel→PDF")
                    self._libreoffice_to_pdf(input_path, output_path)
                else:
                    raise ImportError("Excel PDF conversion requires MS Office, ReportLab, or LibreOffice")
            
            elif input_ext == '.msg':
                # Convert MSG to PDF - requires Outlook
                if WIN32_AVAILABLE:
                    self._msg_to_pdf_com(input_path, output_path)
                else:
                    raise ImportError("MSG to PDF conversion requires pywin32 and Outlook")
            
            else:
                raise ValueError(f"Cannot convert {input_ext} to PDF")
            
            # Verify file was created
            if not os.path.exists(output_path):
                raise RuntimeError(f"PDF file was not created: {output_path}")
            
            print(f"PDF created successfully: {output_path}")
            return output_path
            
        except Exception as e:
            print(f"Error converting to PDF: {str(e)}")
            raise
    
    def _has_complex_print_settings(self, print_settings: Dict) -> bool:
        """Check if print settings are complex enough to require COM."""
        if not print_settings:
            return False
        # If any non-basic settings are present, consider it complex
        complex_keys = ['FitToPagesTall', 'FitToPagesWide', 'PaperSize', 'PrintTitleRows']
        return any(key in print_settings for key in complex_keys)
    
    def _libreoffice_to_pdf(self, input_path: str, output_path: str):
        """
        Convert document to PDF using LibreOffice headless mode.
        60% faster than COM, portable, good quality.
        """
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
        
        abs_input = os.path.abspath(input_path)
        abs_output = os.path.abspath(output_path)
        output_dir = os.path.dirname(abs_output)
        base_name = Path(abs_input).stem
        
        print(f"[LibreOffice] Converting: {Path(abs_input).name}")
        
        # Get LibreOffice executable (portable or system)
        portable_path = _get_portable_soffice_path()
        if portable_path and os.path.exists(portable_path):
            soffice_cmd = portable_path
            print(f"[LibreOffice] Using portable version: {soffice_cmd}")
        else:
            soffice_cmd = 'soffice'
            print(f"[LibreOffice] Using system installation")
        
        try:
            # LibreOffice command: soffice --headless --convert-to pdf --outdir <dir> <file>
            cmd = [
                soffice_cmd,
                '--headless',
                '--invisible',
                '--nocrashreport',
                '--nodefault',
                '--nofirststartwizard',
                '--nolockcheck',
                '--nologo',
                '--norestore',
                '--convert-to', 'pdf',
                '--outdir', output_dir,
                abs_input
            ]
            
            # Hide console window on Windows
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                timeout=60,  # Increased timeout for first run
                text=True,
                startupinfo=startupinfo,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )
            
            if result.returncode != 0:
                error_msg = result.stderr or result.stdout or 'Unknown error'
                raise RuntimeError(f"LibreOffice conversion failed: {error_msg}")
            
            # LibreOffice creates file as <basename>.pdf in output_dir
            expected_file = os.path.join(output_dir, f"{base_name}.pdf")
            
            # Rename if needed
            if expected_file != abs_output and os.path.exists(expected_file):
                if os.path.exists(abs_output):
                    os.remove(abs_output)
                os.rename(expected_file, abs_output)
            
            if not os.path.exists(abs_output):
                raise RuntimeError(f"PDF not created at expected location: {abs_output}")
            
            print(f"[LibreOffice] Conversion successful")
            
        except subprocess.TimeoutExpired:
            raise RuntimeError("LibreOffice conversion timed out after 30 seconds")
        except FileNotFoundError:
            raise RuntimeError("LibreOffice not found. Please install LibreOffice or use MS Office.")
    
    def _docx_to_pdf_com(self, input_path: str, output_path: str):
        """Convert Word to PDF using COM automation."""
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
        
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            
            try:
                abs_input = os.path.abspath(input_path)
                abs_output = os.path.abspath(output_path)
                
                print(f"Opening Word document: {abs_input}")
                doc = word.Documents.Open(abs_input)
                
                print(f"Saving as PDF: {abs_output}")
                doc.SaveAs(abs_output, FileFormat=17)  # 17 = PDF
                doc.Close(False)
                
                print(f"Word document closed")
                
            except Exception as e:
                print(f"Error in Word COM operation: {str(e)}")
                raise
            finally:
                word.Quit()
        finally:
            pythoncom.CoUninitialize()
        
        # Verify PDF was created
        if not os.path.exists(output_path):
            raise RuntimeError(f"PDF file was not created by Word: {output_path}")
    
    def _xlsx_to_pdf_com(self, input_path: str, output_path: str, print_settings: Optional[Dict] = None):
        """Convert Excel to PDF using COM automation."""
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
        
        print(f"Starting Excel to PDF conversion: {input_path} -> {output_path}")
        
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            try:
                abs_input = os.path.abspath(input_path)
                abs_output = os.path.abspath(output_path)
                
                print(f"Opening Excel workbook: {abs_input}")
                wb = excel.Workbooks.Open(abs_input, ReadOnly=False, UpdateLinks=False)
                
                # Save the workbook first to ensure it's not in a temporary state
                wb.Save()
                
                # Apply print settings if provided
                if print_settings:
                    for sheet in wb.Worksheets:
                        # Page setup
                        ps = sheet.PageSetup
                        
                        # Orientation: 1 = Portrait, 2 = Landscape
                        if print_settings.get('orientation') == 'landscape':
                            ps.Orientation = 2
                        else:
                            ps.Orientation = 1
                        
                        # Paper size (e.g., 1 = Letter, 9 = A4)
                        paper_sizes = {
                            'letter': 1, 'a4': 9, 'a3': 8, 'legal': 5,
                            'tabloid': 3, 'a5': 11
                        }
                        paper = print_settings.get('paper_size', 'a4').lower()
                        if paper in paper_sizes:
                            ps.PaperSize = paper_sizes[paper]
                        
                        # Margins (in inches)
                        if 'margins' in print_settings:
                            margins = print_settings['margins']
                            ps.LeftMargin = excel.InchesToPoints(margins.get('left', 0.75))
                            ps.RightMargin = excel.InchesToPoints(margins.get('right', 0.75))
                            ps.TopMargin = excel.InchesToPoints(margins.get('top', 1.0))
                            ps.BottomMargin = excel.InchesToPoints(margins.get('bottom', 1.0))
                        
                        # Scaling
                        if 'scaling' in print_settings:
                            scaling = print_settings['scaling']
                            scaling_type = scaling.get('type', 'percent')
                            
                            if scaling_type == 'no_scaling':
                                # No scaling - 100%
                                ps.Zoom = 100
                                ps.FitToPagesWide = False
                                ps.FitToPagesTall = False
                            elif scaling_type == 'percent':
                                # Scale to percentage
                                ps.Zoom = scaling.get('value', 100)
                                ps.FitToPagesWide = False
                                ps.FitToPagesTall = False
                            elif scaling_type == 'fit_to':
                                # Fit to specific pages wide/tall
                                ps.Zoom = False
                                ps.FitToPagesWide = scaling.get('width', 1)
                                ps.FitToPagesTall = scaling.get('height', 1)
                            elif scaling_type == 'fit_sheet_on_one_page':
                                # Fit entire sheet on one page
                                ps.Zoom = False
                                ps.FitToPagesWide = 1
                                ps.FitToPagesTall = 1
                            elif scaling_type == 'fit_all_columns_on_one_page':
                                # Fit all columns on one page width
                                ps.Zoom = False
                                ps.FitToPagesWide = 1
                                ps.FitToPagesTall = False
                            elif scaling_type == 'fit_all_rows_on_one_page':
                                # Fit all rows on one page height
                                ps.Zoom = False
                                ps.FitToPagesWide = False
                                ps.FitToPagesTall = 1
                        
                        # Print quality
                        if 'print_quality' in print_settings:
                            ps.PrintQuality = print_settings['print_quality']
                        
                        # Center on page
                        if print_settings.get('center_horizontally'):
                            ps.CenterHorizontally = True
                        if print_settings.get('center_vertically'):
                            ps.CenterVertically = True
                
                # Export to PDF
                # IgnorePrintAreas parameter (0 = False, 1 = True)
                ignore_print_areas = 0
                if print_settings and print_settings.get('ignore_print_areas'):
                    ignore_print_areas = 1
                
                # Page range
                from_page = 1
                to_page = 0  # 0 means all pages
                if print_settings and 'page_range' in print_settings:
                    from_page = print_settings['page_range'].get('from', 1)
                    to_page = print_settings['page_range'].get('to', 0)
            
                print(f"Exporting to PDF: {abs_output}")
                
                # Ensure output directory exists
                output_dir_path = os.path.dirname(abs_output)
                if not os.path.exists(output_dir_path):
                    os.makedirs(output_dir_path, exist_ok=True)
                
                # Try different export methods
                try:
                    # Method 1: ExportAsFixedFormat with minimal parameters
                    if from_page == 1 and to_page == 0:
                        # Export all pages
                        wb.ExportAsFixedFormat(
                            Type=0,  # xlTypePDF
                            Filename=abs_output,
                            Quality=0,  # xlQualityStandard
                            IncludeDocProperties=True,
                            IgnorePrintAreas=False,
                            OpenAfterPublish=False
                        )
                    else:
                        # Export specific page range
                        wb.ExportAsFixedFormat(
                            Type=0,
                            Filename=abs_output,
                            Quality=0,
                            IncludeDocProperties=True,
                            IgnorePrintAreas=ignore_print_areas,
                            From=from_page,
                            To=to_page,
                            OpenAfterPublish=False
                        )
                except Exception as export_error:
                    print(f"ExportAsFixedFormat failed, trying alternative method: {str(export_error)}")
                    
                    # Method 2: Use PrintOut to PDF printer as fallback
                    # First try to use Microsoft Print to PDF
                    try:
                        wb.PrintOut(
                            ActivePrinter="Microsoft Print to PDF",
                            PrintToFile=True,
                            PrToFileName=abs_output
                        )
                    except:
                        # If that fails, save as PDF using SaveAs
                        file_format = 57  # xlTypePDF
                        wb.SaveAs(abs_output, FileFormat=file_format)
                
                print(f"Closing Excel workbook")
                wb.Close(SaveChanges=False)
                
                # Small delay to ensure file is written
                import time
                time.sleep(0.5)
                
            except Exception as e:
                print(f"Error in Excel COM operation: {str(e)}")
                # Try to close workbook if it's open
                try:
                    if 'wb' in locals():
                        wb.Close(SaveChanges=False)
                except:
                    pass
                raise
            finally:
                try:
                    excel.Quit()
                    # Give Excel time to fully quit
                    import time
                    time.sleep(0.3)
                except:
                    pass
        finally:
            pythoncom.CoUninitialize()
        
        # Verify PDF was created
        if not os.path.exists(output_path):
            raise RuntimeError(f"PDF file was not created by Excel: {output_path}")
        
        file_size = os.path.getsize(output_path)
        print(f"Excel PDF created successfully: {output_path} ({file_size} bytes)")
    
    def _xlsx_to_pdf_reportlab(self, input_path: str, output_path: str):
        """Convert Excel to PDF using ReportLab (fallback)."""
        if not REPORTLAB_AVAILABLE:
            raise ImportError("reportlab is required for Excel to PDF conversion")
        
        if openpyxl is None:
            raise ImportError("openpyxl is required for Excel reading")
        
        wb = openpyxl.load_workbook(input_path)
        sheet = wb.active
        
        # Create PDF
        doc = SimpleDocTemplate(output_path, pagesize=A4)
        elements = []
        styles = getSampleStyleSheet()
        
        # Add title
        elements.append(Paragraph(f"<b>{Path(input_path).stem}</b>", styles['Title']))
        elements.append(Spacer(1, 12))
        
        # Convert sheet to table data
        data = []
        for row in sheet.iter_rows(values_only=True):
            data.append([str(cell) if cell is not None else "" for cell in row])
        
        if data:
            table = Table(data)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            elements.append(table)
        
        doc.build(elements)
    
    def _msg_to_pdf_com(self, input_path: str, output_path: str):
        """Convert MSG to PDF using COM automation."""
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            
            try:
                msg = outlook.CreateItemFromTemplate(os.path.abspath(input_path))
                
                # Save as Word document first, then convert to PDF
                temp_doc = output_path.replace('.pdf', '_temp.docx')
                msg.SaveAs(temp_doc, 4)  # 4 = Word format
                
                # Convert Word to PDF
                self._docx_to_pdf_com(temp_doc, output_path)
                
                # Clean up temp file
                if os.path.exists(temp_doc):
                    os.remove(temp_doc)
                
                # Verify PDF was created
                if not os.path.exists(output_path):
                    raise RuntimeError(f"PDF file was not created from MSG: {output_path}")
            
            except Exception as e:
                raise ValueError(f"Error converting MSG to PDF: {str(e)}")
        finally:
            pythoncom.CoUninitialize()
    
    def _convert_to_word(self, input_path: str, output_dir: str) -> str:
        """Convert document to Word format."""
        input_ext = Path(input_path).suffix.lower()
        base_name = Path(input_path).stem
        output_path = os.path.join(output_dir, f"{base_name}.docx")
        
        if input_ext == '.docx':
            # Already Word format, just copy
            shutil.copy2(input_path, output_path)
        elif input_ext == '.msg':
            # Convert MSG to Word
            if WIN32_AVAILABLE:
                pythoncom.CoInitialize()
                try:
                    outlook = win32com.client.Dispatch("Outlook.Application")
                    msg = outlook.CreateItemFromTemplate(os.path.abspath(input_path))
                    msg.SaveAs(os.path.abspath(output_path), 4)  # 4 = Word format
                finally:
                    pythoncom.CoUninitialize()
            else:
                raise ImportError("MSG to Word conversion requires pywin32")
        else:
            raise ValueError(f"Cannot convert {input_ext} to Word")
        
        return output_path
    
    def _convert_to_excel_single(self, input_path: str, output_dir: str) -> str:
        """Convert document to single Excel sheet."""
        input_ext = Path(input_path).suffix.lower()
        base_name = Path(input_path).stem
        output_path = os.path.join(output_dir, f"{base_name}.xlsx")
        
        if input_ext == '.xlsx':
            # Already Excel, just copy
            shutil.copy2(input_path, output_path)
        else:
            raise ValueError(f"Cannot convert {input_ext} to Excel single sheet")
        
        return output_path
    
    def _convert_to_excel_workbook(self, input_path: str, output_dir: str) -> str:
        """
        Excel workbook conversion is handled by job_manager._merge_excel_workbook().
        This method should not be called directly during individual file conversion.
        """
        raise RuntimeError(
            "Excel workbook format should not be converted individually. "
            "Workbook merging is handled by job_manager after all individual files are processed."
        )
    
    def _convert_to_msg(self, input_path: str, output_dir: str) -> str:
        """Convert document to MSG format."""
        input_ext = Path(input_path).suffix.lower()
        base_name = Path(input_path).stem
        output_path = os.path.join(output_dir, f"{base_name}.msg")
        
        if input_ext == '.msg':
            # Already MSG format, just copy
            shutil.copy2(input_path, output_path)
        else:
            raise ValueError(f"Cannot convert {input_ext} to MSG")
        
        return output_path
    
    def batch_convert(self, input_paths: List[str], output_formats: List[str], output_dir: str) -> List[str]:
        """
        Convert multiple files to multiple formats.
        
        Args:
            input_paths: List of input file paths
            output_formats: List of output formats
            output_dir: Directory for output files
            
        Returns:
            List of output file paths
        """
        output_files = []
        
        for input_path in input_paths:
            for output_format in output_formats:
                try:
                    output_file = self.convert(input_path, output_format, output_dir)
                    output_files.append(output_file)
                except Exception as e:
                    print(f"Error converting {input_path} to {output_format}: {str(e)}")
        
        return output_files
