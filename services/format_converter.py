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
from utils.file_handlers import open_workbook_safe

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
        # Ensure all paths are absolute
        input_path = str(Path(input_path).absolute())
        output_dir = str(Path(output_dir).absolute())
        
        input_ext = Path(input_path).suffix.lower()
        base_name = Path(input_path).stem
        output_path = str(Path(output_dir) / f"{base_name}.pdf")
        
        print(f"[FormatConverter] Converting {input_ext} to PDF")
        print(f"[FormatConverter] Input: {input_path}")
        print(f"[FormatConverter] Output: {output_path}")
        print(f"[FormatConverter] Input exists: {os.path.exists(input_path)}")
        print(f"[FormatConverter] Output dir exists: {os.path.exists(output_dir)}")
        
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
        
        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)
        
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
                # Convert Excel to PDF
                # CRITICAL: ONLY use MS Excel COM for Excel→PDF conversion
                # LibreOffice produces POOR QUALITY Excel PDFs with formatting issues
                # LibreOffice is OK for generating .xlsx files, but NOT for PDF conversion
                
                if not WIN32_AVAILABLE:
                    raise RuntimeError(
                        "Excel to PDF conversion REQUIRES MS Excel (pywin32). "
                        "LibreOffice produces poor quality Excel PDFs and is NOT used. "
                        "Please ensure MS Excel is installed."
                    )
                
                try:
                    # MS Excel COM: REQUIRED for proper Excel PDF formatting
                    print(f"[FormatConverter] Using MS Excel COM for Excel→PDF (REQUIRED for quality)")
                    self._xlsx_to_pdf_com(input_path, output_path, print_settings)
                except Exception as com_error:
                    # If COM fails, this is a CRITICAL error - do NOT fall back to LibreOffice
                    error_msg = (
                        f"Excel COM PDF conversion FAILED: {str(com_error)}\n"
                        f"Excel PDFs require MS Excel COM for proper formatting.\n"
                        f"LibreOffice/ReportLab produce poor quality Excel PDFs and are NOT used as fallback."
                    )
                    print(f"[FormatConverter] ERROR: {error_msg}")
                    raise RuntimeError(error_msg)
            
            elif input_ext == '.msg':
                # Convert MSG to PDF - requires Outlook
                if WIN32_AVAILABLE:
                    self._msg_to_pdf_com(input_path, output_path)
                else:
                    raise ImportError("MSG to PDF conversion requires pywin32 and Outlook")
            
            else:
                raise ValueError(f"Cannot convert {input_ext} to PDF")
            
            # Verify file was created with detailed logging
            print(f"[FormatConverter] Verifying PDF creation...")
            print(f"[FormatConverter] Checking path: {output_path}")
            print(f"[FormatConverter] Path is absolute: {Path(output_path).is_absolute()}")
            print(f"[FormatConverter] File exists: {os.path.exists(output_path)}")
            
            if not os.path.exists(output_path):
                # Try to list what files ARE in the output directory
                if os.path.exists(output_dir):
                    files_in_dir = list(Path(output_dir).iterdir())
                    print(f"[FormatConverter] Files in output dir: {files_in_dir}")
                raise RuntimeError(f"PDF file was not created: {output_path}")
            
            file_size = os.path.getsize(output_path)
            print(f"[FormatConverter] PDF created successfully: {output_path} ({file_size} bytes)")
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
        # Ensure absolute paths
        input_path = str(Path(input_path).absolute())
        output_path = str(Path(output_path).absolute())
        
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
        
        abs_input = input_path
        abs_output = output_path
        output_dir = os.path.dirname(abs_output)
        base_name = Path(abs_input).stem
        
        print(f"[LibreOffice] Converting: {Path(abs_input).name}")
        print(f"[LibreOffice] Input: {abs_input}")
        print(f"[LibreOffice] Output: {abs_output}")
        print(f"[LibreOffice] Output dir: {output_dir}")
        
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
            # For Excel files, we can add filter options to preserve formatting better
            input_ext = Path(abs_input).suffix.lower()
            
            # Build filter options for Excel files
            filter_opts = []
            if input_ext in ['.xlsx', '.xls']:
                # Excel-specific PDF export options for LibreOffice Calc
                # UseISOPaperFormatting=false preserves Excel page setup
                # EmbedStandardFonts=true ensures fonts display correctly
                # ExportFormFields=false prevents form field issues
                # FormsType=0 means no form export
                filter_opts = [
                    'PageRange=All',
                    'MaxImageResolution=300',
                    'Quality=90',
                    'ReduceImageResolution=false',
                    'UseISOPaperFormatting=false',
                    'EmbedStandardFonts=true',
                    'ExportFormFields=false',
                    'FormsType=0'
                ]
            
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
                '--convert-to'
            ]
            
            # Add filter options if we have them
            if filter_opts:
                filter_str = ':'.join(filter_opts)
                cmd.append(f'pdf:calc_pdf_Export:{filter_str}')
            else:
                cmd.append('pdf')
            
            cmd.extend(['--outdir', output_dir, abs_input])
            
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
                print(f"[LibreOffice] Stderr: {result.stderr}")
                print(f"[LibreOffice] Stdout: {result.stdout}")
                raise RuntimeError(f"LibreOffice conversion failed: {error_msg}")
            
            print(f"[LibreOffice] Conversion command completed successfully")
            
            # LibreOffice creates file as <basename>.pdf in output_dir
            expected_file = os.path.join(output_dir, f"{base_name}.pdf")
            
            print(f"[LibreOffice] Expected file: {expected_file}")
            print(f"[LibreOffice] Target output: {abs_output}")
            print(f"[LibreOffice] Expected file exists: {os.path.exists(expected_file)}")
            
            # List all files in output directory for debugging
            if os.path.exists(output_dir):
                files_in_dir = list(Path(output_dir).iterdir())
                print(f"[LibreOffice] Files in output dir: {[f.name for f in files_in_dir]}")
            
            # Rename if needed
            if expected_file != abs_output and os.path.exists(expected_file):
                print(f"[LibreOffice] Renaming {expected_file} to {abs_output}")
                if os.path.exists(abs_output):
                    os.remove(abs_output)
                os.rename(expected_file, abs_output)
            
            if not os.path.exists(abs_output):
                raise RuntimeError(f"PDF not created at expected location: {abs_output}")
            
            file_size = os.path.getsize(abs_output)
            print(f"[LibreOffice] Conversion successful - {abs_output} ({file_size} bytes)")
            
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
        
        print(f"[Excel COM] Starting Excel to PDF conversion")
        print(f"[Excel COM] Input: {input_path}")
        print(f"[Excel COM] Output: {output_path}")
        
        # Initialize COM for this thread
        try:
            pythoncom.CoInitialize()
            print(f"[Excel COM] COM initialized")
        except Exception as e:
            raise RuntimeError(f"Failed to initialize COM: {str(e)}")
        
        try:
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                print(f"[Excel COM] Excel.Application dispatched")
            except Exception as e:
                raise RuntimeError(f"Failed to create Excel.Application: {str(e)}")
            
            excel.Visible = False
            excel.DisplayAlerts = False
            print(f"[Excel COM] Excel configured (Visible=False, DisplayAlerts=False)")
            
            try:
                abs_input = os.path.abspath(input_path)
                abs_output = os.path.abspath(output_path)
                
                print(f"[Excel COM] Opening workbook: {abs_input}")
                wb = excel.Workbooks.Open(abs_input, ReadOnly=True, UpdateLinks=False)
                print(f"[Excel COM] Workbook opened successfully")
                
                # Apply print settings if provided
                if print_settings:
                    print(f"[Excel COM] Applying print settings...")
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
                print(f"[Excel COM] Exporting to PDF...")
                try:
                    # Method 1: ExportAsFixedFormat with minimal parameters
                    if from_page == 1 and to_page == 0:
                        # Export all pages
                        print(f"[Excel COM] Using ExportAsFixedFormat (all pages)")
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
                        print(f"[Excel COM] Using ExportAsFixedFormat (pages {from_page}-{to_page})")
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
                    print(f"[Excel COM] Export successful")
                except Exception as export_error:
                    print(f"[Excel COM] ExportAsFixedFormat failed: {str(export_error)}")
                    print(f"[Excel COM] Trying alternative method...")
                    
                    # Method 2: Use PrintOut to PDF printer as fallback
                    # First try to use Microsoft Print to PDF
                    try:
                        print(f"[Excel COM] Trying PrintOut method")
                        wb.PrintOut(
                            ActivePrinter="Microsoft Print to PDF",
                            PrintToFile=True,
                            PrToFileName=abs_output
                        )
                        print(f"[Excel COM] PrintOut successful")
                    except Exception as printout_error:
                        print(f"[Excel COM] PrintOut failed: {str(printout_error)}")
                        print(f"[Excel COM] Trying SaveAs method")
                        # If that fails, save as PDF using SaveAs
                        file_format = 57  # xlTypePDF
                        wb.SaveAs(abs_output, FileFormat=file_format)
                        print(f"[Excel COM] SaveAs successful")
                
                print(f"[Excel COM] Closing Excel workbook")
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
                    print(f"[Excel COM] Quitting Excel application")
                    excel.Quit()
                    # Give Excel time to fully quit
                    import time
                    time.sleep(0.3)
                    print(f"[Excel COM] Excel quit successfully")
                except Exception as e:
                    print(f"[Excel COM] Error quitting Excel: {str(e)}")
        finally:
            pythoncom.CoUninitialize()
            print(f"[Excel COM] COM uninitialized")
        
        # Verify PDF was created
        print(f"[Excel COM] Verifying output file: {output_path}")
        if not os.path.exists(output_path):
            raise RuntimeError(f"PDF file was not created by Excel: {output_path}")
        
        file_size = os.path.getsize(output_path)
        print(f"[Excel COM] PDF created successfully: {output_path} ({file_size} bytes)")
    
    def _xlsx_to_pdf_reportlab(self, input_path: str, output_path: str):
        """Convert Excel to PDF using ReportLab (fallback)."""
        if not REPORTLAB_AVAILABLE:
            raise ImportError("reportlab is required for Excel to PDF conversion")
        
        if openpyxl is None:
            raise ImportError("openpyxl is required for Excel reading")
        
        with open_workbook_safe(input_path, data_only=True, read_only=True) as wb:
            sheet = wb.active
            
            # Convert sheet to table data
            data = []
            for row in sheet.iter_rows(values_only=True):
                data.append([str(cell) if cell is not None else "" for cell in row])
            
            # Create PDF
            doc = SimpleDocTemplate(output_path, pagesize=A4)
            elements = []
            styles = getSampleStyleSheet()
            
            # Add title
            elements.append(Paragraph(f"<b>{Path(input_path).stem}</b>", styles['Title']))
            elements.append(Spacer(1, 12))
        
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
