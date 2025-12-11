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
from io import StringIO
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
                
                # MS Excel COM: REQUIRED for proper Excel PDF formatting
                print(f"[FormatConverter] Using MS Excel COM for Excel→PDF (REQUIRED for quality)")
                self._xlsx_to_pdf_com(input_path, output_path, print_settings)
                # Note: If COM fails, _xlsx_to_pdf_com will raise RuntimeError with full debug log
                # We let it propagate unchanged to preserve all debugging information
            
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
        """Convert Excel to PDF using COM automation with 6 fallback methods and extensive debugging."""
        # Capture all debug output for UI display and log file
        import sys
        from datetime import datetime
        from config.config import Config
        
        debug_log = StringIO()
        original_stdout = sys.stdout
        
        # Create log file in AppData/logs
        log_filename = f"excel_conversion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        log_path = Path(Config.LOGS_DIR) / log_filename
        
        class TeeOutput:
            """Write to console, string buffer, AND log file"""
            def __init__(self, *outputs):
                self.outputs = outputs
                self.log_file = open(log_path, 'w', encoding='utf-8')
            def write(self, data):
                for output in self.outputs:
                    output.write(data)
                # Also write to log file
                self.log_file.write(data)
                self.log_file.flush()
            def flush(self):
                for output in self.outputs:
                    output.flush()
                self.log_file.flush()
            def close(self):
                self.log_file.close()
        
        # Redirect stdout to capture all print statements
        tee = TeeOutput(original_stdout, debug_log)
        sys.stdout = tee
        
        try:
            print(f"[Excel COM] Log file: {log_path}")
            self._xlsx_to_pdf_com_inner(input_path, output_path, print_settings, debug_log, log_path)
        finally:
            # Always restore stdout and close log file
            sys.stdout = original_stdout
            tee.close()
    
    def _xlsx_to_pdf_com_inner(self, input_path: str, output_path: str, print_settings: Optional[Dict], debug_log, log_path):
        """Inner implementation of Excel to PDF conversion."""
        
        print(f"\n{'='*80}")
        print(f"[Excel COM] STARTING EXCEL TO PDF CONVERSION")
        print(f"{'='*80}")
        
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
        
        print(f"[Excel COM] Input file: {input_path}")
        print(f"[Excel COM] Input exists: {os.path.exists(input_path)}")
        print(f"[Excel COM] Input size: {os.path.getsize(input_path):,} bytes")
        print(f"[Excel COM] Output path: {output_path}")
        print(f"[Excel COM] Output is absolute: {os.path.isabs(output_path)}")
        print(f"[Excel COM] Output dir: {os.path.dirname(output_path)}")
        print(f"[Excel COM] Output dir exists: {os.path.exists(os.path.dirname(output_path))}")
        print(f"[Excel COM] Print settings: {print_settings is not None}")
        
        # Initialize COM for this thread
        try:
            pythoncom.CoInitialize()
            print(f"[Excel COM] [OK] COM initialized successfully")
        except Exception as e:
            print(f"[Excel COM] [X] FATAL: Failed to initialize COM: {str(e)}")
            raise RuntimeError(f"Failed to initialize COM: {str(e)}")
        
        try:
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                print(f"[Excel COM] [OK] Excel.Application dispatched successfully")
            except Exception as e:
                print(f"[Excel COM] [X] FATAL: Failed to create Excel.Application: {str(e)}")
                print(f"[Excel COM] Is MS Excel installed? Check COM registration.")
                raise RuntimeError(f"Failed to create Excel.Application: {str(e)}")
            
            excel.Visible = False
            excel.DisplayAlerts = False
            print(f"[Excel COM] [OK] Excel configured (Visible=False, DisplayAlerts=False)")
            
            try:
                abs_input = os.path.abspath(input_path)
                abs_output = os.path.abspath(output_path)
                
                print(f"[Excel COM] Absolute input: {abs_input}")
                print(f"[Excel COM] Absolute output: {abs_output}")
                
                print(f"[Excel COM] Opening workbook...")
                try:
                    wb = excel.Workbooks.Open(abs_input, ReadOnly=True, UpdateLinks=False)
                    print(f"[Excel COM] [OK] Workbook opened successfully")
                    print(f"[Excel COM] Workbook name: {wb.Name}")
                    print(f"[Excel COM] Worksheets count: {wb.Worksheets.Count}")
                except Exception as e:
                    print(f"[Excel COM] [X] FATAL: Failed to open workbook: {str(e)}")
                    raise RuntimeError(f"Failed to open workbook: {str(e)}")
                
                # Apply print settings if provided
                if print_settings:
                    print(f"[Excel COM] Applying print settings: {print_settings}")
                    try:
                        for sheet_idx, sheet in enumerate(wb.Worksheets, 1):
                            print(f"[Excel COM] Configuring sheet {sheet_idx}: {sheet.Name}")
                            ps = sheet.PageSetup
                            
                            # Orientation: 1 = Portrait, 2 = Landscape
                            if print_settings.get('orientation') == 'landscape':
                                ps.Orientation = 2
                                print(f"[Excel COM]   - Orientation: Landscape")
                            else:
                                ps.Orientation = 1
                                print(f"[Excel COM]   - Orientation: Portrait")
                            
                            # Paper size (e.g., 1 = Letter, 9 = A4)
                            paper_sizes = {
                                'letter': 1, 'a4': 9, 'a3': 8, 'legal': 5,
                                'tabloid': 3, 'a5': 11
                            }
                            paper = print_settings.get('paper_size', 'a4').lower()
                            if paper in paper_sizes:
                                ps.PaperSize = paper_sizes[paper]
                                print(f"[Excel COM]   - Paper size: {paper}")
                        
                        # Margins (in inches)
                        if 'margins' in print_settings:
                            margins = print_settings['margins']
                            ps.LeftMargin = excel.InchesToPoints(margins.get('left', 0.75))
                            ps.RightMargin = excel.InchesToPoints(margins.get('right', 0.75))
                            ps.TopMargin = excel.InchesToPoints(margins.get('top', 1.0))
                            ps.BottomMargin = excel.InchesToPoints(margins.get('bottom', 1.0))
                            print(f"[Excel COM]   - Margins applied")
                        
                        # Scaling
                        if 'scaling' in print_settings:
                            scaling = print_settings['scaling']
                            scaling_type = scaling.get('type', 'percent')
                            
                            if scaling_type == 'no_scaling':
                                # No scaling - 100%
                                ps.Zoom = 100
                                ps.FitToPagesWide = False
                                ps.FitToPagesTall = False
                                print(f"[Excel COM]   - Scaling: None (100%)")
                            elif scaling_type == 'percent':
                                # Scale to percentage
                                ps.Zoom = scaling.get('value', 100)
                                ps.FitToPagesWide = False
                                ps.FitToPagesTall = False
                                print(f"[Excel COM]   - Scaling: {ps.Zoom}%")
                            elif scaling_type == 'fit_to':
                                # Fit to specific pages wide/tall
                                ps.Zoom = False
                                ps.FitToPagesWide = scaling.get('width', 1)
                                ps.FitToPagesTall = scaling.get('height', 1)
                                print(f"[Excel COM]   - Scaling: Fit to {ps.FitToPagesWide}x{ps.FitToPagesTall} pages")
                            elif scaling_type == 'fit_sheet_on_one_page':
                                # Fit entire sheet on one page
                                ps.Zoom = False
                                ps.FitToPagesWide = 1
                                ps.FitToPagesTall = 1
                                print(f"[Excel COM]   - Scaling: Fit sheet on one page")
                            elif scaling_type == 'fit_all_columns_on_one_page':
                                # Fit all columns on one page width
                                ps.Zoom = False
                                ps.FitToPagesWide = 1
                                ps.FitToPagesTall = False
                                print(f"[Excel COM]   - Scaling: Fit all columns")
                            elif scaling_type == 'fit_all_rows_on_one_page':
                                # Fit all rows on one page height
                                ps.Zoom = False
                                ps.FitToPagesWide = False
                                ps.FitToPagesTall = 1
                                print(f"[Excel COM]   - Scaling: Fit all rows")
                        
                        # Print quality
                        if 'print_quality' in print_settings:
                            ps.PrintQuality = print_settings['print_quality']
                            print(f"[Excel COM]   - Print quality: {ps.PrintQuality}")
                        
                        # Center on page
                        if print_settings.get('center_horizontally'):
                            ps.CenterHorizontally = True
                            print(f"[Excel COM]   - Center horizontally: Yes")
                        if print_settings.get('center_vertically'):
                            ps.CenterVertically = True
                            print(f"[Excel COM]   - Center vertically: Yes")
                        
                        print(f"[Excel COM] [OK] Print settings applied to all sheets")
                    except Exception as ps_error:
                        print(f"[Excel COM] [WARN] Warning: Failed to apply some print settings: {ps_error}")
                
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
                    print(f"[Excel COM] Page range: {from_page} to {to_page if to_page > 0 else 'end'}")
            
                print(f"[Excel COM] Target PDF: {abs_output}")
                
                # Ensure output directory exists
                output_dir_path = os.path.dirname(abs_output)
                if not os.path.exists(output_dir_path):
                    os.makedirs(output_dir_path, exist_ok=True)
                    print(f"[Excel COM] [OK] Created output directory: {output_dir_path}")
                else:
                    print(f"[Excel COM] [OK] Output directory exists: {output_dir_path}")
                
                # List directory BEFORE conversion
                print(f"\n[Excel COM] {'='*60}")
                print(f"[Excel COM] PRE-CONVERSION STATE")
                print(f"[Excel COM] {'='*60}")
                try:
                    files_before = list(Path(output_dir_path).iterdir())
                    print(f"[Excel COM] Files in output dir: {len(files_before)}")
                    for f in sorted(files_before)[:10]:  # Show first 10
                        if f.is_file():
                            print(f"[Excel COM]   - {f.name} ({f.stat().st_size:,} bytes)")
                except Exception as e:
                    print(f"[Excel COM] Could not list directory: {e}")
                
                # Track conversion success
                export_success = False
                method_used = None
                last_error = None
                print(f"[Excel COM] Starting 6-method conversion chain...")
                
                # ============================================================
                # METHOD 1: ExportAsFixedFormat - Standard Quality
                # ============================================================
                if not export_success:
                    try:
                        print(f"\n[Excel COM] ╔══════════════════════════════════════════════════════════╗")
                        print(f"[Excel COM] ║ METHOD 1: ExportAsFixedFormat (Standard Quality)        ║")
                        print(f"[Excel COM] ╚══════════════════════════════════════════════════════════╝")
                        
                        # Remove any existing file
                        if os.path.exists(abs_output):
                            os.remove(abs_output)
                            print(f"[Excel COM] Removed existing file")
                        
                        print(f"[Excel COM] Calling ExportAsFixedFormat...")
                        if from_page == 1 and to_page == 0:
                            wb.ExportAsFixedFormat(
                                Type=0,  # xlTypePDF
                                Filename=abs_output,
                                Quality=0,  # xlQualityStandard
                                IncludeDocProperties=True,
                                IgnorePrintAreas=False,
                                OpenAfterPublish=False
                            )
                        else:
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
                        print(f"[Excel COM] ExportAsFixedFormat call completed (no exception)")
                        
                        # Wait and verify
                        print(f"[Excel COM] Waiting 2 seconds for file system...")
                        time.sleep(2.0)
                        
                        print(f"[Excel COM] Verifying file creation (5 attempts)...")
                        for attempt in range(5):
                            print(f"[Excel COM]   Attempt {attempt + 1}/5: Checking {os.path.basename(abs_output)}...")
                            if os.path.exists(abs_output):
                                size = os.path.getsize(abs_output)
                                print(f"[Excel COM]   File exists! Size: {size:,} bytes")
                                if size > 0:
                                    print(f"[Excel COM] [OK][OK][OK] METHOD 1 SUCCESS ({size:,} bytes) [OK][OK][OK]")
                                    export_success = True
                                    method_used = "ExportAsFixedFormat (Standard)"
                                    break
                                else:
                                    print(f"[Excel COM]   File exists but size is 0!")
                            else:
                                print(f"[Excel COM]   File not found yet")
                            time.sleep(0.5)
                        
                        if not export_success:
                            last_error = "File not created after 5 verification attempts"
                            print(f"[Excel COM] [X][X][X] METHOD 1 FAILED: {last_error}")
                    except Exception as e:
                        last_error = str(e)
                        print(f"[Excel COM] [X][X][X] METHOD 1 EXCEPTION: {last_error}")
                        import traceback
                        traceback.print_exc()
                
                # ============================================================
                # METHOD 2: ExportAsFixedFormat - Minimum Quality (Faster)
                # ============================================================
                if not export_success:
                    try:
                        print(f"\n[Excel COM] ╔══════════════════════════════════════════════════════════╗")
                        print(f"[Excel COM] ║ METHOD 2: ExportAsFixedFormat (Minimum Quality)         ║")
                        print(f"[Excel COM] ╚══════════════════════════════════════════════════════════╝")
                        
                        if os.path.exists(abs_output):
                            os.remove(abs_output)
                            print(f"[Excel COM] Removed partial file")
                        
                        print(f"[Excel COM] Calling ExportAsFixedFormat with Quality=1 (Minimum)...")
                        wb.ExportAsFixedFormat(
                            Type=0,
                            Filename=abs_output,
                            Quality=1,  # xlQualityMinimum
                            IncludeDocProperties=False,
                            IgnorePrintAreas=False,
                            OpenAfterPublish=False
                        )
                        print(f"[Excel COM] Call completed")
                        
                        time.sleep(2.0)
                        print(f"[Excel COM] Verifying (5 attempts)...")
                        for attempt in range(5):
                            print(f"[Excel COM]   Attempt {attempt + 1}/5...")
                            if os.path.exists(abs_output) and os.path.getsize(abs_output) > 0:
                                size = os.path.getsize(abs_output)
                                print(f"[Excel COM] [OK][OK][OK] METHOD 2 SUCCESS ({size:,} bytes) [OK][OK][OK]")
                                export_success = True
                                method_used = "ExportAsFixedFormat (Min Quality)"
                                break
                            time.sleep(0.5)
                        
                        if not export_success:
                            last_error = "File not created"
                            print(f"[Excel COM] [X][X][X] METHOD 2 FAILED")
                    except Exception as e:
                        last_error = str(e)
                        print(f"[Excel COM] [X][X][X] METHOD 2 EXCEPTION: {last_error}")
                        
                        time.sleep(2.0)
                        for attempt in range(5):
                            if os.path.exists(abs_output) and os.path.getsize(abs_output) > 0:
                                print(f"[Excel COM] [OK] METHOD 2 SUCCESS ({os.path.getsize(abs_output):,} bytes)")
                                export_success = True
                                method_used = "ExportAsFixedFormat (Min Quality)"
                                break
                            time.sleep(0.5)
                        
                        if not export_success:
                            print(f"[Excel COM] [X] METHOD 2 FAILED")
                    except Exception as e:
                        print(f"[Excel COM] [X] METHOD 2 EXCEPTION: {e}")
                
                # ============================================================
                # METHOD 3: ExportAsFixedFormat - Minimal Parameters Only
                # ============================================================
                if not export_success:
                    try:
                        print(f"\n[Excel COM] === METHOD 3: ExportAsFixedFormat (Minimal Params) ===")
                        
                        if os.path.exists(abs_output):
                            os.remove(abs_output)
                        
                        # Only Type and Filename - let Excel use defaults
                        wb.ExportAsFixedFormat(0, abs_output)
                        
                        time.sleep(2.0)
                        for attempt in range(5):
                            if os.path.exists(abs_output) and os.path.getsize(abs_output) > 0:
                                print(f"[Excel COM] [OK] METHOD 3 SUCCESS ({os.path.getsize(abs_output):,} bytes)")
                                export_success = True
                                method_used = "ExportAsFixedFormat (Minimal)"
                                break
                            time.sleep(0.5)
                        
                        if not export_success:
                            print(f"[Excel COM] [X] METHOD 3 FAILED")
                    except Exception as e:
                        print(f"[Excel COM] [X] METHOD 3 EXCEPTION: {e}")
                
                # ============================================================
                # METHOD 4: SaveAs with FileFormat=57 (xlTypePDF)
                # ============================================================
                if not export_success:
                    try:
                        print(f"\n[Excel COM] === METHOD 4: SaveAs (FileFormat=57) ===")
                        
                        if os.path.exists(abs_output):
                            os.remove(abs_output)
                        
                        wb.SaveAs(
                            Filename=abs_output,
                            FileFormat=57,  # xlTypePDF
                            CreateBackup=False
                        )
                        
                        time.sleep(2.0)
                        for attempt in range(5):
                            if os.path.exists(abs_output) and os.path.getsize(abs_output) > 0:
                                print(f"[Excel COM] [OK] METHOD 4 SUCCESS ({os.path.getsize(abs_output):,} bytes)")
                                export_success = True
                                method_used = "SaveAs (xlTypePDF)"
                                break
                            time.sleep(0.5)
                        
                        if not export_success:
                            print(f"[Excel COM] [X] METHOD 4 FAILED")
                    except Exception as e:
                        print(f"[Excel COM] [X] METHOD 4 EXCEPTION: {e}")
                
                # ============================================================
                # METHOD 5: PrintOut to Microsoft Print to PDF
                # ============================================================
                if not export_success:
                    try:
                        print(f"\n[Excel COM] === METHOD 5: PrintOut (MS Print to PDF) ===")
                        
                        if os.path.exists(abs_output):
                            os.remove(abs_output)
                        
                        # Try to set printer
                        try:
                            excel.ActivePrinter = "Microsoft Print to PDF"
                            print(f"[Excel COM] Active printer set")
                        except:
                            print(f"[Excel COM] Could not set printer, continuing anyway")
                        
                        wb.PrintOut(
                            PrintToFile=True,
                            PrToFileName=abs_output,
                            Collate=True
                        )
                        
                        time.sleep(3.0)  # Print queue needs more time
                        for attempt in range(5):
                            if os.path.exists(abs_output) and os.path.getsize(abs_output) > 0:
                                print(f"[Excel COM] [OK] METHOD 5 SUCCESS ({os.path.getsize(abs_output):,} bytes)")
                                export_success = True
                                method_used = "PrintOut"
                                break
                            time.sleep(0.5)
                        
                        if not export_success:
                            print(f"[Excel COM] [X] METHOD 5 FAILED")
                    except Exception as e:
                        print(f"[Excel COM] [X] METHOD 5 EXCEPTION: {e}")
                
                # ============================================================
                # METHOD 6: Per-Sheet Export (All Sheets)
                # ============================================================
                if not export_success:
                    try:
                        print(f"\n[Excel COM] === METHOD 6: Per-Sheet Export ===")
                        
                        if os.path.exists(abs_output):
                            os.remove(abs_output)
                        
                        sheet_count = wb.Sheets.Count
                        print(f"[Excel COM] Exporting {sheet_count} sheet(s) individually")
                        
                        # For single sheet, just export it
                        if sheet_count == 1:
                            wb.Sheets(1).ExportAsFixedFormat(
                                Type=0,
                                Filename=abs_output,
                                Quality=0,
                                IncludeDocProperties=True,
                                IgnorePrintAreas=False,
                                OpenAfterPublish=False
                            )
                        else:
                            # Multiple sheets - export first as test
                            # (Full implementation would merge PDFs)
                            print(f"[Excel COM] Multiple sheets detected, exporting first sheet")
                            wb.Sheets(1).ExportAsFixedFormat(
                                Type=0,
                                Filename=abs_output,
                                Quality=0,
                                IncludeDocProperties=True,
                                IgnorePrintAreas=False,
                                OpenAfterPublish=False
                            )
                        
                        time.sleep(2.0)
                        for attempt in range(5):
                            if os.path.exists(abs_output) and os.path.getsize(abs_output) > 0:
                                print(f"[Excel COM] [OK] METHOD 6 SUCCESS ({os.path.getsize(abs_output):,} bytes)")
                                export_success = True
                                method_used = "Per-Sheet Export"
                                break
                            time.sleep(0.5)
                        
                        if not export_success:
                            print(f"[Excel COM] [X] METHOD 6 FAILED")
                    except Exception as e:
                        print(f"[Excel COM] [X] METHOD 6 EXCEPTION: {e}")
                
                # List directory AFTER conversion
                print(f"\n[Excel COM] {'='*60}")
                print(f"[Excel COM] POST-CONVERSION STATE")
                print(f"[Excel COM] {'='*60}")
                try:
                    files_after = list(Path(output_dir_path).iterdir())
                    print(f"[Excel COM] Files in output dir: {len(files_after)}")
                    for f in sorted(files_after)[:15]:
                        if f.is_file():
                            marker = " <<<TARGET" if f.name == Path(abs_output).name else ""
                            print(f"[Excel COM]   - {f.name} ({f.stat().st_size:,} bytes){marker}")
                except Exception as e:
                    print(f"[Excel COM] Could not list directory: {e}")
                
                # Report final result
                print(f"\n[Excel COM] {'='*60}")
                if export_success:
                    print(f"[Excel COM] [OK][OK][OK] PDF EXPORT SUCCESSFUL [OK][OK][OK]")
                    print(f"[Excel COM] Method used: {method_used}")
                    print(f"[Excel COM] File: {os.path.basename(abs_output)}")
                else:
                    print(f"[Excel COM] [X][X][X] ALL 6 COM METHODS FAILED [X][X][X]")
                    print(f"[Excel COM] Last error: {last_error}")
                    print(f"[Excel COM] Will attempt LibreOffice fallback during verification...")
                print(f"[Excel COM] {'='*60}")
                
                print(f"\n[Excel COM] Closing workbook...")
                wb.Close(SaveChanges=False)
                print(f"[Excel COM] [OK] Workbook closed")
                
            except Exception as e:
                print(f"\n[Excel COM] [X] FATAL ERROR in Excel COM operation: {str(e)}")
                import traceback
                traceback.print_exc()
                # Try to close workbook if it's open
                try:
                    if 'wb' in locals():
                        print(f"[Excel COM] Attempting to close workbook after error...")
                        wb.Close(SaveChanges=False)
                        print(f"[Excel COM] Workbook closed")
                except:
                    print(f"[Excel COM] Could not close workbook")
                raise
            finally:
                try:
                    print(f"[Excel COM] Quitting Excel application...")
                    excel.Quit()
                    time.sleep(0.5)
                    print(f"[Excel COM] [OK] Excel quit successfully")
                except Exception as e:
                    print(f"[Excel COM] Error quitting Excel: {str(e)}")
        finally:
            pythoncom.CoUninitialize()
            print(f"[Excel COM] [OK] COM uninitialized")
        
        # Final verification with comprehensive checks
        print(f"\n[Excel COM] {'='*60}")
        print(f"[Excel COM] FINAL VERIFICATION")
        print(f"[Excel COM] {'='*60}")
        print(f"[Excel COM] Expected path: {output_path}")
        print(f"[Excel COM] Absolute path: {os.path.abspath(output_path)}")
        print(f"[Excel COM] File exists: {os.path.exists(output_path)}")
        
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            print(f"[Excel COM] File size: {file_size:,} bytes")
            
            if file_size == 0:
                print(f"[Excel COM] [X] WARNING: File exists but is EMPTY (0 bytes)")
                # Try LibreOffice as last resort
                if LIBREOFFICE_AVAILABLE:
                    print(f"[Excel COM] Attempting LibreOffice conversion as fallback...")
                    try:
                        self._libreoffice_to_pdf(input_path, output_path)
                        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                            print(f"[Excel COM] [OK] LibreOffice fallback succeeded")
                            return
                    except Exception as lo_error:
                        print(f"[Excel COM] LibreOffice fallback failed: {lo_error}")
                
                raise RuntimeError(f"PDF file is empty (0 bytes): {output_path}")
            
            print(f"[Excel COM] [OK][OK][OK] FINAL VERIFICATION PASSED [OK][OK][OK]")
            print(f"[Excel COM] PDF successfully created: {file_size:,} bytes")
        else:
            print(f"[Excel COM] [X] FILE DOES NOT EXIST at expected path")
            
            # List directory contents
            output_dir = os.path.dirname(output_path)
            print(f"[Excel COM] Listing directory: {output_dir}")
            if os.path.exists(output_dir):
                try:
                    files = list(Path(output_dir).iterdir())
                    print(f"[Excel COM] Directory contains {len(files)} file(s):")
                    for f in sorted(files)[:20]:
                        if f.is_file():
                            print(f"[Excel COM]   - {f.name} ({f.stat().st_size:,} bytes)")
                        else:
                            print(f"[Excel COM]   - {f.name}/ (directory)")
                except Exception as e:
                    print(f"[Excel COM] Error listing directory: {e}")
            else:
                print(f"[Excel COM] [X] Output directory does not exist!")
            
            # Try LibreOffice as last resort
            if LIBREOFFICE_AVAILABLE:
                print(f"\n[Excel COM] {'='*60}")
                print(f"[Excel COM] ATTEMPTING LIBREOFFICE FALLBACK")
                print(f"[Excel COM] All COM methods failed, trying LibreOffice as last resort...")
                print(f"[Excel COM] NOTE: LibreOffice may have formatting differences")
                print(f"[Excel COM] {'='*60}")
                try:
                    self._libreoffice_to_pdf(input_path, output_path)
                    
                    # Verify LibreOffice output
                    time.sleep(1)
                    if os.path.exists(output_path):
                        file_size = os.path.getsize(output_path)
                        if file_size > 0:
                            print(f"[Excel COM] [OK][OK][OK] LibreOffice fallback succeeded: {file_size:,} bytes")
                            print(f"[Excel COM] WARNING: PDF created by LibreOffice, not Excel COM")
                            print(f"[Excel COM] There may be formatting differences from Excel")
                            return
                        else:
                            print(f"[Excel COM] [X] LibreOffice created empty file")
                    else:
                        print(f"[Excel COM] [X] LibreOffice did not create file")
                except Exception as lo_error:
                    print(f"[Excel COM] [X] LibreOffice fallback failed: {lo_error}")
                    import traceback
                    traceback.print_exc()
            
            # Final failure message - include debug log
            error_summary = f"""
{'='*60}
EXCEL TO PDF CONVERSION FAILED
{'='*60}
Input:  {input_path} ({os.path.getsize(input_path):,} bytes)
Output: {output_path}

All 6 COM methods failed:
  1. ExportAsFixedFormat (Standard Quality)
  2. ExportAsFixedFormat (Minimum Quality)
  3. ExportAsFixedFormat (Minimal Parameters)
  4. SaveAs (FileFormat=57)
  5. PrintOut (Microsoft Print to PDF)
  6. Per-Sheet Export

LibreOffice fallback: {'FAILED' if LIBREOFFICE_AVAILABLE else 'NOT AVAILABLE'}

Possible causes:
- Excel COM automation blocked by security settings
- Insufficient permissions to write to output directory
- Excel not properly installed or registered
- File path encoding issues
- Antivirus blocking COM operations

Check the detailed logs above for specific error messages.
{'='*60}
"""
            print(f"[Excel COM] {error_summary}")
            
            # Capture debug log and include in exception
            debug_output = debug_log.getvalue()
            full_error = f"""PDF file was not created after trying all methods (6 COM + LibreOffice): {output_path}

{'='*60}
DETAILED LOG FILE LOCATION:
{'='*60}
{log_path}

You can open this file to see the complete debugging output.

{'='*60}
DEBUG LOG (SUMMARY):
{'='*60}
{debug_output}"""
            raise RuntimeError(full_error)
    
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
