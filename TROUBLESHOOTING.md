# Troubleshooting Guide

## Common Issues and Solutions

### Installation Issues

#### Issue: `pip install` fails
**Symptoms:**
- Error installing dependencies
- Package conflicts
- Build errors

**Solutions:**
```bash
# Update pip
python -m pip install --upgrade pip

# Install with verbose output
pip install -r requirements.txt -v

# Install individually if needed
pip install Flask flask-cors python-docx openpyxl pandas reportlab

# For Windows users having issues with pywin32:
pip install --force-reinstall pywin32
python venv\Scripts\pywin32_postinstall.py -install
```

#### Issue: Virtual environment not activating
**Symptoms:**
- `venv\Scripts\activate` does nothing
- Wrong Python version

**Solutions:**
```bash
# Ensure Python is in PATH
python --version

# Recreate virtual environment
rmdir /s venv
python -m venv venv
venv\Scripts\activate
```

### Server Issues

#### Issue: Port 5000 already in use
**Symptoms:**
```
OSError: [WinError 10048] Only one usage of each socket address
```

**Solutions:**
1. Change port in `.env`:
   ```env
   PORT=5001
   ```

2. Or kill process using port 5000:
   ```bash
   # Find process
   netstat -ano | findstr :5000
   
   # Kill process (replace PID)
   taskkill /PID <PID> /F
   ```

#### Issue: Server won't start
**Symptoms:**
- Import errors
- Module not found

**Solutions:**
```bash
# Ensure virtual environment is activated
venv\Scripts\activate

# Check Python path
python -c "import sys; print(sys.executable)"

# Reinstall dependencies
pip install -r requirements.txt

# Check for syntax errors
python -m py_compile run.py
```

#### Issue: 404 on all API endpoints
**Symptoms:**
- API calls return 404
- Routes not found

**Solutions:**
- Check Flask is running
- Verify URL: `http://localhost:5000/api/jobs`
- Check browser console for CORS errors
- Ensure routes.py is properly registered

### File Processing Issues

#### Issue: COM initialization error (CoInitialize)
**Symptoms:**
```
Error processing row 1: (-2147221008, 'No se ha llamado a CoInitialize.', None, None)
pywintypes.com_error: (-2147221008, 'CoInitialize has not been called.', None, None)
```

**Cause:**
Windows COM automation (used for Word, Excel, Outlook integration) requires explicit initialization in each thread.

**Solution:**
✅ **FIXED** - The application now automatically calls `pythoncom.CoInitialize()` before all COM operations and `pythoncom.CoUninitialize()` after completion. This fix is in:
- `services/format_converter.py` (all Word/Excel/Outlook conversions)
- `services/template_processor.py` (MSG template processing)

If you still encounter this error:
1. Ensure pywin32 is properly installed:
   ```bash
   pip install --force-reinstall pywin32
   python venv\Scripts\pywin32_postinstall.py -install
   ```
2. Restart the Flask server after installing pywin32
3. Check that Microsoft Office is installed on Windows

#### Issue: Excel file not parsing
**Symptoms:**
- "Error parsing Excel file"
- Empty data returned

**Solutions:**
1. Check Excel format:
   - First row must have `##variable##` format
   - At least 2 rows (header + data)
   - Save as .xlsx format

2. Verify file:
   ```python
   import pandas as pd
   df = pd.read_excel('your_file.xlsx')
   print(df.head())
   ```

#### Issue: Variables not being replaced
**Symptoms:**
- Output has `##variable##` still visible
- Placeholders not substituted

**Solutions:**
1. Check variable names match:
   - Excel: `##name##`
   - Template: `##name##` (exact match, case-sensitive)

2. Check for extra spaces:
   - Wrong: `## name ##`
   - Correct: `##name##`

3. Verify data row:
   ```python
   # Test data parsing
   from services.document_parser import DocumentParser
   parser = DocumentParser()
   data = parser.parse_excel_data('data.xlsx')
   print(data['variables'])
   print(data['data'])
   ```

#### Issue: .msg files not processing
**Symptoms:**
- "pywin32 is required for .msg templates"
- COM errors

**Solutions:**
1. Windows only feature - ensure on Windows

2. Install/reinstall pywin32:
   ```bash
   pip install --force-reinstall pywin32
   python venv\Scripts\pywin32_postinstall.py -install
   ```

3. Ensure Outlook is installed

4. Run as administrator if needed

#### Issue: PDF conversion fails
**Symptoms:**
- "PDF conversion requires either pywin32 or docx2pdf"
- PDF files not generated

**Solutions:**
1. **Best solution:** Install Microsoft Office
   - Enables COM automation
   - Better quality PDFs

2. **Alternative:** Use ReportLab fallback
   - Already included in requirements
   - Basic PDF generation

3. **For docx2pdf:**
   ```bash
   pip install docx2pdf
   ```

### Job Issues

#### Issue: Job stuck in PROCESSING
**Symptoms:**
- Job doesn't complete
- Progress bar frozen

**Solutions:**
1. Check console logs for errors

2. Check job metadata:
   ```bash
   type jobs\<job-id>\metadata.json
   ```

3. Restart server and recreate job

4. Check for file permission issues

#### Issue: Job fails immediately
**Symptoms:**
- Status: FAILED
- Error message in job card

**Solutions:**
1. Read error message carefully

2. Common causes:
   - File not found
   - Invalid file format
   - Missing variables
   - Insufficient permissions

3. Check job metadata for detailed error:
   ```json
   {
     "error_message": "Detailed error here"
   }
   ```

### Upload Issues

#### Issue: File upload fails
**Symptoms:**
- "Invalid file format"
- "File too large"

**Solutions:**
1. Check file size:
   - Default limit: 100MB
   - Increase in `.env`: `MAX_CONTENT_LENGTH=209715200` (200MB)

2. Check file extension:
   - Templates: .docx, .xlsx, .msg
   - Data: .xlsx, .xls

3. Check file is not corrupted:
   - Open file manually
   - Try saving a new copy

#### Issue: Path input not working
**Symptoms:**
- "File not found"
- Path errors

**Solutions:**
1. Use absolute paths:
   - ✓ `C:\Users\pc\documents\template.docx`
   - ✗ `documents\template.docx`

2. Use forward slashes or double backslashes:
   - ✓ `C:/Users/pc/template.docx`
   - ✓ `C:\\Users\\pc\\template.docx`
   - ✗ `C:\Users\pc\template.docx` (may fail)

3. Check file permissions

### Frontend Issues

#### Issue: Dashboard not loading
**Symptoms:**
- Blank page
- Loading spinner forever

**Solutions:**
1. Check browser console (F12)

2. Common errors:
   ```javascript
   // CORS error
   // Solution: Check Flask CORS settings
   
   // 404 on API
   // Solution: Ensure server is running
   
   // JavaScript error
   // Solution: Clear browser cache
   ```

3. Try different browser

4. Disable browser extensions

#### Issue: Auto-refresh not working
**Symptoms:**
- Jobs don't update automatically
- Need to manually refresh page

**Solutions:**
1. Check browser console for errors

2. Verify JavaScript is enabled

3. Check interval is running:
   ```javascript
   // In browser console
   console.log('Refresh interval:', refreshInterval);
   ```

#### Issue: Preview not working
**Symptoms:**
- Preview button does nothing
- Files won't open

**Solutions:**
1. For PDFs:
   - Allow popups for localhost
   - Check PDF.js loaded

2. For Office files:
   - Files will download instead
   - Open with appropriate program

### Performance Issues

#### Issue: Slow processing
**Symptoms:**
- Jobs take very long
- Server unresponsive

**Solutions:**
1. Reduce dataset size for testing

2. Check system resources:
   - CPU usage
   - Memory usage
   - Disk space

3. Process fewer formats at once

4. Consider async processing (Celery)

#### Issue: Large output files
**Symptoms:**
- ZIP file very large
- Download fails

**Solutions:**
1. Reduce number of records

2. Process in batches

3. Compress output more:
   - Use PDF instead of Word
   - Optimize images in templates

### Database/Storage Issues

#### Issue: Jobs disappear after restart
**Symptoms:**
- Job list empty after server restart

**Solutions:**
1. Check jobs/ directory exists

2. Check metadata.json files:
   ```bash
   dir /s jobs\*.json
   ```

3. Verify file permissions

4. Check for disk space

#### Issue: Old files accumulating
**Symptoms:**
- Disk space filling up
- Too many old jobs

**Solutions:**
1. Manual cleanup:
   ```bash
   # Delete old jobs
   rmdir /s jobs\<old-job-id>
   ```

2. Implement automatic cleanup:
   ```python
   from utils.helpers import cleanup_old_files
   cleanup_old_files('jobs', days=7)
   ```

### Windows-Specific Issues

#### Issue: Permission denied errors
**Symptoms:**
- Cannot write to directory
- Access denied

**Solutions:**
1. Run as administrator

2. Check folder permissions

3. Disable antivirus temporarily

4. Check file is not open in another program

#### Issue: Path too long error
**Symptoms:**
- Error with file paths
- Path limit exceeded

**Solutions:**
1. Move project closer to root:
   - Good: `C:\autoarendt`
   - Bad: `C:\Users\Username\Documents\Projects\Python\autoarendt`

2. Enable long paths (Windows 10+):
   - Run as admin:
   ```
   reg add HKLM\SYSTEM\CurrentControlSet\Control\FileSystem /v LongPathsEnabled /t REG_DWORD /d 1 /f
   ```

## Debugging Tips

### Enable Debug Mode
In `.env`:
```env
DEBUG=True
```

### Check Logs
Monitor console output for errors:
```bash
python run.py > output.log 2>&1
```

### Test Individual Components

```python
# Test FileTracker
from services.file_tracker import FileTracker
tracker = FileTracker('storage')
info = tracker.track_file('path/to/file.docx')
print(info)

# Test DocumentParser
from services.document_parser import DocumentParser
parser = DocumentParser()
data = parser.parse_excel_data('data.xlsx')
print(data)

# Test TemplateProcessor
from services.template_processor import TemplateProcessor
processor = TemplateProcessor()
vars = processor.extract_template_variables('template.docx')
print(vars)
```

### Verify Installation

```bash
# Check Python version
python --version

# Check installed packages
pip list

# Check package imports
python -c "import flask; import openpyxl; import docx; print('OK')"
```

### Network Issues

If running on a different machine:

1. Change HOST in `.env`:
   ```env
   HOST=0.0.0.0  # Allow external access
   ```

2. Update CORS:
   ```env
   CORS_ORIGINS=http://your-ip:5000
   ```

3. Check firewall allows port 5000

## Getting Additional Help

1. **Check logs:**
   - Console output
   - Browser console (F12)
   - Job metadata files

2. **Verify setup:**
   - Run CHECKLIST.md items
   - Ensure all dependencies installed
   - Check file permissions

3. **Test with examples:**
   - Create simple test files
   - Use minimal data set
   - Try one format at a time

4. **Isolate the issue:**
   - Does it work with example files?
   - Does it work with different browser?
   - Does it work after fresh install?

## Still Having Issues?

Create a bug report with:
- Error message (full text)
- Steps to reproduce
- Python version
- OS version
- Installed packages (`pip list`)
- Console logs
- Job metadata (if applicable)
