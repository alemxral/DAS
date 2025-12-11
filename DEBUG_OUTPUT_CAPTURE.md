# Debug Output Capture for UI Display

## Problem
User reported: "No records were processed successfully" error, but debugging information was only visible in the console, not in the web UI error messages.

## Solution
Implemented automatic capture of ALL debugging output to display in UI error messages.

## How It Works

### 1. Output Redirection (`format_converter.py`)
- **TeeOutput class**: Writes to both console AND a StringIO buffer simultaneously
- All `print()` statements are automatically captured without modification
- Original stdout is restored after conversion completes

### 2. Error Message Enhancement
When Excel to PDF conversion fails, the exception message includes:
```
PDF file was not created after trying all methods...

============================================================
DEBUG LOG:
============================================================
[Full debugging output from all 6 conversion attempts]
```

### 3. UI Display (`app.js` line 847)
Error messages are displayed in red in the job details modal automatically.

## What You'll See in UI Errors

When an Excel template fails, the error message will now include:

1. **Method Attempts**: All 6 COM methods tried (ExportAsFixedFormat variants, SaveAs, PrintOut, Per-Sheet)
2. **Verification Steps**: 5 verification attempts per method with timestamps
3. **File System Info**: Directory listings, file sizes, permissions
4. **Error Details**: Specific COM error codes and messages
5. **LibreOffice Fallback**: Whether LibreOffice was attempted and results

## Technical Implementation

```python
# Outer wrapper captures output
def _xlsx_to_pdf_com(self, input_path, output_path, print_settings):
    debug_log = StringIO()
    sys.stdout = TeeOutput(original_stdout, debug_log)
    try:
        self._xlsx_to_pdf_com_inner(...)  # Does all the work
    finally:
        sys.stdout = original_stdout  # Always restore

# Inner method does conversion and raises error with debug log
def _xlsx_to_pdf_com_inner(..., debug_log):
    # ... all conversion attempts with print() debugging ...
    if all_methods_failed:
        full_error = f"Conversion failed\n\nDEBUG LOG:\n{debug_log.getvalue()}"
        raise RuntimeError(full_error)
```

## Benefits

✅ **No code changes needed**: All existing `print()` statements automatically captured
✅ **Always visible**: Debug output goes to BOTH console AND error message
✅ **Complete history**: Full trace of all 6 methods + LibreOffice attempts
✅ **Production ready**: stdout always restored even if exceptions occur
✅ **User friendly**: Error messages in UI contain all diagnostic information

## Testing

To verify this is working:
1. Create/run a job with an Excel template
2. If conversion fails, click on the job in the UI
3. Check the error message (red text) - it should contain:
   - "DEBUG LOG:" section
   - All 6 method attempts with fancy box formatting
   - Verification steps and file system checks
   - Specific error messages from each method

## Files Modified

- `services/format_converter.py`:
  - Added `TeeOutput` class for dual output
  - Split `_xlsx_to_pdf_com` into wrapper + inner method
  - Enhanced final RuntimeError to include debug log
  
- `services/job_manager.py`:
  - Already captures detailed errors (no changes needed)
  
- `static/js/app.js`:
  - Already displays error messages (no changes needed)
