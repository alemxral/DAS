# Document Automation System

A powerful client-server Python application for automating document generation from Excel data and multi-format templates.

## Features

- **Multi-Format Support**: Templates in Word (.docx), Excel (.xlsx), and Outlook (.msg) formats
- **Flexible Output**: Generate documents in PDF, Word, Excel (single sheet or workbook), and MSG formats
- **Variable Substitution**: Use `##variable##` format in Excel data and templates for automatic replacement
- **File Tracking**: SHA-256 based change detection ensures local copies are always up-to-date
- **Job Management**: Track multiple document generation jobs with real-time progress monitoring
- **Web Interface**: Professional dashboard built with Flask, JavaScript, and Tailwind CSS
- **File Preview**: View generated documents directly in the browser
- **Batch Processing**: Generate multiple documents from multiple data rows automatically

## Architecture

### Backend
- **Flask API**: RESTful API for job management and file operations
- **Services**:
  - `FileTracker`: SHA-256 based file change detection
  - `DocumentParser`: Excel data parsing with variable extraction
  - `TemplateProcessor`: Multi-format template processing
  - `FormatConverter`: Document format conversion
  - `JobManager`: Job lifecycle and metadata management

### Frontend
- **Tailwind CSS**: Modern, responsive UI
- **JavaScript**: Dynamic job management and real-time updates
- **PDF.js**: In-browser PDF preview

## Project Structure

```
autoarendt/
├── app/                      # Flask application
│   ├── __init__.py          # App initialization
│   └── routes.py            # API endpoints
├── config/                   # Configuration
│   └── config.py            # Settings and environment
├── models/                   # Data models
│   └── job.py               # Job class definition
├── services/                 # Business logic
│   ├── file_tracker.py      # File tracking and SHA validation
│   ├── document_parser.py   # Excel parsing
│   ├── template_processor.py # Template processing
│   ├── format_converter.py  # Format conversion
│   └── job_manager.py       # Job orchestration
├── static/                   # Frontend assets
│   ├── js/
│   │   └── app.js           # Frontend JavaScript
│   └── css/
├── templates/                # HTML templates
│   └── index.html           # Dashboard
├── jobs/                     # Job data (generated)
├── storage/                  # Tracked file copies (generated)
├── uploads/                  # Uploaded files (generated)
├── requirements.txt          # Python dependencies
├── .env.example             # Environment template
├── .gitignore               # Git ignore rules
├── run.py                   # Application entry point
└── README.md                # This file
```

## Installation

### Prerequisites
- Python 3.8 or higher
- Windows OS (for .msg file support via COM automation)
- Microsoft Office (optional, for better PDF conversion)

### Setup

1. **Clone or navigate to the project directory**
   ```bash
   cd c:\Users\pc\autoarendt
   ```

2. **Create virtual environment**
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

4. **Configure environment**
   ```bash
   copy .env.example .env
   ```
   Edit `.env` file with your settings.

5. **Run the application**
   ```bash
   python run.py
   ```

6. **Access the dashboard**
   Open your browser and navigate to: `http://localhost:5000`

## Usage

### Creating a Job

1. **Click "New Job"** in the dashboard
2. **Select Template**: Upload a template file or provide a file path
   - Supported formats: `.docx`, `.xlsx`, `.msg`
   - Use `##variable##` format for placeholders
3. **Select Data File**: Upload Excel file with variables in first row
   - First row format: `##name##`, `##email##`, `##amount##`, etc.
   - Each subsequent row contains one record to process
4. **Choose Output Formats**: Select desired output formats (PDF, Word, Excel, MSG)
5. **Click "Create Job"**: Job will automatically start processing

### Excel Data Format

Your Excel data file should have:
- **First row**: Variable names in `##variable##` format
- **Subsequent rows**: Data values for each variable

Example:
```
| ##name##    | ##email##           | ##amount## |
|-------------|---------------------|------------|
| John Smith  | john@example.com    | $1,000     |
| Jane Doe    | jane@example.com    | $2,500     |
```

### Template Format

Create templates with placeholders matching your Excel variables:

**Word Template Example:**
```
Dear ##name##,

Thank you for your order of ##amount##.
We will contact you at ##email##.
```

### Job Statuses

- **Pending**: Job created, waiting to process
- **Processing**: Currently generating documents
- **Completed**: All documents generated successfully
- **Failed**: Error occurred during processing

### Downloading Results

When a job completes:
1. Click **"Download"** to get a ZIP file with all generated documents
2. Click **"View"** to preview individual files in the browser

## API Endpoints

### Jobs
- `GET /api/jobs` - List all jobs
- `GET /api/jobs/<id>` - Get job details
- `POST /api/jobs` - Create new job
- `POST /api/jobs/<id>/process` - Start job processing
- `DELETE /api/jobs/<id>` - Delete job

### Files
- `GET /api/jobs/<id>/download` - Download job output ZIP
- `GET /api/jobs/<id>/files` - List job output files
- `GET /api/jobs/<id>/preview/<path>` - Preview specific file

### Dashboard
- `GET /api/dashboard/stats` - Get dashboard statistics
- `GET /api/formats` - Get available output formats
- `GET /api/health` - Health check

## Configuration

Edit `.env` file to customize:

```env
# Server settings
HOST=0.0.0.0
PORT=5000
DEBUG=True

# File size limits
MAX_CONTENT_LENGTH=104857600  # 100MB

# Processing
MAX_CONCURRENT_JOBS=5
JOB_TIMEOUT=3600
```

## File Tracking

The system automatically:
- Tracks file changes using SHA-256 hashing
- Maintains local copies in `storage/` directory
- Updates copies when source files change
- Never modifies original files

## Troubleshooting

### .msg Files Not Working
- Requires Windows OS with Outlook installed
- Ensure `pywin32` is installed: `pip install pywin32`
- Run: `python venv\Scripts\pywin32_postinstall.py -install`

### PDF Conversion Issues
- Install Microsoft Office for best results
- Alternative: System uses ReportLab for basic PDF generation

### Large Files
- Adjust `MAX_CONTENT_LENGTH` in `.env`
- Consider increasing `JOB_TIMEOUT` for large datasets

## Development

### Adding New Template Formats
1. Update `TemplateProcessor` in `services/template_processor.py`
2. Add extraction and processing methods
3. Update `ALLOWED_TEMPLATE_EXTENSIONS` in config

### Adding New Output Formats
1. Update `FormatConverter` in `services/format_converter.py`
2. Add conversion method
3. Update `AVAILABLE_OUTPUT_FORMATS` in config

## License

This project is provided as-is for document automation purposes.

## Support

For issues or questions, please check:
- Error messages in the browser console
- Application logs in the terminal
- Job metadata files in `jobs/<job-id>/metadata.json`
