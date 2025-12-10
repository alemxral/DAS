# Quick Setup Guide

## Installation Steps

### 1. Prerequisites
- Python 3.8 or higher installed
- Windows operating system (for .msg file support)
- Microsoft Office installed (optional, for better PDF conversion)

### 2. Quick Start (Easiest Method)

Simply double-click `start.bat` in the project root directory. This will:
- Create a virtual environment
- Install all dependencies
- Create configuration file
- Start the server

The application will be available at: http://localhost:5000

### 3. Manual Setup

If you prefer manual setup:

```bash
# Navigate to project directory
cd c:\Users\pc\autoarendt

# Create virtual environment
python -m venv venv

# Activate virtual environment
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Create configuration file
copy .env.example .env

# Start the application
python run.py
```

### 4. First Time Configuration

The default configuration works out of the box, but you can customize it:

1. Open `.env` file in a text editor
2. Modify settings as needed:
   - `PORT`: Change server port (default: 5000)
   - `DEBUG`: Set to False for production
   - `MAX_CONTENT_LENGTH`: Adjust file size limit
   - `MAX_CONCURRENT_JOBS`: Set concurrent job limit

### 5. Create Example Files

Before creating your first job, create example files:

1. **Create Excel Data File** (`examples\example_data.xlsx`):
   - Open Excel
   - Row 1: `##name##` | `##email##` | `##amount##` | `##date##`
   - Row 2: `John Smith` | `john@example.com` | `$1,000` | `2025-01-15`
   - Row 3: `Jane Doe` | `jane@example.com` | `$2,500` | `2025-01-20`
   - Save in `examples` folder

2. **Create Word Template** (`examples\example_template.docx`):
   - Open Word
   - Add text with `##variable##` placeholders
   - Example: "Dear ##name##, your amount is ##amount##"
   - Save in `examples` folder

### 6. Create Your First Job

1. Open browser: http://localhost:5000
2. Click "New Job"
3. Choose "File Path" option
4. Enter paths to your example files:
   - Template: `c:\Users\pc\autoarendt\examples\example_template.docx`
   - Data: `c:\Users\pc\autoarendt\examples\example_data.xlsx`
5. Select output format (PDF recommended)
6. Click "Create Job"
7. Wait for completion
8. Click "Download" to get results

## Troubleshooting

### Port Already in Use
If port 5000 is already in use:
1. Edit `.env` file
2. Change `PORT=5000` to another port (e.g., `PORT=5001`)
3. Restart the application

### Module Not Found Errors
```bash
# Ensure virtual environment is activated
venv\Scripts\activate

# Reinstall dependencies
pip install -r requirements.txt
```

### .msg Files Not Working
```bash
# Install pywin32 postinstall
python venv\Scripts\pywin32_postinstall.py -install

# Or install manually
pip install --force-reinstall pywin32
```

### PDF Conversion Issues
- Option 1: Install Microsoft Office
- Option 2: System will use ReportLab for basic PDF generation

### File Upload Fails
- Check file size (default limit: 100MB)
- Verify file format is supported
- Check `MAX_CONTENT_LENGTH` in `.env`

## Directory Structure After Setup

```
autoarendt/
├── venv/                 # Virtual environment (created)
├── jobs/                 # Job data (created on first job)
├── storage/              # File cache (created on first job)
├── uploads/              # Uploaded files (created on first upload)
├── app/                  # Application code
├── config/               # Configuration
├── models/               # Data models
├── services/             # Business logic
├── static/               # Frontend assets
├── templates/            # HTML templates
├── examples/             # Example files
├── .env                  # Environment config (created)
├── requirements.txt      # Dependencies
├── run.py               # Entry point
├── start.bat            # Quick start script
└── README.md            # Documentation
```

## Next Steps

1. **Test with Examples**: Use the example files to verify everything works
2. **Create Your Templates**: Build your own Word/Excel templates
3. **Prepare Your Data**: Format Excel files with `##variable##` in first row
4. **Run Jobs**: Process your documents through the system
5. **Download Results**: Get ZIP files with generated documents

## Getting Help

- Check `README.md` for detailed documentation
- Review `examples/README.md` for template examples
- Check console logs for error messages
- Review job metadata in `jobs/<job-id>/metadata.json`

## Production Deployment

For production use:

1. Set `DEBUG=False` in `.env`
2. Change `SECRET_KEY` to a secure random value
3. Set `CORS_ORIGINS` to specific domains
4. Use a production WSGI server (gunicorn, waitress)
5. Set up proper logging
6. Configure file backup strategy
7. Implement authentication if needed

## Support

For issues or questions:
- Check error messages in browser console
- Review terminal output
- Check job metadata files
- Verify file formats and paths
