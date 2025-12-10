"""
Configuration Module
Handles application configuration settings.
"""
import os
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables
load_dotenv()


class Config:
    """Application configuration."""
    
    # Flask settings
    SECRET_KEY = os.getenv('SECRET_KEY', 'dev-secret-key-change-in-production')
    DEBUG = os.getenv('DEBUG', 'True').lower() == 'true'
    HOST = os.getenv('HOST', '0.0.0.0')
    PORT = int(os.getenv('PORT', '5000'))
    
    # Directories
    BASE_DIR = Path(__file__).parent.parent  # Project root directory
    print(f"[CONFIG] __file__ = {__file__}")
    print(f"[CONFIG] BASE_DIR = {BASE_DIR}")
    print(f"[CONFIG] BASE_DIR (absolute) = {BASE_DIR.absolute()}")
    
    # Handle both absolute and relative paths from environment
    _jobs_dir = os.getenv('JOBS_DIR', 'jobs')
    _storage_dir = os.getenv('STORAGE_DIR', 'storage')
    _upload_dir = os.getenv('UPLOAD_DIR', 'uploads')
    
    # Convert to absolute paths if they're relative
    JOBS_DIR = str(BASE_DIR / _jobs_dir) if not Path(_jobs_dir).is_absolute() else _jobs_dir
    STORAGE_DIR = str(BASE_DIR / _storage_dir) if not Path(_storage_dir).is_absolute() else _storage_dir
    UPLOAD_DIR = str(BASE_DIR / _upload_dir) if not Path(_upload_dir).is_absolute() else _upload_dir
    
    print(f"[CONFIG] JOBS_DIR = {JOBS_DIR}")
    print(f"[CONFIG] STORAGE_DIR = {STORAGE_DIR}")
    print(f"[CONFIG] UPLOAD_DIR = {UPLOAD_DIR}")
    
    # File upload settings
    MAX_CONTENT_LENGTH = int(os.getenv('MAX_CONTENT_LENGTH', 100 * 1024 * 1024))  # 100MB default
    ALLOWED_TEMPLATE_EXTENSIONS = {'.docx', '.xlsx', '.msg'}
    ALLOWED_DATA_EXTENSIONS = {'.xlsx', '.xls'}
    
    # Output format settings
    AVAILABLE_OUTPUT_FORMATS = ['pdf', 'pdf_merged', 'word', 'excel', 'excel_workbook', 'msg']
    
    # Processing settings
    MAX_CONCURRENT_JOBS = int(os.getenv('MAX_CONCURRENT_JOBS', '5'))
    JOB_TIMEOUT = int(os.getenv('JOB_TIMEOUT', '3600'))  # 1 hour default
    
    # CORS settings
    CORS_ORIGINS = os.getenv('CORS_ORIGINS', '*')
    
    @staticmethod
    def init_app(app):
        """Initialize application with configuration."""
        # Create directories if they don't exist
        for directory in [Config.JOBS_DIR, Config.STORAGE_DIR, Config.UPLOAD_DIR]:
            Path(directory).mkdir(parents=True, exist_ok=True)


# Configuration dictionary
config = {
    'development': Config,
    'production': Config,
    'default': Config
}
