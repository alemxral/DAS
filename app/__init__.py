"""
Flask Application Initialization
"""
import os
import sys
from pathlib import Path
from flask import Flask, render_template
from flask_cors import CORS
from config.config import Config


def create_app(config_class=Config):
    """
    Create and configure Flask application.
    
    Args:
        config_class: Configuration class to use
        
    Returns:
        Configured Flask application
    """
    # Get the project root directory - handle PyInstaller
    if getattr(sys, 'frozen', False):
        # Running in PyInstaller bundle
        project_root = Path(sys._MEIPASS)
    else:
        # Running in normal Python environment
        project_root = Path(__file__).parent.parent
    
    # Create Flask app with correct template and static folders
    app = Flask(
        __name__,
        template_folder=str(project_root / 'templates'),
        static_folder=str(project_root / 'static')
    )
    app.config.from_object(config_class)
    
    # Initialize configuration
    config_class.init_app(app)
    
    # Debug: Print configuration paths
    print(f"Configuration loaded:")
    print(f"  BASE_DIR: {config_class.BASE_DIR}")
    print(f"  JOBS_DIR: {app.config['JOBS_DIR']}")
    print(f"  STORAGE_DIR: {app.config['STORAGE_DIR']}")
    print(f"  UPLOAD_DIR: {app.config['UPLOAD_DIR']}")
    
    # Enable CORS
    CORS(app, resources={r"/api/*": {"origins": app.config['CORS_ORIGINS']}})
    
    # Register blueprints
    from app.routes import api_bp
    app.register_blueprint(api_bp)
    
    # Main route
    @app.route('/')
    def index():
        return render_template('index.html')
    
    # Error handlers
    @app.errorhandler(404)
    def not_found(error):
        # Check if this is an API request
        from flask import request
        if request.path.startswith('/api/'):
            return {'error': 'Not found'}, 404
        return render_template('index.html')  # SPA fallback
    
    @app.errorhandler(500)
    def internal_error(error):
        return {'error': 'Internal server error'}, 500
    
    @app.errorhandler(413)
    def too_large(error):
        return {'error': 'File too large'}, 413
    
    return app
