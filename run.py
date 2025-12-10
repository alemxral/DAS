"""
Main Application Entry Point
"""
from app import create_app
from config.config import Config

# Create Flask application
app = create_app(Config)

if __name__ == '__main__':
    print("=" * 60)
    print("Document Automation System")
    print("=" * 60)
    print(f"Server starting on http://{Config.HOST}:{Config.PORT}")
    print(f"Debug mode: {Config.DEBUG}")
    print("=" * 60)
    
    app.run(
        host=Config.HOST,
        port=Config.PORT,
        debug=Config.DEBUG
    )
