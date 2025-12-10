"""
Main Entry Point for Desktop Application
Uses pywebview to create a native desktop window.
"""
import sys
import os
from pathlib import Path
import threading
import time
import webview
from app import create_app

# Handle PyInstaller frozen state
if getattr(sys, 'frozen', False):
    # Running in PyInstaller bundle
    BASE_DIR = Path(sys._MEIPASS)
    os.chdir(BASE_DIR)
else:
    # Running in normal Python environment
    BASE_DIR = Path(__file__).parent

def start_flask():
    """Start Flask server in a separate thread."""
    app = create_app()
    app.run(host='127.0.0.1', port=5000, debug=False, use_reloader=False, threaded=True)

def main():
    """Initialize and start the desktop application."""
    print("Starting Document Automation System...")
    
    # Start Flask in background thread
    flask_thread = threading.Thread(target=start_flask, daemon=True)
    flask_thread.start()
    
    # Wait for Flask to start
    time.sleep(3)
    
    # Find icon file
    icon_path = BASE_DIR / 'static' / 'icon.png'
    if not icon_path.exists():
        icon_path = None
    
    # Create native desktop window
    webview.create_window(
        'DAS - Document Automation System',
        'http://127.0.0.1:5000',
        width=1400,
        height=900,
        resizable=True,
        min_size=(800, 600)
    )
    
    webview.start()

if __name__ == '__main__':
    main()
