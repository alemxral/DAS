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
from services.license_validator import LicenseValidator
import base64

class Api:
    """API class for PyWebView JavaScript bridge."""
    
    def save_file(self, filename, data_base64):
        """Save file dialog for PyWebView.
        
        Args:
            filename: Suggested filename
            data_base64: Base64 encoded file content
            
        Returns:
            dict with success status and path
        """
        try:
            # Decode base64 data
            file_data = base64.b64decode(data_base64)
            
            # Show save dialog
            result = webview.windows[0].create_file_dialog(
                webview.SAVE_DIALOG,
                save_filename=filename
            )
            
            if result:
                # User selected a location
                save_path = result[0] if isinstance(result, tuple) else result
                
                # Write file
                with open(save_path, 'wb') as f:
                    f.write(file_data)
                
                return {'success': True, 'path': save_path}
            else:
                # User cancelled
                return {'success': False, 'cancelled': True}
                
        except Exception as e:
            return {'success': False, 'error': str(e)}

# Handle PyInstaller frozen state
if getattr(sys, 'frozen', False):
    # Running in PyInstaller bundle
    BASE_DIR = Path(sys._MEIPASS)
    os.chdir(BASE_DIR)
    # Set encoding for stdout to avoid Unicode errors in frozen exe
    try:
        import io
        if hasattr(sys.stdout, 'buffer') and not isinstance(sys.stdout, io.TextIOWrapper):
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        if hasattr(sys.stderr, 'buffer') and not isinstance(sys.stderr, io.TextIOWrapper):
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    except (AttributeError, ValueError):
        # If wrapping fails, just continue without it
        pass
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
    
    # Validate license
    validator = LicenseValidator()
    is_valid, message = validator.validate()
    
    if not is_valid:
        print(f"\n{'='*60}")
        print("[X] SERVICE NOT AVAILABLE")
        print(f"{'='*60}")
        print(f"Message: {message}")
        print(f"The service is currently not available.")
        print(f"{'='*60}\n")
        
        # Show error window
        webview.create_window(
            'Service Unavailable - DAS',
            html=f"""
            <!DOCTYPE html>
            <html>
            <head>
                <title>Service Unavailable</title>
                <style>
                    body {{
                        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        display: flex;
                        align-items: center;
                        justify-content: center;
                        height: 100vh;
                        margin: 0;
                        padding: 20px;
                    }}
                    .container {{
                        background: white;
                        border-radius: 20px;
                        box-shadow: 0 20px 60px rgba(0,0,0,0.3);
                        padding: 60px;
                        max-width: 500px;
                        text-align: center;
                    }}
                    .icon {{
                        font-size: 80px;
                        margin-bottom: 20px;
                    }}
                    h1 {{
                        color: #e53e3e;
                        margin: 0 0 10px 0;
                        font-size: 32px;
                    }}
                    .subtitle {{
                        color: #718096;
                        margin: 0 0 30px 0;
                        font-size: 18px;
                    }}
                    .message {{
                        color: #4a5568;
                        font-size: 16px;
                        margin: 20px 0;
                        padding: 20px;
                        background: #f7fafc;
                        border-radius: 10px;
                    }}
                    button {{
                        background: #667eea;
                        color: white;
                        border: none;
                        padding: 12px 30px;
                        border-radius: 8px;
                        font-size: 14px;
                        font-weight: 600;
                        cursor: pointer;
                        margin-top: 20px;
                    }}
                    button:hover {{
                        background: #5a67d8;
                    }}
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="icon">[X]</div>
                    <h1>Service Not Available</h1>
                    <p class="subtitle">The application cannot start at this time</p>
                    
                    <div class="message">
                        {message}
                    </div>
                    
                    <button onclick="window.close()">Close</button>
                </div>
            </body>
            </html>
            """,
            width=600,
            height=500,
            resizable=False
        )
        webview.start()
        sys.exit(1)
    
    print(f"[License] [OK] {message}")
    
    # Start Flask in background thread
    try:
        flask_thread = threading.Thread(target=start_flask, daemon=True)
        flask_thread.start()
        
        # Wait for Flask to start
        time.sleep(3)
        
        # Find icon file
        icon_path = BASE_DIR / 'static' / 'icon.png'
        if not icon_path.exists():
            icon_path = None
        
        # Create API instance
        api = Api()
        
        # Create native desktop window
        webview.create_window(
            'DAS - Document Automation System',
            'http://127.0.0.1:5000',
            width=1400,
            height=900,
            resizable=True,
            min_size=(800, 600),
            js_api=api
        )
        
        webview.start()
    
    except Exception as e:
        print(f"[Error] Failed to start application: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == '__main__':
    main()
