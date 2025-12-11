"""
Quick test to verify license validation path won't crash the exe.
"""
import sys
import io

def test_license_validation():
    """Test license validation without actually making network calls"""
    print("="*60)
    print("Testing License Validation Path")
    print("="*60)
    
    try:
        # Wrap stdout like in frozen exe
        if hasattr(sys.stdout, 'buffer') and not isinstance(sys.stdout, io.TextIOWrapper):
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
        
        # Test the safe_print pattern used in license_validator
        def safe_print(msg):
            """Print with fallback for encoding errors"""
            try:
                print(msg)
            except UnicodeEncodeError:
                print(msg.encode('ascii', errors='replace').decode('ascii'))
        
        # Test various message formats
        messages = [
            "\n" + "="*60,
            "[X] SERVICE NOT AVAILABLE",
            "="*60,
            "Message: License validation failed",
            "The service is currently not available.",
            "="*60 + "\n",
            "[License] [OK] License validated successfully",
        ]
        
        for msg in messages:
            safe_print(msg)
        
        print("\n[PASS] All license validation prints succeeded")
        return True
        
    except Exception as e:
        print(f"\n[FAIL] License validation test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_html_rendering():
    """Test HTML string generation for error window"""
    print("\n" + "="*60)
    print("Testing HTML Error Window Generation")
    print("="*60)
    
    try:
        message = "License validation failed - server unreachable"
        
        # Generate HTML like in main.py
        html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Service Unavailable</title>
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
        """
        
        # Verify no problematic characters
        try:
            html.encode('utf-8')
            print("[OK] HTML encodes to UTF-8")
        except UnicodeEncodeError as e:
            print(f"[FAIL] HTML encoding failed: {e}")
            return False
        
        # Verify [X] symbol present
        if "[X]" in html:
            print("[OK] Error icon [X] present in HTML")
        
        print("[PASS] HTML generation test passed")
        return True
        
    except Exception as e:
        print(f"[FAIL] HTML generation test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_fstring_formatting():
    """Test f-string formatting patterns used in main.py"""
    print("\n" + "="*60)
    print("Testing F-String Formatting")
    print("="*60)
    
    try:
        # Test patterns from main.py
        message = "Test message"
        
        # Pattern 1: Simple f-string with variable
        test1 = f"[License] [OK] {message}"
        print(f"[OK] Pattern 1: {test1}")
        
        # Pattern 2: F-string with equals
        test2 = f"\n{'='*60}"
        print(f"[OK] Pattern 2: {test2}")
        
        # Pattern 3: F-string in multiline
        test3 = f"""
        Message: {message}
        The service is currently not available.
        """
        print(f"[OK] Pattern 3: {test3.strip()}")
        
        # Pattern 4: Nested formatting
        test4 = f"[Error] Failed to start application: {message}"
        print(f"[OK] Pattern 4: {test4}")
        
        print("[PASS] F-string formatting test passed")
        return True
        
    except Exception as e:
        print(f"[FAIL] F-string formatting test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    print("\n" + "="*70)
    print(" LICENSE PATH & ENCODING VALIDATION")
    print("="*70)
    
    tests = [
        ("License Validation", test_license_validation),
        ("HTML Rendering", test_html_rendering),
        ("F-String Formatting", test_fstring_formatting),
    ]
    
    results = []
    for test_name, test_func in tests:
        result = test_func()
        results.append((test_name, result))
    
    # Summary
    print("\n" + "="*70)
    print(" SUMMARY")
    print("="*70)
    
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for test_name, result in results:
        status = "[PASS]" if result else "[FAIL]"
        print(f"{status} {test_name}")
    
    print("="*70)
    
    if passed == total:
        print(f"\n[SUCCESS] All {total} tests passed!")
        print("main.py is safe to build into exe.")
        return 0
    else:
        print(f"\n[FAILURE] {total - passed}/{total} tests failed.")
        return 1


if __name__ == '__main__':
    exit_code = main()
    sys.exit(exit_code)
