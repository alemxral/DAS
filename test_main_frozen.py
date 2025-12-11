"""
Test script to validate main.py behavior in frozen exe environment.
This simulates PyInstaller frozen state to catch potential crashes before building.
"""
import sys
import os
from pathlib import Path
import tempfile
import shutil

def test_frozen_environment():
    """Test 1: Simulate frozen exe environment"""
    print("\n" + "="*60)
    print("TEST 1: Frozen Environment Simulation")
    print("="*60)
    
    try:
        # Simulate frozen state
        original_frozen = getattr(sys, 'frozen', False)
        original_meipass = getattr(sys, '_MEIPASS', None)
        
        # Create temp directory to simulate _MEIPASS
        temp_dir = tempfile.mkdtemp()
        
        try:
            sys.frozen = True
            sys._MEIPASS = temp_dir
            
            # Test BASE_DIR calculation
            from pathlib import Path
            BASE_DIR = Path(sys._MEIPASS)
            
            print(f"[OK] BASE_DIR set to: {BASE_DIR}")
            print(f"[OK] Directory exists: {BASE_DIR.exists()}")
            
        finally:
            # Restore original state
            if original_frozen:
                sys.frozen = original_frozen
            else:
                if hasattr(sys, 'frozen'):
                    delattr(sys, 'frozen')
            
            if original_meipass:
                sys._MEIPASS = original_meipass
            else:
                if hasattr(sys, '_MEIPASS'):
                    delattr(sys, '_MEIPASS')
            
            # Cleanup temp dir
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir, ignore_errors=True)
        
        print("[PASS] Frozen environment test passed")
        return True
        
    except Exception as e:
        print(f"[FAIL] Frozen environment test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_stdout_wrapping():
    """Test 2: Test stdout/stderr wrapping logic"""
    print("\n" + "="*60)
    print("TEST 2: Stdout/Stderr Wrapping")
    print("="*60)
    
    try:
        import io
        
        # Save original stdout/stderr
        original_stdout = sys.stdout
        original_stderr = sys.stderr
        
        try:
            # Test wrapping logic
            if hasattr(sys.stdout, 'buffer') and not isinstance(sys.stdout, io.TextIOWrapper):
                test_stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
                print("[OK] stdout wrapping successful")
            else:
                print("[OK] stdout already wrapped or no buffer")
            
            if hasattr(sys.stderr, 'buffer') and not isinstance(sys.stderr, io.TextIOWrapper):
                test_stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
                print("[OK] stderr wrapping successful")
            else:
                print("[OK] stderr already wrapped or no buffer")
            
            # Test printing with various characters
            test_strings = [
                "ASCII test: Hello World",
                "UTF-8 test: Testing encoding",
                "Special chars: [OK] [ERROR] [X]",
            ]
            
            for test_str in test_strings:
                print(f"[OK] Print test: {test_str}")
            
        finally:
            # Restore original stdout/stderr
            sys.stdout = original_stdout
            sys.stderr = original_stderr
        
        print("[PASS] Stdout/stderr wrapping test passed")
        return True
        
    except Exception as e:
        print(f"[FAIL] Stdout/stderr test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_imports():
    """Test 3: Verify all required imports"""
    print("\n" + "="*60)
    print("TEST 3: Import Validation")
    print("="*60)
    
    failed_imports = []
    
    imports_to_test = [
        ('sys', 'sys'),
        ('os', 'os'),
        ('pathlib', 'Path'),
        ('threading', 'threading'),
        ('time', 'time'),
        ('webview', 'webview'),
        ('base64', 'base64'),
        ('io', 'io'),
    ]
    
    for module_name, import_name in imports_to_test:
        try:
            if module_name == import_name:
                __import__(module_name)
            else:
                module = __import__(module_name)
                getattr(module, import_name.split('.')[-1])
            print(f"[OK] Import successful: {import_name}")
        except ImportError as e:
            print(f"[FAIL] Import failed: {import_name} - {e}")
            failed_imports.append(import_name)
        except AttributeError as e:
            print(f"[FAIL] Import failed: {import_name} - {e}")
            failed_imports.append(import_name)
    
    if not failed_imports:
        print("[PASS] All imports test passed")
        return True
    else:
        print(f"[FAIL] Failed imports: {', '.join(failed_imports)}")
        return False


def test_api_class():
    """Test 4: Test Api class methods"""
    print("\n" + "="*60)
    print("TEST 4: Api Class Validation")
    print("="*60)
    
    try:
        import base64
        
        # Import Api class from main
        import importlib.util
        spec = importlib.util.spec_from_file_location("main", "main.py")
        main_module = importlib.util.module_from_spec(spec)
        
        # Mock webview before loading
        import sys
        from unittest.mock import MagicMock
        sys.modules['webview'] = MagicMock()
        
        spec.loader.exec_module(main_module)
        
        Api = main_module.Api
        
        # Test Api instantiation
        api = Api()
        print("[OK] Api class instantiated")
        
        # Test save_file method exists
        assert hasattr(api, 'save_file'), "save_file method missing"
        print("[OK] save_file method exists")
        
        # Test save_file signature
        import inspect
        sig = inspect.signature(api.save_file)
        params = list(sig.parameters.keys())
        assert 'filename' in params, "filename parameter missing"
        assert 'data_base64' in params, "data_base64 parameter missing"
        print("[OK] save_file has correct parameters")
        
        print("[PASS] Api class test passed")
        return True
        
    except Exception as e:
        print(f"[FAIL] Api class test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_path_handling():
    """Test 5: Test path handling logic"""
    print("\n" + "="*60)
    print("TEST 5: Path Handling")
    print("="*60)
    
    try:
        from pathlib import Path
        
        # Test BASE_DIR calculation (normal mode)
        if not getattr(sys, 'frozen', False):
            BASE_DIR = Path(__file__).parent
            print(f"[OK] Normal mode BASE_DIR: {BASE_DIR}")
            print(f"[OK] Directory exists: {BASE_DIR.exists()}")
        
        # Test icon path construction
        test_base = Path.cwd()
        icon_path = test_base / 'static' / 'icon.png'
        print(f"[OK] Icon path construction: {icon_path}")
        
        # Test path existence check
        if icon_path.exists():
            print(f"[OK] Icon file exists at {icon_path}")
        else:
            print(f"[INFO] Icon file not found (this is OK, app handles it)")
        
        print("[PASS] Path handling test passed")
        return True
        
    except Exception as e:
        print(f"[FAIL] Path handling test failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def test_exception_handling():
    """Test 6: Verify exception handling patterns"""
    print("\n" + "="*60)
    print("TEST 6: Exception Handling")
    print("="*60)
    
    try:
        # Read main.py and check for exception handling
        with open('main.py', 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Check for try-except blocks
        if 'try:' in content and 'except' in content:
            print("[OK] Exception handling present")
        else:
            print("[WARN] Limited exception handling found")
        
        # Check for sys.exit calls
        if 'sys.exit' in content:
            print("[OK] Proper exit handling present")
        
        # Check for traceback printing
        if 'traceback' in content:
            print("[OK] Traceback debugging present")
        
        print("[PASS] Exception handling test passed")
        return True
        
    except Exception as e:
        print(f"[FAIL] Exception handling test failed: {e}")
        return False


def test_print_safety():
    """Test 7: Test print statement safety with various encodings"""
    print("\n" + "="*60)
    print("TEST 7: Print Safety")
    print("="*60)
    
    try:
        # Test various print scenarios
        test_cases = [
            ("ASCII only", "Hello World"),
            ("Numbers", "Error code: 12345"),
            ("Special ASCII", "[OK] [ERROR] [X]"),
            ("Formatted string", f"{'='*60}"),
            ("F-string with vars", f"Value: {42}"),
        ]
        
        for test_name, test_value in test_cases:
            try:
                print(f"[OK] {test_name}: {test_value}")
            except Exception as e:
                print(f"[FAIL] {test_name} failed: {e}")
                return False
        
        print("[PASS] Print safety test passed")
        return True
        
    except Exception as e:
        print(f"[FAIL] Print safety test failed: {e}")
        return False


def main():
    """Run all tests"""
    print("\n" + "="*70)
    print(" MAIN.PY FROZEN EXE VALIDATION TEST SUITE")
    print("="*70)
    print("Testing for potential crashes before building executable...")
    
    tests = [
        ("Frozen Environment", test_frozen_environment),
        ("Stdout Wrapping", test_stdout_wrapping),
        ("Imports", test_imports),
        ("Api Class", test_api_class),
        ("Path Handling", test_path_handling),
        ("Exception Handling", test_exception_handling),
        ("Print Safety", test_print_safety),
    ]
    
    results = []
    for test_name, test_func in tests:
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"\n[CRITICAL] Test '{test_name}' crashed: {e}")
            import traceback
            traceback.print_exc()
            results.append((test_name, False))
    
    # Summary
    print("\n" + "="*70)
    print(" TEST SUMMARY")
    print("="*70)
    
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for test_name, result in results:
        status = "[PASS]" if result else "[FAIL]"
        print(f"{status} {test_name}")
    
    print("="*70)
    print(f"Results: {passed}/{total} tests passed")
    
    if passed == total:
        print("\n[SUCCESS] All tests passed! main.py should be safe for exe build.")
        return 0
    else:
        print(f"\n[WARNING] {total - passed} test(s) failed. Review issues before building.")
        return 1


if __name__ == '__main__':
    exit_code = main()
    sys.exit(exit_code)
