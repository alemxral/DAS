"""
Test Runner Script - Execute All Tests with Reporting
Runs the complete test suite and generates reports.
"""
import sys
import subprocess
from pathlib import Path
from datetime import datetime


def run_tests():
    """Run all tests with pytest and generate reports."""
    
    print("=" * 80)
    print("Document Automation System - Test Suite")
    print("=" * 80)
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    # Ensure we're in the project root
    project_root = Path(__file__).parent
    
    # Test categories
    test_suites = {
        'basic': {
            'name': 'Basic Functionality Tests',
            'path': 'tests/test_suite.py',
            'markers': None
        },
        'integration': {
            'name': 'Integration Tests',
            'path': 'tests/test_integration.py',
            'markers': 'integration'
        },
        'validators': {
            'name': 'Validation Tests',
            'path': 'tests/test_validators.py',
            'markers': None
        },
        'performance': {
            'name': 'Performance Tests',
            'path': 'tests/test_performance.py',
            'markers': 'performance'
        }
    }
    
    # Run options
    run_mode = 'all'  # Options: 'all', 'fast', 'integration', 'performance'
    
    if len(sys.argv) > 1:
        run_mode = sys.argv[1]
    
    print(f"Run Mode: {run_mode}\n")
    
    # Build pytest command
    base_cmd = [
        sys.executable,
        '-m', 'pytest',
        '-v',
        '--tb=short',
        '--color=yes',
        '--durations=10',
        f'--html=tests/report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.html',
        '--self-contained-html'
    ]
    
    # Select tests based on mode
    if run_mode == 'fast':
        print("Running FAST tests (excluding slow and LibreOffice-dependent)...\n")
        cmd = base_cmd + [
            '-m', 'not slow and not requires_libreoffice',
            'tests/'
        ]
    elif run_mode == 'integration':
        print("Running INTEGRATION tests only...\n")
        cmd = base_cmd + [
            '-m', 'integration',
            'tests/'
        ]
    elif run_mode == 'performance':
        print("Running PERFORMANCE tests only...\n")
        cmd = base_cmd + [
            '-m', 'performance',
            'tests/'
        ]
    elif run_mode == 'no-libreoffice':
        print("Running tests WITHOUT LibreOffice dependency...\n")
        cmd = base_cmd + [
            '-m', 'not requires_libreoffice',
            'tests/'
        ]
    else:  # all
        print("Running ALL tests...\n")
        cmd = base_cmd + ['tests/']
    
    # Run tests
    try:
        result = subprocess.run(cmd, cwd=project_root)
        exit_code = result.returncode
    except Exception as e:
        print(f"\n[ERROR] Error running tests: {e}")
        exit_code = 1
    
    # Summary
    print("\n" + "=" * 80)
    print(f"Completed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    if exit_code == 0:
        print("[OK] All tests PASSED")
    else:
        print(f"[ERROR] Tests FAILED (exit code: {exit_code})")
    
    print("=" * 80)
    
    return exit_code


if __name__ == '__main__':
    print("""
Usage:
    python run_tests.py [mode]
    
Modes:
    all             - Run all tests (default)
    fast            - Run fast tests only (skip slow and LibreOffice tests)
    integration     - Run integration tests only
    performance     - Run performance tests only
    no-libreoffice  - Run tests without LibreOffice dependency
    """)
    
    exit_code = run_tests()
    sys.exit(exit_code)
