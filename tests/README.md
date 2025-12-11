# Document Automation System - Test Suite

## Overview
Comprehensive testing module for the Document Automation System covering all conversion options, template processing, Excel features, and LibreOffice integration.

## Test Structure

```
tests/
├── __init__.py              # Test package initialization
├── conftest.py              # Pytest configuration and fixtures
├── test_suite.py            # Main test suite (template processing, conversions)
├── test_integration.py      # End-to-end integration tests
├── test_performance.py      # Performance and load tests
├── test_validators.py       # Output validation and integrity tests
├── fixtures/                # Test data and templates
├── output/                  # Test output directory
├── requirements.txt         # Additional test dependencies
└── README.md               # This file
```

## Test Categories

### 1. Basic Functionality Tests (`test_suite.py`)
- ✅ DOCX variable substitution
- ✅ XLSX variable substitution
- ✅ Excel auto-adjust height
- ✅ Excel auto-adjust width
- ✅ Excel auto-adjust specific range
- ✅ Template caching performance
- ✅ Format conversions (DOCX→PDF, XLSX→PDF)
- ✅ Job creation and management
- ✅ Edge cases and error handling

### 2. Integration Tests (`test_integration.py`)
- ✅ Single template, single output
- ✅ Single template, multiple outputs
- ✅ Multiple output formats
- ✅ Excel auto-adjust integration
- ✅ ZIP archive creation
- ✅ Custom output directories
- ✅ Multi-template workflows
- ✅ Error recovery and handling

### 3. Performance Tests (`test_performance.py`)
- ✅ Template caching speedup measurement
- ✅ Large dataset processing (100+ records)
- ✅ Memory usage tracking
- ✅ Conversion speed benchmarks
- ✅ Concurrent processing
- ✅ Multi-sheet processing
- ✅ Large file handling
- ✅ Cache memory efficiency
- ✅ Disk I/O efficiency

### 4. Validation Tests (`test_validators.py`)
- ✅ DOCX file validity
- ✅ XLSX file validity
- ✅ PDF file validity
- ✅ Variable substitution correctness
- ✅ Excel auto-adjust application
- ✅ ZIP archive integrity

## Installation

1. **Install test dependencies:**
   ```bash
   pip install -r tests/requirements.txt
   ```

2. **Or install individually:**
   ```bash
   pip install pytest pytest-html pytest-cov psutil
   ```

## Running Tests

### Quick Start
```bash
# Run all tests
python run_tests.py

# Or use batch file (Windows)
run_tests.bat
```

### Test Modes

**Fast Mode** (skip slow tests and LibreOffice-dependent tests):
```bash
python run_tests.py fast
```

**Integration Tests Only**:
```bash
python run_tests.py integration
```

**Performance Tests Only**:
```bash
python run_tests.py performance
```

**Without LibreOffice** (useful for CI/CD):
```bash
python run_tests.py no-libreoffice
```

### Direct Pytest Usage

**Run specific test file:**
```bash
pytest tests/test_suite.py -v
```

**Run specific test class:**
```bash
pytest tests/test_suite.py::TestTemplateProcessor -v
```

**Run specific test:**
```bash
pytest tests/test_suite.py::TestTemplateProcessor::test_docx_variable_substitution -v
```

**Run with markers:**
```bash
# Skip slow tests
pytest -m "not slow" tests/

# Only integration tests
pytest -m integration tests/

# Only performance tests
pytest -m performance tests/
```

**Generate coverage report:**
```bash
pytest --cov=services --cov=models --cov-report=html tests/
```

## Test Markers

Tests are marked with custom markers for selective execution:

- `@pytest.mark.slow` - Long-running tests
- `@pytest.mark.integration` - Integration tests
- `@pytest.mark.performance` - Performance benchmarks
- `@pytest.mark.requires_libreoffice` - Needs LibreOffice installed

## CI/CD Integration

### GitHub Actions Example
```yaml
name: Tests
on: [push, pull_request]
jobs:
  test:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v2
      - uses: actions/setup-python@v2
        with:
          python-version: '3.12'
      - run: pip install -r requirements.txt
      - run: pip install -r tests/requirements.txt
      - run: python run_tests.py no-libreoffice
      - uses: actions/upload-artifact@v2
        with:
          name: test-report
          path: tests/report_*.html
```

## Test Reports

After running tests, HTML reports are generated in the `tests/` directory with names like:
- `report_20251211_143022.html` (timestamped)

Open in browser to view detailed results including:
- Pass/fail status for each test
- Execution times
- Error tracebacks
- Test duration rankings

## Performance Benchmarks

Expected performance targets:

| Metric | Target | Test |
|--------|--------|------|
| Template caching speedup | > 1.5x | `test_template_caching_speedup` |
| Records per second | > 10 | `test_large_dataset_processing` |
| Memory overhead | < 200 MB | `test_memory_usage` |
| Conversion time | < 30s | `test_conversion_speed` |
| I/O operations/sec | > 5 | `test_disk_io_efficiency` |

## Adding New Tests

1. **Choose appropriate test file:**
   - Basic functionality → `test_suite.py`
   - End-to-end workflows → `test_integration.py`
   - Performance metrics → `test_performance.py`
   - Output validation → `test_validators.py`

2. **Create test class and methods:**
   ```python
   class TestNewFeature:
       def test_feature_basic(self, template_processor, output_dir):
           # Arrange
           template = create_template()
           
           # Act
           result = process_template(template)
           
           # Assert
           assert result is not None
   ```

3. **Use fixtures from conftest.py:**
   - `output_dir` - Clean output directory
   - `template_processor` - TemplateProcessor instance
   - `format_converter` - FormatConverter instance
   - `job_manager` - JobManager instance
   - `sample_data` - Test data dictionary

4. **Add appropriate markers:**
   ```python
   @pytest.mark.slow
   @pytest.mark.requires_libreoffice
   def test_complex_conversion(self):
       pass
   ```

## Troubleshooting

### Tests failing with "LibreOffice not found"
- Install LibreOffice or run with: `python run_tests.py no-libreoffice`

### ImportError for test dependencies
- Run: `pip install -r tests/requirements.txt`

### Slow test execution
- Run fast tests only: `python run_tests.py fast`
- Skip performance tests: `pytest -m "not performance" tests/`

### Permission errors in output directory
- Ensure `tests/output/` is writable
- Close any files opened from previous test runs

## Test Coverage

Current test coverage by module:

- `services/template_processor.py` - Template processing, variable substitution, Excel auto-adjust
- `services/format_converter.py` - Format conversions, LibreOffice integration
- `services/job_manager.py` - Job lifecycle, output generation, ZIP creation
- `services/document_parser.py` - Data file parsing
- `models/job.py` - Job model serialization

## Continuous Improvement

The test suite is designed to grow with the application. When adding new features:

1. ✅ Write tests first (TDD approach)
2. ✅ Ensure tests pass before committing
3. ✅ Add integration tests for complete workflows
4. ✅ Include performance tests if relevant
5. ✅ Update this README with new test descriptions

## Contact

For issues or questions about the test suite, refer to the main project documentation or create an issue in the repository.
