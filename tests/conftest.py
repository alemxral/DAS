"""
Pytest configuration and shared fixtures for test suite.
"""
import os
import sys
import shutil
import pytest
from pathlib import Path
from datetime import datetime

# Add project root to path
PROJECT_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from services.job_manager import JobManager
from services.template_processor import TemplateProcessor
from services.format_converter import FormatConverter
from services.document_parser import DocumentParser


@pytest.fixture(scope="session")
def project_root():
    """Get project root directory."""
    return PROJECT_ROOT


@pytest.fixture(scope="session")
def test_dir():
    """Get test directory."""
    return PROJECT_ROOT / "tests"


@pytest.fixture(scope="session")
def fixtures_dir():
    """Get fixtures directory."""
    return PROJECT_ROOT / "tests" / "fixtures"


@pytest.fixture(scope="function")
def output_dir():
    """Create and clean output directory for each test."""
    output_path = PROJECT_ROOT / "tests" / "output"
    output_path.mkdir(parents=True, exist_ok=True)
    
    # Clean output directory before test
    for item in output_path.iterdir():
        if item.is_file():
            item.unlink()
        elif item.is_dir():
            shutil.rmtree(item)
    
    yield output_path
    
    # Optional: Clean after test (commented out to keep test artifacts)
    # shutil.rmtree(output_path, ignore_errors=True)


@pytest.fixture(scope="function")
def temp_jobs_dir(tmp_path):
    """Create temporary jobs directory for testing."""
    jobs_dir = tmp_path / "test_jobs"
    jobs_dir.mkdir(parents=True, exist_ok=True)
    return jobs_dir


@pytest.fixture(scope="function")
def temp_storage_dir(tmp_path):
    """Create temporary storage directory for testing."""
    storage_dir = tmp_path / "test_storage"
    storage_dir.mkdir(parents=True, exist_ok=True)
    return storage_dir


@pytest.fixture(scope="function")
def job_manager(temp_jobs_dir, temp_storage_dir):
    """Create JobManager instance for testing."""
    return JobManager(
        jobs_dir=str(temp_jobs_dir),
        storage_dir=str(temp_storage_dir)
    )


@pytest.fixture(scope="function")
def template_processor():
    """Create TemplateProcessor instance for testing."""
    return TemplateProcessor()


@pytest.fixture(scope="function")
def format_converter():
    """Create FormatConverter instance for testing."""
    return FormatConverter()


@pytest.fixture(scope="function")
def document_parser():
    """Create DocumentParser instance for testing."""
    return DocumentParser()


@pytest.fixture(scope="session")
def sample_data():
    """Sample data for template processing."""
    return [
        {
            'name': 'John Doe',
            'email': 'john.doe@example.com',
            'phone': '555-1234',
            'company': 'Acme Corporation',
            'position': 'Senior Developer',
            'filename': 'john_doe_document',
            'tabname': 'John'
        },
        {
            'name': 'Jane Smith',
            'email': 'jane.smith@example.com',
            'phone': '555-5678',
            'company': 'Tech Solutions Inc',
            'position': 'Project Manager',
            'filename': 'jane_smith_document',
            'tabname': 'Jane'
        },
        {
            'name': 'Bob Johnson',
            'email': 'bob.j@example.com',
            'phone': '555-9012',
            'company': 'Global Enterprises',
            'position': 'Chief Technology Officer',
            'filename': 'bob_johnson_document',
            'tabname': 'Bob'
        }
    ]


@pytest.fixture(scope="session")
def excel_auto_adjust_options():
    """Sample Excel auto-adjust options."""
    return {
        'auto_adjust_height': True,
        'auto_adjust_width': True,
        'adjust_range': None
    }


@pytest.fixture(scope="session")
def excel_print_settings():
    """Sample Excel print settings for PDF conversion."""
    return {
        'page_range': {'from': 1, 'to': 0},
        'orientation': 'portrait',
        'paper_size': 'a4',
        'margins': {
            'left': 0.75,
            'right': 0.75,
            'top': 1.0,
            'bottom': 1.0
        },
        'scaling': {
            'type': 'percent',
            'value': 100,
            'width': None,
            'height': None
        },
        'center_horizontally': False,
        'center_vertically': False,
        'ignore_print_areas': False
    }


def pytest_configure(config):
    """Configure pytest with custom markers."""
    config.addinivalue_line(
        "markers", "slow: marks tests as slow (deselect with '-m \"not slow\"')"
    )
    config.addinivalue_line(
        "markers", "integration: marks tests as integration tests"
    )
    config.addinivalue_line(
        "markers", "performance: marks tests as performance tests"
    )
    config.addinivalue_line(
        "markers", "requires_libreoffice: marks tests requiring LibreOffice"
    )


def pytest_collection_modifyitems(config, items):
    """Modify test items during collection."""
    # Add markers automatically based on test names
    for item in items:
        if "slow" in item.nodeid.lower():
            item.add_marker(pytest.mark.slow)
        if "integration" in item.nodeid.lower():
            item.add_marker(pytest.mark.integration)
        if "performance" in item.nodeid.lower():
            item.add_marker(pytest.mark.performance)
        if "libreoffice" in item.nodeid.lower() or "pdf" in item.nodeid.lower():
            item.add_marker(pytest.mark.requires_libreoffice)
