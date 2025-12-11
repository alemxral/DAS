"""
Safe file handling utilities with automatic fallback for locked files
"""
import shutil
import time
import tempfile
import logging
from pathlib import Path
from typing import Optional
from contextlib import contextmanager
import openpyxl

logger = logging.getLogger(__name__)


class FileLockError(Exception):
    """Raised when a file cannot be accessed due to being locked"""
    pass


def create_safe_copy(file_path: str, max_retries: int = 3) -> Optional[str]:
    """
    Create a temporary copy of a file, with retry logic for locked files.
    
    Args:
        file_path: Path to the original file
        max_retries: Number of retry attempts
        
    Returns:
        Path to temporary copy, or None if failed
    """
    file_path = Path(file_path).absolute()
    
    for attempt in range(max_retries):
        try:
            # Create temp file with same extension
            suffix = file_path.suffix
            temp_fd, temp_path = tempfile.mkstemp(suffix=suffix, prefix='das_safe_')
            
            # Close the file descriptor and copy the file
            import os
            os.close(temp_fd)
            
            shutil.copy2(str(file_path), temp_path)
            logger.info(f"Created safe copy of {file_path.name} at {temp_path}")
            return temp_path
            
        except (PermissionError, OSError) as e:
            if attempt < max_retries - 1:
                wait_time = 0.2 * (attempt + 1)
                logger.warning(f"File copy attempt {attempt + 1} failed, retrying in {wait_time}s: {e}")
                time.sleep(wait_time)
            else:
                logger.error(f"Failed to create safe copy of {file_path} after {max_retries} attempts: {e}")
                return None
    
    return None


@contextmanager
def open_workbook_safe(file_path: str, data_only: bool = False, read_only: bool = False):
    """
    Safely open an Excel workbook with automatic fallback to temp copy if file is locked.
    
    Usage:
        with open_workbook_safe('data.xlsx') as wb:
            # Use workbook
            pass
    
    Args:
        file_path: Path to Excel file
        data_only: Load only values (no formulas)
        read_only: Open in read-only mode
        
    Yields:
        openpyxl.Workbook object
        
    Raises:
        FileLockError: If file cannot be opened after all retries
    """
    file_path = Path(file_path).absolute()
    wb = None
    temp_copy = None
    using_copy = False
    
    try:
        # Try opening the original file first
        try:
            wb = openpyxl.load_workbook(str(file_path), data_only=data_only, read_only=read_only)
            logger.debug(f"Opened workbook directly: {file_path.name}")
            
        except (PermissionError, OSError) as e:
            # File is locked, try working with a copy
            logger.warning(f"File locked, creating safe copy: {file_path.name} - {e}")
            temp_copy = create_safe_copy(str(file_path))
            
            if temp_copy is None:
                raise FileLockError(f"Cannot access file (locked) and failed to create copy: {file_path}")
            
            wb = openpyxl.load_workbook(temp_copy, data_only=data_only, read_only=read_only)
            using_copy = True
            logger.info(f"Using temporary copy for processing: {file_path.name}")
        
        yield wb
        
    finally:
        # Always close the workbook
        if wb is not None:
            try:
                wb.close()
                logger.debug(f"Closed workbook: {file_path.name}")
            except Exception as e:
                logger.error(f"Error closing workbook {file_path.name}: {e}")
        
        # Clean up temporary copy if used
        if temp_copy and using_copy:
            try:
                Path(temp_copy).unlink(missing_ok=True)
                logger.debug(f"Cleaned up temporary copy: {temp_copy}")
            except Exception as e:
                logger.warning(f"Could not delete temp copy {temp_copy}: {e}")


def safe_file_operation(file_path: str, operation_func, *args, max_retries: int = 3, **kwargs):
    """
    Execute a file operation with automatic retry and temp copy fallback.
    
    Args:
        file_path: Path to the file
        operation_func: Function to execute (receives file_path as first arg)
        *args: Additional positional arguments for operation_func
        max_retries: Number of retry attempts
        **kwargs: Additional keyword arguments for operation_func
        
    Returns:
        Result from operation_func
        
    Raises:
        FileLockError: If operation fails after all retries
    """
    file_path = Path(file_path).absolute()
    
    for attempt in range(max_retries):
        try:
            # Try with original file
            return operation_func(str(file_path), *args, **kwargs)
            
        except (PermissionError, OSError) as e:
            if attempt < max_retries - 1:
                wait_time = 0.2 * (attempt + 1)
                logger.warning(f"Operation attempt {attempt + 1} failed, retrying in {wait_time}s: {e}")
                time.sleep(wait_time)
            else:
                # Last attempt - try with temp copy
                logger.warning(f"All direct attempts failed, trying with temp copy: {file_path.name}")
                temp_copy = create_safe_copy(str(file_path))
                
                if temp_copy is None:
                    raise FileLockError(f"Cannot access file and failed to create copy: {file_path}")
                
                try:
                    result = operation_func(temp_copy, *args, **kwargs)
                    return result
                finally:
                    # Clean up temp copy
                    try:
                        Path(temp_copy).unlink(missing_ok=True)
                    except:
                        pass
    
    raise FileLockError(f"Operation failed after {max_retries} attempts: {file_path}")
