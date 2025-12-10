"""
Utility Functions
Common helper functions used across the application.
"""
import os
import hashlib
from pathlib import Path
from typing import Optional


def get_file_extension(file_path: str) -> str:
    """
    Get file extension from path.
    
    Args:
        file_path: Path to file
        
    Returns:
        File extension including dot (e.g., '.docx')
    """
    return Path(file_path).suffix.lower()


def get_file_name(file_path: str) -> str:
    """
    Get filename from path without extension.
    
    Args:
        file_path: Path to file
        
    Returns:
        Filename without extension
    """
    return Path(file_path).stem


def ensure_dir(directory: str) -> str:
    """
    Ensure directory exists, create if not.
    
    Args:
        directory: Directory path
        
    Returns:
        Absolute path to directory
    """
    path = Path(directory)
    path.mkdir(parents=True, exist_ok=True)
    return str(path.absolute())


def safe_filename(filename: str) -> str:
    """
    Create safe filename by removing/replacing invalid characters.
    
    Args:
        filename: Original filename
        
    Returns:
        Safe filename
    """
    # Remove or replace invalid characters
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename


def calculate_file_hash(file_path: str, algorithm: str = 'sha256') -> str:
    """
    Calculate hash of a file.
    
    Args:
        file_path: Path to file
        algorithm: Hash algorithm ('sha256', 'md5', etc.)
        
    Returns:
        Hash as hexadecimal string
    """
    hash_obj = hashlib.new(algorithm)
    with open(file_path, 'rb') as f:
        for chunk in iter(lambda: f.read(4096), b''):
            hash_obj.update(chunk)
    return hash_obj.hexdigest()


def format_bytes(bytes_size: int) -> str:
    """
    Format bytes to human-readable string.
    
    Args:
        bytes_size: Size in bytes
        
    Returns:
        Formatted string (e.g., '1.5 MB')
    """
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if bytes_size < 1024.0:
            return f"{bytes_size:.2f} {unit}"
        bytes_size /= 1024.0
    return f"{bytes_size:.2f} PB"


def is_valid_path(path: str) -> bool:
    """
    Check if path is valid and accessible.
    
    Args:
        path: File or directory path
        
    Returns:
        True if path exists and is accessible
    """
    try:
        return os.path.exists(path) and os.access(path, os.R_OK)
    except Exception:
        return False


def get_file_info(file_path: str) -> Optional[dict]:
    """
    Get detailed file information.
    
    Args:
        file_path: Path to file
        
    Returns:
        Dictionary with file info or None
    """
    if not os.path.exists(file_path):
        return None
    
    stat = os.stat(file_path)
    return {
        'path': file_path,
        'name': os.path.basename(file_path),
        'size': stat.st_size,
        'size_formatted': format_bytes(stat.st_size),
        'extension': get_file_extension(file_path),
        'modified': stat.st_mtime,
        'created': stat.st_ctime
    }


def cleanup_old_files(directory: str, days: int = 7) -> int:
    """
    Remove files older than specified days.
    
    Args:
        directory: Directory to clean
        days: Age threshold in days
        
    Returns:
        Number of files removed
    """
    import time
    
    if not os.path.exists(directory):
        return 0
    
    current_time = time.time()
    threshold = days * 24 * 60 * 60
    removed_count = 0
    
    for root, dirs, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            if current_time - os.path.getmtime(file_path) > threshold:
                try:
                    os.remove(file_path)
                    removed_count += 1
                except Exception as e:
                    print(f"Error removing {file_path}: {str(e)}")
    
    return removed_count
