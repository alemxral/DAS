"""
File Tracker Service
Tracks files using SHA-256 hashing to detect changes and maintain local copies.
"""
import os
import hashlib
import shutil
import json
from pathlib import Path
from typing import Optional, Dict
from datetime import datetime


class FileTracker:
    """Manages file tracking and synchronization using SHA-256 hashing."""
    
    def __init__(self, storage_dir: str):
        """
        Initialize FileTracker.
        
        Args:
            storage_dir: Directory to store tracked file copies
        """
        self.storage_dir = Path(storage_dir)
        self.storage_dir.mkdir(parents=True, exist_ok=True)
        self.metadata_file = self.storage_dir / "file_metadata.json"
        self.metadata = self._load_metadata()
    
    def _load_metadata(self) -> Dict:
        """Load metadata from JSON file."""
        if self.metadata_file.exists():
            with open(self.metadata_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {}
    
    def _save_metadata(self):
        """Save metadata to JSON file."""
        with open(self.metadata_file, 'w', encoding='utf-8') as f:
            json.dump(self.metadata, f, indent=2, ensure_ascii=False)
    
    def _calculate_sha256(self, file_path: str) -> str:
        """
        Calculate SHA-256 hash of a file.
        
        Args:
            file_path: Path to the file
            
        Returns:
            SHA-256 hash as hexadecimal string
        """
        sha256_hash = hashlib.sha256()
        with open(file_path, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    
    def get_file_id(self, original_path: str) -> str:
        """
        Generate a unique ID for a file based on its path.
        
        Args:
            original_path: Original file path
            
        Returns:
            Unique file ID
        """
        return hashlib.md5(original_path.encode()).hexdigest()
    
    def track_file(self, original_path: str, force_update: bool = False) -> Dict:
        """
        Track a file and create/update local copy if needed.
        
        Args:
            original_path: Path to the original file
            force_update: Force update even if hash matches
            
        Returns:
            Dictionary with file information including local path
            
        Raises:
            FileNotFoundError: If original file doesn't exist
        """
        if not os.path.exists(original_path):
            raise FileNotFoundError(f"File not found: {original_path}")
        
        file_id = self.get_file_id(original_path)
        current_hash = self._calculate_sha256(original_path)
        file_ext = Path(original_path).suffix
        local_filename = f"{file_id}{file_ext}"
        local_path = self.storage_dir / local_filename
        
        # Check if we need to update the local copy
        needs_update = (
            force_update or
            file_id not in self.metadata or
            self.metadata[file_id]['sha256'] != current_hash or
            not local_path.exists()
        )
        
        if needs_update:
            # Copy file to storage
            shutil.copy2(original_path, local_path)
            
            # Update metadata
            self.metadata[file_id] = {
                'original_path': original_path,
                'local_path': str(local_path),
                'sha256': current_hash,
                'last_updated': datetime.now().isoformat(),
                'file_name': Path(original_path).name,
                'file_size': os.path.getsize(original_path)
            }
            self._save_metadata()
        
        return {
            'file_id': file_id,
            'local_path': str(local_path),
            'original_path': original_path,
            'updated': needs_update,
            'sha256': current_hash
        }
    
    def get_local_path(self, file_id: str) -> Optional[str]:
        """
        Get local path for a tracked file.
        
        Args:
            file_id: File ID
            
        Returns:
            Local path if file is tracked, None otherwise
        """
        if file_id in self.metadata:
            return self.metadata[file_id]['local_path']
        return None
    
    def is_file_changed(self, original_path: str) -> bool:
        """
        Check if a file has changed since last tracking.
        
        Args:
            original_path: Path to the original file
            
        Returns:
            True if file has changed, False otherwise
        """
        if not os.path.exists(original_path):
            return True
        
        file_id = self.get_file_id(original_path)
        if file_id not in self.metadata:
            return True
        
        current_hash = self._calculate_sha256(original_path)
        return current_hash != self.metadata[file_id]['sha256']
    
    def get_file_info(self, file_id: str) -> Optional[Dict]:
        """
        Get information about a tracked file.
        
        Args:
            file_id: File ID
            
        Returns:
            File metadata dictionary or None
        """
        return self.metadata.get(file_id)
    
    def cleanup_orphaned_files(self):
        """Remove local copies of files that no longer exist at their original path."""
        orphaned_ids = []
        for file_id, info in self.metadata.items():
            if not os.path.exists(info['original_path']):
                local_path = Path(info['local_path'])
                if local_path.exists():
                    local_path.unlink()
                orphaned_ids.append(file_id)
        
        for file_id in orphaned_ids:
            del self.metadata[file_id]
        
        if orphaned_ids:
            self._save_metadata()
        
        return len(orphaned_ids)
