"""
Job Model
Represents a document generation job with all its metadata and state.
"""
import json
import uuid
from datetime import datetime
from typing import List, Dict, Optional
from enum import Enum


class JobStatus(Enum):
    """Job status enumeration."""
    PENDING = "pending"
    PROCESSING = "processing"
    COMPLETED = "completed"
    FAILED = "failed"
    CANCELLED = "cancelled"


class Job:
    """Represents a document generation job."""
    
    def __init__(
        self,
        template_path: Optional[str] = None,
        data_path: Optional[str] = None,
        output_formats: Optional[List[str]] = None,
        job_id: Optional[str] = None
    ):
        """
        Initialize a Job.
        
        Args:
            template_path: Path to template file
            data_path: Path to data file
            output_formats: List of desired output formats
            job_id: Optional job ID (generated if not provided)
        """
        self.id = job_id or str(uuid.uuid4())
        self.template_path = template_path
        self.data_path = data_path
        self.output_formats = output_formats or []
        self.status = JobStatus.PENDING
        self.created_at = datetime.now()
        self.updated_at = datetime.now()
        self.started_at: Optional[datetime] = None
        self.completed_at: Optional[datetime] = None
        
        # File tracking
        self.template_file_id: Optional[str] = None
        self.data_file_id: Optional[str] = None
        self.local_template_path: Optional[str] = None
        self.local_data_path: Optional[str] = None
        
        # Processing details
        self.total_records: int = 0
        self.processed_records: int = 0
        self.failed_records: int = 0
        self.error_message: Optional[str] = None
        self.output_files: List[str] = []
        self.zip_file_path: Optional[str] = None
        
        # Metadata
        self.metadata: Dict = {}
        
        # Excel printing settings
        self.excel_print_settings: Optional[Dict] = None
        
        # Custom output directory
        self.output_directory: Optional[str] = None
    
    def to_dict(self) -> Dict:
        """
        Convert Job to dictionary.
        
        Returns:
            Dictionary representation of the job
        """
        return {
            'id': self.id,
            'template_path': self.template_path,
            'data_path': self.data_path,
            'output_formats': self.output_formats,
            'status': self.status.value,
            'created_at': self.created_at.isoformat(),
            'updated_at': self.updated_at.isoformat(),
            'started_at': self.started_at.isoformat() if self.started_at else None,
            'completed_at': self.completed_at.isoformat() if self.completed_at else None,
            'template_file_id': self.template_file_id,
            'data_file_id': self.data_file_id,
            'local_template_path': self.local_template_path,
            'local_data_path': self.local_data_path,
            'total_records': self.total_records,
            'processed_records': self.processed_records,
            'failed_records': self.failed_records,
            'error_message': self.error_message,
            'output_files': self.output_files,
            'zip_file_path': self.zip_file_path,
            'metadata': self.metadata,
            'excel_print_settings': self.excel_print_settings,
            'output_directory': self.output_directory
        }
    
    @classmethod
    def from_dict(cls, data: Dict) -> 'Job':
        """
        Create Job from dictionary.
        
        Args:
            data: Dictionary containing job data
            
        Returns:
            Job instance
        """
        job = cls(
            template_path=data.get('template_path'),
            data_path=data.get('data_path'),
            output_formats=data.get('output_formats', []),
            job_id=data.get('id')
        )
        
        # Set status
        status_value = data.get('status', 'pending')
        job.status = JobStatus(status_value)
        
        # Set timestamps
        if data.get('created_at'):
            job.created_at = datetime.fromisoformat(data['created_at'])
        if data.get('updated_at'):
            job.updated_at = datetime.fromisoformat(data['updated_at'])
        if data.get('started_at'):
            job.started_at = datetime.fromisoformat(data['started_at'])
        if data.get('completed_at'):
            job.completed_at = datetime.fromisoformat(data['completed_at'])
        
        # Set file tracking
        job.template_file_id = data.get('template_file_id')
        job.data_file_id = data.get('data_file_id')
        job.local_template_path = data.get('local_template_path')
        job.local_data_path = data.get('local_data_path')
        
        # Set processing details
        job.total_records = data.get('total_records', 0)
        job.processed_records = data.get('processed_records', 0)
        job.failed_records = data.get('failed_records', 0)
        job.error_message = data.get('error_message')
        job.output_files = data.get('output_files', [])
        job.zip_file_path = data.get('zip_file_path')
        job.metadata = data.get('metadata', {})
        job.excel_print_settings = data.get('excel_print_settings')
        job.output_directory = data.get('output_directory')
        
        return job
    
    def to_json(self) -> str:
        """
        Convert Job to JSON string.
        
        Returns:
            JSON string representation
        """
        return json.dumps(self.to_dict(), indent=2, ensure_ascii=False)
    
    @classmethod
    def from_json(cls, json_str: str) -> 'Job':
        """
        Create Job from JSON string.
        
        Args:
            json_str: JSON string containing job data
            
        Returns:
            Job instance
        """
        data = json.loads(json_str)
        return cls.from_dict(data)
    
    def update_status(self, status: JobStatus, error_message: Optional[str] = None):
        """
        Update job status.
        
        Args:
            status: New status
            error_message: Optional error message for failed jobs
        """
        self.status = status
        self.updated_at = datetime.now()
        
        if status == JobStatus.PROCESSING and not self.started_at:
            self.started_at = datetime.now()
        elif status in [JobStatus.COMPLETED, JobStatus.FAILED, JobStatus.CANCELLED]:
            self.completed_at = datetime.now()
        
        if error_message:
            self.error_message = error_message
    
    def get_progress_percentage(self) -> float:
        """
        Calculate job progress percentage.
        
        Returns:
            Progress as percentage (0-100)
        """
        if self.total_records == 0:
            return 0.0
        return (self.processed_records / self.total_records) * 100
    
    def get_duration(self) -> Optional[float]:
        """
        Get job duration in seconds.
        
        Returns:
            Duration in seconds or None if not completed
        """
        if self.started_at and self.completed_at:
            return (self.completed_at - self.started_at).total_seconds()
        elif self.started_at:
            return (datetime.now() - self.started_at).total_seconds()
        return None
    
    def add_output_file(self, file_path: str):
        """Add an output file to the job."""
        if file_path not in self.output_files:
            self.output_files.append(file_path)
            self.updated_at = datetime.now()
    
    def set_zip_file(self, zip_path: str):
        """Set the ZIP file path for the job."""
        self.zip_file_path = zip_path
        self.updated_at = datetime.now()
    
    def increment_processed(self):
        """Increment the processed records counter."""
        self.processed_records += 1
        self.updated_at = datetime.now()
    
    def increment_failed(self):
        """Increment the failed records counter."""
        self.failed_records += 1
        self.updated_at = datetime.now()
    
    def get_summary(self) -> Dict:
        """
        Get a summary of the job for display.
        
        Returns:
            Dictionary with job summary
        """
        return {
            'id': self.id,
            'status': self.status.value,
            'progress': f"{self.get_progress_percentage():.1f}%",
            'records': f"{self.processed_records}/{self.total_records}",
            'duration': f"{self.get_duration():.1f}s" if self.get_duration() else "N/A",
            'output_formats': ', '.join(self.output_formats),
            'created_at': self.created_at.strftime('%Y-%m-%d %H:%M:%S')
        }
