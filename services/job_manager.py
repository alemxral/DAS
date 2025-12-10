"""
Job Manager Service
Manages job lifecycle, file operations, and output generation.
"""
import os
import json
import zipfile
import shutil
from pathlib import Path
from typing import List, Dict, Optional
from datetime import datetime

from models.job import Job, JobStatus
from services.file_tracker import FileTracker
from services.document_parser import DocumentParser
from services.template_processor import TemplateProcessor
from services.format_converter import FormatConverter


class JobManager:
    """Manages document generation jobs."""
    
    def __init__(self, jobs_dir: str, storage_dir: str):
        """
        Initialize JobManager.
        
        Args:
            jobs_dir: Directory to store job data
            storage_dir: Directory for file tracking
        """
        self.jobs_dir = Path(jobs_dir)
        self.storage_dir = Path(storage_dir)
        self.jobs_dir.mkdir(parents=True, exist_ok=True)
        self.storage_dir.mkdir(parents=True, exist_ok=True)
        
        self.file_tracker = FileTracker(str(self.storage_dir))
        self.document_parser = DocumentParser()
        self.template_processor = TemplateProcessor()
        self.format_converter = FormatConverter()
        
        self.jobs: Dict[str, Job] = {}
        self._load_all_jobs()
    
    def _load_all_jobs(self):
        """Load all jobs from disk."""
        if not self.jobs_dir.exists():
            return
        
        for job_dir in self.jobs_dir.iterdir():
            if job_dir.is_dir():
                metadata_file = job_dir / "metadata.json"
                if metadata_file.exists():
                    try:
                        with open(metadata_file, 'r', encoding='utf-8') as f:
                            job_data = json.load(f)
                            job = Job.from_dict(job_data)
                            self.jobs[job.id] = job
                    except Exception as e:
                        print(f"Error loading job {job_dir.name}: {str(e)}")
    
    def get_job_dir(self, job_id: str) -> Path:
        """Get the directory for a specific job."""
        return self.jobs_dir / job_id
    
    def save_job_metadata(self, job: Job):
        """Save job metadata to disk."""
        job_dir = self.get_job_dir(job.id)
        job_dir.mkdir(parents=True, exist_ok=True)
        
        metadata_file = job_dir / "metadata.json"
        with open(metadata_file, 'w', encoding='utf-8') as f:
            json.dump(job.to_dict(), f, indent=2, ensure_ascii=False)
    
    def create_job(
        self,
        template_path: str,
        data_path: str,
        output_formats: List[str],
        excel_print_settings: Optional[Dict] = None,
        output_directory: Optional[str] = None,
        filename_variable: str = '##filename##'
    ) -> Job:
        """
        Create a new job.
        
        Args:
            template_path: Path to template file
            data_path: Path to data file
            output_formats: List of desired output formats
            excel_print_settings: Optional Excel print settings for PDF conversion
            output_directory: Optional custom output directory
            filename_variable: Variable to use for output filenames (default: ##filename##)
            
        Returns:
            Created Job instance
        """
        job = Job(
            template_path=template_path,
            data_path=data_path,
            output_formats=output_formats
        )
        
        # Add Excel print settings if provided
        if excel_print_settings:
            job.excel_print_settings = excel_print_settings
        
        # Store custom output directory
        if output_directory:
            job.output_directory = output_directory
        
        # Store filename variable
        job.metadata['filename_variable'] = filename_variable
        
        # Create job directory
        job_dir = self.get_job_dir(job.id)
        job_dir.mkdir(parents=True, exist_ok=True)
        
        # Track and copy files
        try:
            # Track template file
            template_info = self.file_tracker.track_file(template_path)
            job.template_file_id = template_info['file_id']
            job.local_template_path = template_info['local_path']
            
            # Track data file
            data_info = self.file_tracker.track_file(data_path)
            job.data_file_id = data_info['file_id']
            job.local_data_path = data_info['local_path']
            
            # Copy files to job directory for processing
            job_template_path = job_dir / f"template{Path(template_path).suffix}"
            job_data_path = job_dir / f"data{Path(data_path).suffix}"
            
            shutil.copy2(job.local_template_path, job_template_path)
            shutil.copy2(job.local_data_path, job_data_path)
            
            job.metadata['job_template_path'] = str(job_template_path)
            job.metadata['job_data_path'] = str(job_data_path)
            
        except Exception as e:
            job.update_status(JobStatus.FAILED, f"Error tracking files: {str(e)}")
        
        # Save job
        self.jobs[job.id] = job
        self.save_job_metadata(job)
        
        return job
    
    def get_job(self, job_id: str) -> Optional[Job]:
        """Get a job by ID."""
        return self.jobs.get(job_id)
    
    def get_all_jobs(self) -> List[Job]:
        """Get all jobs."""
        return list(self.jobs.values())
    
    def get_jobs_by_status(self, status: JobStatus) -> List[Job]:
        """Get all jobs with a specific status."""
        return [job for job in self.jobs.values() if job.status == status]
    
    def delete_job(self, job_id: str) -> bool:
        """
        Delete a job and its associated files.
        
        Args:
            job_id: Job ID
            
        Returns:
            True if deleted successfully
        """
        if job_id not in self.jobs:
            return False
        
        # Delete job directory
        job_dir = self.get_job_dir(job_id)
        if job_dir.exists():
            shutil.rmtree(job_dir)
        
        # Remove from memory
        del self.jobs[job_id]
        
        return True
    
    def process_job(self, job_id: str) -> Job:
        """
        Process a job to generate documents.
        
        Args:
            job_id: Job ID
            
        Returns:
            Updated Job instance
        """
        job = self.get_job(job_id)
        if not job:
            raise ValueError(f"Job not found: {job_id}")
        
        if job.status != JobStatus.PENDING:
            raise ValueError(f"Job {job_id} cannot be processed (status: {job.status.value})")
        
        job.update_status(JobStatus.PROCESSING)
        self.save_job_metadata(job)
        
        try:
            # Parse data file
            data_result = self.document_parser.parse_excel_data(
                job.metadata['job_data_path']
            )
            job.total_records = data_result['total_rows']
            
            # Create output directory
            output_dir = self.get_job_dir(job.id) / "outputs"
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # Process each data row
            for idx, row_data in enumerate(data_result['data'], start=1):
                try:
                    # Determine output filename
                    # Check for ##filename## variable or custom filename variable in job metadata
                    filename_var = job.metadata.get('filename_variable', '##filename##')
                    # Remove the ## markers to get the key
                    filename_key = filename_var.replace('##', '')
                    
                    # Try to get custom filename from row data
                    custom_filename = None
                    if filename_key in row_data and row_data[filename_key]:
                        custom_filename = str(row_data[filename_key]).strip()
                        # Sanitize filename - remove invalid characters
                        invalid_chars = '<>:"/\\|?*'
                        for char in invalid_chars:
                            custom_filename = custom_filename.replace(char, '_')
                    
                    # Use custom filename or default to processed_{idx}
                    if custom_filename:
                        base_filename = custom_filename
                    else:
                        base_filename = f"processed_{idx}"
                    
                    # Generate document from template
                    template_ext = Path(job.metadata['job_template_path']).suffix
                    processed_doc = output_dir / f"{base_filename}{template_ext}"
                    print(f"Processing row {idx}: Generating {base_filename}{template_ext}...")
                    
                    self.template_processor.process_template(
                        job.metadata['job_template_path'],
                        row_data,
                        str(processed_doc)
                    )
                    
                    # Verify processed document was created
                    if not processed_doc.exists():
                        raise RuntimeError(f"Failed to create processed document: {processed_doc}")
                    
                    print(f"Row {idx}: Template processed successfully")
                    
                    # Convert to requested formats
                    for output_format in job.output_formats:
                        # Skip pdf_merged format in loop - will be handled after all files are processed
                        if output_format == 'pdf_merged':
                            continue
                        
                        format_dir = output_dir / output_format
                        format_dir.mkdir(parents=True, exist_ok=True)
                        
                        print(f"Row {idx}: Converting to {output_format}...")
                        
                        # Pass Excel print settings if converting to PDF from Excel
                        print_settings = None
                        if output_format == 'pdf' and job.excel_print_settings:
                            template_ext = Path(job.metadata['job_template_path']).suffix.lower()
                            if template_ext in ['.xlsx', '.xls']:
                                print_settings = job.excel_print_settings
                                print(f"Row {idx}: Using Excel print settings for PDF conversion")
                        
                        try:
                            output_file = self.format_converter.convert(
                                str(processed_doc),
                                output_format,
                                str(format_dir),
                                print_settings
                            )
                            
                            # Verify output file was created
                            if not os.path.exists(output_file):
                                raise RuntimeError(f"Output file was not created: {output_file}")
                            
                            print(f"Row {idx}: Successfully created {output_format} file: {output_file}")
                            job.add_output_file(output_file)
                            
                        except Exception as conv_error:
                            print(f"Row {idx}: Error converting to {output_format}: {str(conv_error)}")
                            raise
                    
                    job.increment_processed()
                    print(f"Row {idx}: Completed successfully")
                    
                except Exception as e:
                    error_msg = f"Error processing row {idx}: {str(e)}"
                    print(error_msg)
                    import traceback
                    traceback.print_exc()
                    job.increment_failed()
                    if not job.error_message:
                        job.error_message = error_msg
                
                self.save_job_metadata(job)
            
            # Validate that we have output files
            if job.processed_records == 0:
                raise RuntimeError("No records were processed successfully")
            
            # Verify output directory has files
            output_files_exist = any(output_dir.rglob('*.*'))
            if not output_files_exist:
                raise RuntimeError(f"No output files were generated in {output_dir}")
            
            # Handle PDF merging if pdf_merged format was requested
            if 'pdf_merged' in job.output_formats:
                # Check if individual PDFs were also requested
                if 'pdf' not in job.output_formats:
                    # Need to create temporary PDFs for merging
                    print(f"Job {job.id}: Creating PDFs for merging...")
                    temp_pdf_dir = output_dir / 'pdf'
                    temp_pdf_dir.mkdir(parents=True, exist_ok=True)
                    
                    # Convert all processed documents to PDF
                    for processed_file in output_dir.glob(f"processed_*{Path(job.metadata['job_template_path']).suffix}"):
                        try:
                            print_settings = None
                            if job.excel_print_settings:
                                template_ext = Path(job.metadata['job_template_path']).suffix.lower()
                                if template_ext in ['.xlsx', '.xls']:
                                    print_settings = job.excel_print_settings
                            
                            output_file = self.format_converter.convert(
                                str(processed_file),
                                'pdf',
                                str(temp_pdf_dir),
                                print_settings
                            )
                            print(f"Job {job.id}: Created PDF for merging: {output_file}")
                        except Exception as e:
                            print(f"Job {job.id}: Error creating PDF for merging: {str(e)}")
                
                print(f"Job {job.id}: Merging PDF files...")
                self._merge_pdfs(output_dir, job)
            
            print(f"Job {job.id}: Creating ZIP archive...")
            
            # Create ZIP file with all outputs
            zip_path = self.get_job_dir(job.id) / f"job_{job.id}_output.zip"
            self._create_zip_archive(output_dir, zip_path)
            
            # Verify ZIP was created
            if not zip_path.exists():
                raise RuntimeError(f"Failed to create ZIP file: {zip_path}")
            
            zip_size = zip_path.stat().st_size
            if zip_size == 0:
                raise RuntimeError(f"ZIP file is empty: {zip_path}")
            
            print(f"Job {job.id}: ZIP created successfully ({zip_size} bytes)")
            job.set_zip_file(str(zip_path))
            
            # Copy ZIP to custom output directory if specified
            if job.output_directory and os.path.exists(job.output_directory):
                try:
                    custom_zip_path = os.path.join(job.output_directory, f"job_{job.id}_output.zip")
                    shutil.copy2(zip_path, custom_zip_path)
                    print(f"Output copied to: {custom_zip_path}")
                except Exception as e:
                    print(f"Failed to copy output to custom directory: {str(e)}")
            
            # Final validation
            print(f"Job {job.id}: Processing completed. Processed: {job.processed_records}, Failed: {job.failed_records}")
            
            # Mark as completed
            job.update_status(JobStatus.COMPLETED)
            
        except Exception as e:
            job.update_status(JobStatus.FAILED, str(e))
        
        self.save_job_metadata(job)
        return job
    
    def _create_zip_archive(self, source_dir: Path, zip_path: Path):
        """Create a ZIP archive from a directory."""
        if not source_dir.exists():
            raise RuntimeError(f"Source directory does not exist: {source_dir}")
        
        file_count = 0
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(source_dir):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(source_dir)
                    zipf.write(file_path, arcname)
                    file_count += 1
                    print(f"Added to ZIP: {arcname}")
        
        if file_count == 0:
            raise RuntimeError(f"No files found to archive in {source_dir}")
        
        print(f"ZIP archive created with {file_count} files")
    
    def get_job_output_files(self, job_id: str) -> List[str]:
        """
        Get list of output files for a job.
        
        Args:
            job_id: Job ID
            
        Returns:
            List of output file paths
        """
        job = self.get_job(job_id)
        if not job:
            return []
        return job.output_files
    
    def get_job_zip_file(self, job_id: str) -> Optional[str]:
        """
        Get ZIP file path for a job.
        
        Args:
            job_id: Job ID
            
        Returns:
            ZIP file path or None (returns absolute path)
        """
        job = self.get_job(job_id)
        if not job:
            return None
        
        if not job.zip_file_path:
            return None
        
        # Ensure the path is absolute
        zip_path = Path(job.zip_file_path)
        if not zip_path.is_absolute():
            # Convert relative path to absolute
            zip_path = self.jobs_dir.parent / job.zip_file_path
        
        return str(zip_path)
    
    def check_and_update_files(self, job_id: str) -> Dict:
        """
        Check if job files need updating and update if necessary.
        
        Args:
            job_id: Job ID
            
        Returns:
            Dictionary with update status
        """
        job = self.get_job(job_id)
        if not job:
            raise ValueError(f"Job not found: {job_id}")
        
        updates = {
            'template_updated': False,
            'data_updated': False
        }
        
        try:
            # Check template file
            if job.template_path and self.file_tracker.is_file_changed(job.template_path):
                template_info = self.file_tracker.track_file(job.template_path, force_update=True)
                job.local_template_path = template_info['local_path']
                updates['template_updated'] = True
            
            # Check data file
            if job.data_path and self.file_tracker.is_file_changed(job.data_path):
                data_info = self.file_tracker.track_file(job.data_path, force_update=True)
                job.local_data_path = data_info['local_path']
                updates['data_updated'] = True
            
            if updates['template_updated'] or updates['data_updated']:
                self.save_job_metadata(job)
        
        except Exception as e:
            updates['error'] = str(e)
        
        return updates
    
    def _merge_pdfs(self, output_dir: Path, job: Job):
        """
        Merge all PDF files from the pdf directory into a single merged.pdf.
        
        Args:
            output_dir: Output directory containing format subdirectories
            job: Job instance
        """
        from PyPDF2 import PdfMerger
        
        # Check if we have individual PDFs to merge
        pdf_dir = output_dir / 'pdf'
        if not pdf_dir.exists():
            print(f"Job {job.id}: No PDF directory found to merge")
            return
        
        # Get all PDF files sorted by name
        pdf_files = sorted(pdf_dir.glob('*.pdf'))
        if not pdf_files:
            print(f"Job {job.id}: No PDF files found to merge")
            return
        
        # Create pdf_merged directory
        merged_dir = output_dir / 'pdf_merged'
        merged_dir.mkdir(parents=True, exist_ok=True)
        merged_file = merged_dir / 'merged.pdf'
        
        try:
            # Create merger and add all PDFs
            merger = PdfMerger()
            
            for pdf_file in pdf_files:
                print(f"Job {job.id}: Adding {pdf_file.name} to merged PDF")
                merger.append(str(pdf_file))
            
            # Write merged PDF
            merger.write(str(merged_file))
            merger.close()
            
            # Verify merged file was created
            if not merged_file.exists():
                raise RuntimeError(f"Failed to create merged PDF: {merged_file}")
            
            merged_size = merged_file.stat().st_size
            print(f"Job {job.id}: Merged PDF created successfully ({merged_size} bytes, {len(pdf_files)} files merged)")
            
            # Add to job output files
            job.add_output_file(str(merged_file))
            
        except Exception as e:
            print(f"Job {job.id}: Error merging PDFs: {str(e)}")
            import traceback
            traceback.print_exc()
            # Don't fail the entire job if merge fails
    
    def get_dashboard_stats(self) -> Dict:
        """
        Get statistics for dashboard display.
        
        Returns:
            Dictionary with dashboard statistics
        """
        all_jobs = self.get_all_jobs()
        
        return {
            'total_jobs': len(all_jobs),
            'pending_jobs': len([j for j in all_jobs if j.status == JobStatus.PENDING]),
            'processing_jobs': len([j for j in all_jobs if j.status == JobStatus.PROCESSING]),
            'completed_jobs': len([j for j in all_jobs if j.status == JobStatus.COMPLETED]),
            'failed_jobs': len([j for j in all_jobs if j.status == JobStatus.FAILED]),
            'total_records_processed': sum(j.processed_records for j in all_jobs),
            'total_files_generated': sum(len(j.output_files) for j in all_jobs)
        }
