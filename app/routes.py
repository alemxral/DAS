"""
API Routes
Handles all API endpoints for the document automation system.
"""
import os
from flask import Blueprint, request, jsonify, send_file, current_app
from werkzeug.utils import secure_filename
from pathlib import Path
import threading

from services.job_manager import JobManager
from models.job import JobStatus

# Create blueprint
api_bp = Blueprint('api', __name__, url_prefix='/api')

# Initialize job manager (will be set in create_app)
job_manager = None


def get_job_manager():
    """Get or create job manager instance."""
    global job_manager
    if job_manager is None:
        print(f"Creating JobManager with:")
        print(f"  jobs_dir: {current_app.config['JOBS_DIR']}")
        print(f"  storage_dir: {current_app.config['STORAGE_DIR']}")
        job_manager = JobManager(
            jobs_dir=current_app.config['JOBS_DIR'],
            storage_dir=current_app.config['STORAGE_DIR']
        )
    return job_manager


def allowed_file(filename, allowed_extensions):
    """Check if file has allowed extension."""
    return Path(filename).suffix.lower() in allowed_extensions


@api_bp.route('/jobs', methods=['GET'])
def get_jobs():
    """Get all jobs."""
    try:
        manager = get_job_manager()
        jobs = manager.get_all_jobs()
        
        # Sort by creation date (newest first)
        jobs.sort(key=lambda x: x.created_at, reverse=True)
        
        return jsonify({
            'success': True,
            'jobs': [job.to_dict() for job in jobs],
            'count': len(jobs)
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@api_bp.route('/jobs/<job_id>', methods=['GET'])
def get_job(job_id):
    """Get specific job details with validation."""
    try:
        manager = get_job_manager()
        job = manager.get_job(job_id)
        
        if not job:
            return jsonify({'success': False, 'error': 'Job not found'}), 404
        
        # Add validation warnings
        warnings = []
        if job.status == JobStatus.COMPLETED:
            # Check if ZIP file exists
            if job.zip_file_path:
                if not os.path.exists(job.zip_file_path):
                    warnings.append('Output ZIP file is missing')
            else:
                warnings.append('ZIP file path not set')
            
            # Check if output files exist
            if not job.output_files:
                warnings.append('No output files recorded')
        
        job_dict = job.to_dict()
        if warnings:
            job_dict['warnings'] = warnings
        
        return jsonify({
            'success': True,
            'job': job_dict
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@api_bp.route('/excel/sheets', methods=['POST'])
def get_excel_sheets():
    """Get list of sheets from an Excel file."""
    try:
        from services.document_parser import DocumentParser
        parser = DocumentParser()
        
        file_path = None
        
        # Handle file upload
        if 'file' in request.files:
            file = request.files['file']
            if file and file.filename:
                from werkzeug.utils import secure_filename
                filename = secure_filename(file.filename)
                upload_dir = Path(current_app.config['UPLOAD_DIR'])
                upload_dir.mkdir(parents=True, exist_ok=True)
                file_path = upload_dir / f"temp_{filename}"
                file.save(str(file_path))
        
        # Handle file path
        if not file_path and request.form.get('file_path'):
            file_path = request.form.get('file_path')
        
        if not file_path or not os.path.exists(file_path):
            return jsonify({'success': False, 'error': 'File not found'}), 400
        
        # Get sheets
        sheets = parser.get_excel_sheets(str(file_path))
        
        # Detect sheet with ##variable## headers
        detected_sheet = parser.detect_data_sheet(str(file_path))
        
        # Clean up temp file if uploaded
        if 'file' in request.files and file_path and os.path.exists(file_path):
            try:
                os.remove(file_path)
            except:
                pass
        
        return jsonify({
            'success': True,
            'sheets': sheets,
            'detected_sheet': detected_sheet
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@api_bp.route('/jobs', methods=['POST'])
def create_job():
    """Create a new job."""
    try:
        manager = get_job_manager()
        
        # Check if files are uploaded or paths provided
        template_path = None
        data_path = None
        
        # Handle file uploads
        if 'template_file' in request.files:
            template_file = request.files['template_file']
            if template_file and template_file.filename:
                if not allowed_file(template_file.filename, current_app.config['ALLOWED_TEMPLATE_EXTENSIONS']):
                    return jsonify({'success': False, 'error': 'Invalid template file format'}), 400
                
                filename = secure_filename(template_file.filename)
                upload_dir = Path(current_app.config['UPLOAD_DIR'])
                upload_dir.mkdir(parents=True, exist_ok=True)
                template_path = upload_dir / filename
                template_file.save(str(template_path))
        
        if 'data_file' in request.files:
            data_file = request.files['data_file']
            if data_file and data_file.filename:
                if not allowed_file(data_file.filename, current_app.config['ALLOWED_DATA_EXTENSIONS']):
                    return jsonify({'success': False, 'error': 'Invalid data file format'}), 400
                
                filename = secure_filename(data_file.filename)
                upload_dir = Path(current_app.config['UPLOAD_DIR'])
                upload_dir.mkdir(parents=True, exist_ok=True)
                data_path = upload_dir / filename
                data_file.save(str(data_path))
        
        # Handle file paths (if files not uploaded)
        if not template_path and request.form.get('template_path'):
            template_path = request.form.get('template_path')
            if not os.path.exists(template_path):
                return jsonify({'success': False, 'error': 'Template file not found'}), 400
        
        if not data_path and request.form.get('data_path'):
            data_path = request.form.get('data_path')
            if not os.path.exists(data_path):
                return jsonify({'success': False, 'error': 'Data file not found'}), 400
        
        # Validate inputs
        if not template_path or not data_path:
            return jsonify({'success': False, 'error': 'Both template and data files are required'}), 400
        
        # Get output formats
        output_formats = request.form.get('output_formats', 'pdf')
        if isinstance(output_formats, str):
            output_formats = [f.strip() for f in output_formats.split(',')]
        
        # Validate output formats
        for fmt in output_formats:
            if fmt not in current_app.config['AVAILABLE_OUTPUT_FORMATS']:
                return jsonify({'success': False, 'error': f'Invalid output format: {fmt}'}), 400
        
        # Get Excel print settings if provided
        excel_print_settings = None
        if request.form.get('excel_print_settings'):
            import json
            try:
                excel_print_settings = json.loads(request.form.get('excel_print_settings'))
            except:
                pass
        
        # Get filename variable if provided
        filename_variable = request.form.get('filename_variable', '##filename##').strip()
        
        # Get tabname variable if provided
        tabname_variable = request.form.get('tabname_variable', '##tabname##').strip()
        
        # Get sheet names if provided (for Excel files with multiple sheets)
        data_sheet = request.form.get('data_sheet', '').strip() or None
        template_sheet = request.form.get('template_sheet', '').strip() or None
        
        # Get output directory if provided
        output_directory = request.form.get('output_directory', '').strip()
        if output_directory and not os.path.exists(output_directory):
            try:
                os.makedirs(output_directory, exist_ok=True)
            except:
                output_directory = None  # Ignore if can't create
        
        # Create job
        job = manager.create_job(
            template_path=str(template_path),
            data_path=str(data_path),
            output_formats=output_formats,
            excel_print_settings=excel_print_settings,
            output_directory=output_directory if output_directory else None,
            filename_variable=filename_variable,
            tabname_variable=tabname_variable,
            data_sheet=data_sheet,
            template_sheet=template_sheet
        )
        
        # Start processing in background thread
        auto_process = request.form.get('auto_process', 'true').lower() == 'true'
        if auto_process:
            thread = threading.Thread(target=manager.process_job, args=(job.id,))
            thread.daemon = True
            thread.start()
        
        return jsonify({
            'success': True,
            'job': job.to_dict(),
            'message': 'Job created successfully'
        }), 201
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@api_bp.route('/jobs/<job_id>/process', methods=['POST'])
def process_job(job_id):
    """Start processing a job."""
    try:
        manager = get_job_manager()
        job = manager.get_job(job_id)
        
        if not job:
            return jsonify({'success': False, 'error': 'Job not found'}), 404
        
        if job.status != JobStatus.PENDING:
            return jsonify({'success': False, 'error': f'Job cannot be processed (status: {job.status.value})'}), 400
        
        # Process in background thread
        thread = threading.Thread(target=manager.process_job, args=(job_id,))
        thread.daemon = True
        thread.start()
        
        return jsonify({
            'success': True,
            'message': 'Job processing started',
            'job_id': job_id
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@api_bp.route('/jobs/<job_id>', methods=['PUT'])
def update_job(job_id):
    """Update an existing job."""
    try:
        manager = get_job_manager()
        job = manager.get_job(job_id)
        
        if not job:
            return jsonify({'success': False, 'error': 'Job not found'}), 404
        
        if job.status != JobStatus.PENDING:
            return jsonify({'success': False, 'error': 'Can only edit pending jobs'}), 400
        
        # Handle file updates (same logic as create)
        template_path = job.template_path
        data_path = job.data_path
        
        if 'template_file' in request.files:
            template_file = request.files['template_file']
            if template_file and template_file.filename:
                if not allowed_file(template_file.filename, current_app.config['ALLOWED_TEMPLATE_EXTENSIONS']):
                    return jsonify({'success': False, 'error': 'Invalid template file format'}), 400
                
                filename = secure_filename(template_file.filename)
                upload_dir = Path(current_app.config['UPLOAD_DIR'])
                upload_dir.mkdir(parents=True, exist_ok=True)
                template_path = upload_dir / filename
                template_file.save(str(template_path))
        elif request.form.get('template_path'):
            template_path = request.form.get('template_path')
        
        if 'data_file' in request.files:
            data_file = request.files['data_file']
            if data_file and data_file.filename:
                if not allowed_file(data_file.filename, current_app.config['ALLOWED_DATA_EXTENSIONS']):
                    return jsonify({'success': False, 'error': 'Invalid data file format'}), 400
                
                filename = secure_filename(data_file.filename)
                upload_dir = Path(current_app.config['UPLOAD_DIR'])
                upload_dir.mkdir(parents=True, exist_ok=True)
                data_path = upload_dir / filename
                data_file.save(str(data_path))
        elif request.form.get('data_path'):
            data_path = request.form.get('data_path')
        
        # Update job properties
        job.template_path = str(template_path)
        job.data_path = str(data_path)
        
        # Update output formats
        if request.form.get('output_formats'):
            output_formats = request.form.get('output_formats')
            if isinstance(output_formats, str):
                job.output_formats = [f.strip() for f in output_formats.split(',')]
        
        # Update Excel print settings
        if request.form.get('excel_print_settings'):
            import json
            try:
                job.excel_print_settings = json.loads(request.form.get('excel_print_settings'))
            except:
                pass
        
        # Update output directory
        output_directory = request.form.get('output_directory', '').strip()
        if output_directory:
            job.output_directory = output_directory
        
        # Re-track files with updated paths
        job_dir = manager.get_job_dir(job.id)
        template_info = manager.file_tracker.track_file(job.template_path)
        data_info = manager.file_tracker.track_file(job.data_path)
        
        job.template_file_id = template_info['file_id']
        job.local_template_path = template_info['local_path']
        job.data_file_id = data_info['file_id']
        job.local_data_path = data_info['local_path']
        
        # Copy files to job directory
        job_template_path = job_dir / f"template{Path(job.template_path).suffix}"
        job_data_path = job_dir / f"data{Path(job.data_path).suffix}"
        
        shutil.copy2(job.local_template_path, job_template_path)
        shutil.copy2(job.local_data_path, job_data_path)
        
        job.metadata['job_template_path'] = str(job_template_path)
        job.metadata['job_data_path'] = str(job_data_path)
        
        # Save updated job
        manager.save_job_metadata(job)
        
        return jsonify({
            'success': True,
            'job': job.to_dict(),
            'message': 'Job updated successfully'
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@api_bp.route('/jobs/<job_id>/rerun', methods=['POST'])
def rerun_job(job_id):
    """Rerun a job with the same settings."""
    try:
        manager = get_job_manager()
        original_job = manager.get_job(job_id)
        
        if not original_job:
            return jsonify({'success': False, 'error': 'Job not found'}), 404
        
        # Create new job with same settings
        filename_variable = original_job.metadata.get('filename_variable', '##filename##')
        tabname_variable = original_job.metadata.get('tabname_variable', '##tabname##')
        data_sheet = original_job.metadata.get('data_sheet', None)
        template_sheet = original_job.metadata.get('template_sheet', None)
        
        new_job = manager.create_job(
            template_path=original_job.template_path,
            data_path=original_job.data_path,
            output_formats=original_job.output_formats,
            excel_print_settings=original_job.excel_print_settings,
            output_directory=original_job.output_directory,
            filename_variable=filename_variable,
            tabname_variable=tabname_variable,
            data_sheet=data_sheet,
            template_sheet=template_sheet
        )
        
        # Start processing
        thread = threading.Thread(target=manager.process_job, args=(new_job.id,))
        thread.daemon = True
        thread.start()
        
        return jsonify({
            'success': True,
            'job': new_job.to_dict(),
            'message': 'Job rerun started successfully'
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@api_bp.route('/jobs/<job_id>', methods=['DELETE'])
def delete_job(job_id):
    """Delete a job."""
    try:
        manager = get_job_manager()
        success = manager.delete_job(job_id)
        
        if not success:
            return jsonify({'success': False, 'error': 'Job not found'}), 404
        
        return jsonify({
            'success': True,
            'message': 'Job deleted successfully'
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@api_bp.route('/jobs/<job_id>/download', methods=['GET'])
def download_job_output(job_id):
    """Download job output as ZIP file."""
    try:
        manager = get_job_manager()
        job = manager.get_job(job_id)
        
        if not job:
            return jsonify({'success': False, 'error': 'Job not found'}), 404
        
        if job.status != JobStatus.COMPLETED:
            return jsonify({
                'success': False, 
                'error': f'Job is not completed yet (status: {job.status.value})'
            }), 400
        
        zip_path = manager.get_job_zip_file(job_id)
        
        print(f"Download request for job {job_id}")
        print(f"ZIP path from metadata: {zip_path}")
        print(f"Jobs directory: {manager.jobs_dir}")
        
        if not zip_path:
            return jsonify({
                'success': False, 
                'error': 'ZIP file path not set in job metadata'
            }), 404
        
        # Convert relative path to absolute if needed
        zip_path_obj = Path(zip_path)
        if not zip_path_obj.is_absolute():
            # Path is relative, make it absolute from BASE_DIR
            from config.config import Config
            zip_path = str(Config.BASE_DIR / zip_path)
            print(f"Converted to absolute path: {zip_path}")
        
        if not os.path.exists(zip_path):
            # Try to find the ZIP file in case path is wrong
            job_dir = manager.get_job_dir(job_id)
            expected_zip = job_dir / f"job_{job_id}_output.zip"
            
            print(f"ZIP not found at metadata path, checking: {expected_zip}")
            print(f"Job directory exists: {job_dir.exists()}")
            
            if job_dir.exists():
                print(f"Job directory contents: {list(job_dir.iterdir())}")
            
            if expected_zip.exists():
                print(f"Found ZIP at expected location, updating metadata")
                zip_path = str(expected_zip)
                job.set_zip_file(zip_path)
                manager.save_job_metadata(job)
            else:
                # List all files in job directory for debugging
                files_in_dir = []
                if job_dir.exists():
                    for item in job_dir.rglob('*'):
                        if item.is_file():
                            files_in_dir.append(str(item.relative_to(job_dir)))
                
                return jsonify({
                    'success': False, 
                    'error': f'Output file not found. Expected at: {expected_zip}. Job directory: {job_dir}. Files found: {", ".join(files_in_dir) if files_in_dir else "none"}'
                }), 404
        
        # Verify file size
        file_size = os.path.getsize(zip_path)
        if file_size == 0:
            return jsonify({
                'success': False, 
                'error': 'Output file is empty. Job may have encountered errors.'
            }), 500
        
        return send_file(
            zip_path,
            as_attachment=True,
            download_name=f'job_{job_id}_output.zip',
            mimetype='application/zip'
        )
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


@api_bp.route('/jobs/<job_id>/files', methods=['GET'])
def get_job_files(job_id):
    """Get list of output files for a job."""
    try:
        manager = get_job_manager()
        files = manager.get_job_output_files(job_id)
        job_dir = manager.get_job_dir(job_id)
        
        # Convert to relative paths and organize by format
        output_files = {}
        for file_path in files:
            path = Path(file_path)
            format_name = path.parent.name
            if format_name not in output_files:
                output_files[format_name] = []
            
            # Get relative path from job directory
            try:
                relative_path = path.relative_to(job_dir)
                # Convert to forward slashes for URL compatibility
                relative_path_str = str(relative_path).replace('\\', '/')
            except ValueError:
                # If relative_to fails, try to extract from outputs onwards
                relative_path_str = str(path).replace('\\', '/')
                if '/outputs/' in relative_path_str:
                    relative_path_str = 'outputs/' + relative_path_str.split('/outputs/')[-1]
                else:
                    relative_path_str = path.name
            
            output_files[format_name].append({
                'name': path.name,
                'path': relative_path_str,
                'size': os.path.getsize(file_path) if os.path.exists(file_path) else 0
            })
        
        return jsonify({
            'success': True,
            'files': output_files
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@api_bp.route('/jobs/<job_id>/preview/<path:file_path>', methods=['GET'])
def preview_file(job_id, file_path):
    """Preview a specific output file."""
    try:
        print(f"[PREVIEW] job_id: {job_id}")
        print(f"[PREVIEW] file_path received: {file_path}")
        
        manager = get_job_manager()
        job = manager.get_job(job_id)
        
        if not job:
            return jsonify({'success': False, 'error': 'Job not found'}), 404
        
        # Construct full path
        job_dir = manager.get_job_dir(job_id)
        full_path = job_dir / file_path
        
        print(f"[PREVIEW] job_dir: {job_dir}")
        print(f"[PREVIEW] full_path: {full_path}")
        print(f"[PREVIEW] file exists: {full_path.exists()}")
        
        if not full_path.exists():
            return jsonify({'success': False, 'error': f'File not found: {full_path}'}), 404
        
        # Determine mime type
        ext = full_path.suffix.lower()
        mime_types = {
            '.pdf': 'application/pdf',
            '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            '.msg': 'application/vnd.ms-outlook'
        }
        
        return send_file(
            str(full_path),
            mimetype=mime_types.get(ext, 'application/octet-stream')
        )
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@api_bp.route('/dashboard/stats', methods=['GET'])
def get_dashboard_stats():
    """Get dashboard statistics."""
    try:
        manager = get_job_manager()
        stats = manager.get_dashboard_stats()
        
        return jsonify({
            'success': True,
            'stats': stats
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@api_bp.route('/formats', methods=['GET'])
def get_available_formats():
    """Get available output formats."""
    return jsonify({
        'success': True,
        'formats': current_app.config['AVAILABLE_OUTPUT_FORMATS']
    })


@api_bp.route('/health', methods=['GET'])
def health_check():
    """Health check endpoint."""
    return jsonify({
        'success': True,
        'status': 'healthy',
        'service': 'Document Automation API'
    })


@api_bp.route('/browse-file', methods=['POST'])
def browse_file():
    """Open native file dialog and return selected file path."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        
        # Get file type from request
        file_type = request.json.get('type', 'template')  # 'template' or 'data'
        
        # Create root window (hidden)
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        
        # Set file types based on request
        if file_type == 'template':
            filetypes = [
                ('All Template Files', '*.docx *.xlsx *.msg'),
                ('Word Documents', '*.docx'),
                ('Excel Files', '*.xlsx'),
                ('Outlook Messages', '*.msg'),
                ('All Files', '*.*')
            ]
        else:  # data
            filetypes = [
                ('Excel Files', '*.xlsx *.xls'),
                ('All Files', '*.*')
            ]
        
        # Open file dialog
        file_path = filedialog.askopenfilename(
            title=f'Select {file_type} file',
            filetypes=filetypes
        )
        
        root.destroy()
        
        if file_path:
            return jsonify({
                'success': True,
                'path': file_path
            })
        else:
            return jsonify({
                'success': False,
                'error': 'No file selected'
            })
            
    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'Failed to open file dialog: {str(e)}'
        }), 500


@api_bp.route('/browse-directory', methods=['POST'])
def browse_directory():
    """Open native directory dialog and return selected directory path."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        
        # Create root window (hidden)
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        
        # Open directory dialog
        dir_path = filedialog.askdirectory(
            title='Select output directory'
        )
        
        root.destroy()
        
        if dir_path:
            return jsonify({
                'success': True,
                'path': dir_path
            })
        else:
            return jsonify({
                'success': False,
                'error': 'No directory selected'
            })
            
    except Exception as e:
        return jsonify({
            'success': False,
            'error': f'Failed to open directory dialog: {str(e)}'
        }), 500
