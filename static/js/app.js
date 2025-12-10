// Document Automation System - Frontend JavaScript

const API_BASE = '/api';
let refreshInterval = null;
let templatesData = [];  // Store templates for multi-template mode

// Initialize on page load
document.addEventListener('DOMContentLoaded', function() {
    loadDashboardData();
    startAutoRefresh();
    setupFormHandlers();
});

// Auto-refresh jobs every 5 seconds
function startAutoRefresh() {
    refreshInterval = setInterval(() => {
        loadDashboardData();
    }, 5000);
}

function stopAutoRefresh() {
    if (refreshInterval) {
        clearInterval(refreshInterval);
    }
}

// Load dashboard stats and jobs
async function loadDashboardData() {
    try {
        // Load stats
        const statsResponse = await fetch(`${API_BASE}/dashboard/stats`);
        const statsData = await statsResponse.json();
        
        if (statsData.success) {
            updateDashboardStats(statsData.stats);
        }
        
        // Load jobs
        const jobsResponse = await fetch(`${API_BASE}/jobs`);
        const jobsData = await jobsResponse.json();
        
        if (jobsData.success) {
            renderJobs(jobsData.jobs);
        }
    } catch (error) {
        console.error('Error loading dashboard data:', error);
        showError('Failed to load dashboard data');
    }
}

// Update dashboard statistics
function updateDashboardStats(stats) {
    document.getElementById('stat-total').textContent = stats.total_jobs || 0;
    document.getElementById('stat-processing').textContent = stats.processing_jobs || 0;
    document.getElementById('stat-completed').textContent = stats.completed_jobs || 0;
    document.getElementById('stat-failed').textContent = stats.failed_jobs || 0;
}

// Render jobs grid
function renderJobs(jobs) {
    const container = document.getElementById('jobs-container');
    
    if (jobs.length === 0) {
        container.innerHTML = `
            <div class="text-center py-12">
                <i class="fas fa-inbox text-6xl text-gray-300 mb-4"></i>
                <p class="text-gray-500 text-lg">No jobs yet. Create your first job to get started!</p>
            </div>
        `;
        return;
    }
    
    const jobsHTML = jobs.map(job => createJobCard(job)).join('');
    container.innerHTML = `<div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">${jobsHTML}</div>`;
}

// Create job card HTML
function createJobCard(job) {
    const statusClass = `status-${job.status}`;
    const statusIcon = {
        'pending': 'fa-clock',
        'processing': 'fa-spinner fa-spin',
        'completed': 'fa-check-circle',
        'failed': 'fa-exclamation-circle'
    }[job.status] || 'fa-question-circle';
    
    const progress = job.total_records > 0 
        ? Math.round((job.processed_records / job.total_records) * 100) 
        : 0;
    
    return `
        <div class="job-card bg-white rounded-lg shadow-md p-6 cursor-pointer hover:shadow-lg transition-shadow" onclick="openJobDetails('${job.id}')">
            <div class="flex justify-between items-start mb-4">
                <div>
                    <h3 class="text-lg font-semibold text-gray-900 truncate">Job ${job.id.substring(0, 8)}</h3>
                    <p class="text-sm text-gray-500">${new Date(job.created_at).toLocaleString()}</p>
                </div>
                <span class="status-badge ${statusClass}">
                    <i class="fas ${statusIcon} mr-1"></i>
                    ${job.status}
                </span>
            </div>
            
            <div class="space-y-2 mb-4">
                <div class="flex items-center text-sm text-gray-600">
                    <i class="fas fa-file-alt w-4 mr-2"></i>
                    <span class="truncate" title="${job.template_path}">${getFileName(job.template_path)}</span>
                </div>
                <div class="flex items-center text-sm text-gray-600">
                    <i class="fas fa-database w-4 mr-2"></i>
                    <span class="truncate" title="${job.data_path}">${getFileName(job.data_path)}</span>
                </div>
                <div class="flex items-center text-sm text-gray-600">
                    <i class="fas fa-layer-group w-4 mr-2"></i>
                    <span>${job.output_formats.join(', ')}</span>
                </div>
            </div>
            
            ${job.status === 'processing' || job.status === 'completed' ? `
                <div class="mb-4">
                    <div class="flex justify-between text-sm text-gray-600 mb-1">
                        <span>Progress</span>
                        <span>${job.processed_records}/${job.total_records} (${progress}%)</span>
                    </div>
                    <div class="w-full bg-gray-200 rounded-full h-2">
                        <div class="bg-blue-600 h-2 rounded-full transition-all duration-300" style="width: ${progress}%"></div>
                    </div>
                </div>
            ` : ''}
            
            ${job.error_message ? `
                <div class="mb-4 p-3 bg-red-50 border border-red-200 rounded-lg">
                    <p class="text-sm text-red-800"><strong>Error:</strong> ${job.error_message}</p>
                </div>
            ` : ''}
            
            ${job.warnings && job.warnings.length > 0 ? `
                <div class="mb-4 p-3 bg-yellow-50 border border-yellow-200 rounded-lg">
                    <p class="text-sm text-yellow-800"><strong>Warning:</strong> ${job.warnings.join(', ')}</p>
                </div>
            ` : ''}
            
            <div class="flex flex-wrap gap-2" onclick="event.stopPropagation()">
                ${job.status === 'completed' ? `
                    <button onclick="downloadJob('${job.id}'); event.stopPropagation();" class="flex-1 bg-blue-600 hover:bg-blue-700 text-white px-3 py-2 rounded-lg text-sm transition">
                        <i class="fas fa-download mr-1"></i> Download
                    </button>
                    <button onclick="viewJobFiles('${job.id}'); event.stopPropagation();" class="flex-1 bg-gray-600 hover:bg-gray-700 text-white px-3 py-2 rounded-lg text-sm transition">
                        <i class="fas fa-eye mr-1"></i> View
                    </button>
                    <button onclick="rerunJob('${job.id}'); event.stopPropagation();" class="flex-1 bg-green-600 hover:bg-green-700 text-white px-3 py-2 rounded-lg text-sm transition">
                        <i class="fas fa-redo mr-1"></i> Rerun
                    </button>
                ` : job.status === 'pending' ? `
                    <button onclick="processJob('${job.id}'); event.stopPropagation();" class="flex-1 bg-green-600 hover:bg-green-700 text-white px-3 py-2 rounded-lg text-sm transition">
                        <i class="fas fa-play mr-1"></i> Process
                    </button>
                    <button onclick="editJob('${job.id}'); event.stopPropagation();" class="flex-1 bg-purple-600 hover:bg-purple-700 text-white px-3 py-2 rounded-lg text-sm transition">
                        <i class="fas fa-edit mr-1"></i> Edit
                    </button>
                ` : job.status === 'failed' ? `
                    <button onclick="rerunJob('${job.id}'); event.stopPropagation();" class="flex-1 bg-green-600 hover:bg-green-700 text-white px-3 py-2 rounded-lg text-sm transition">
                        <i class="fas fa-redo mr-1"></i> Retry
                    </button>
                ` : ''}
                <button onclick="deleteJob('${job.id}'); event.stopPropagation();" class="bg-red-600 hover:bg-red-700 text-white px-3 py-2 rounded-lg text-sm transition">
                    <i class="fas fa-trash"></i>
                </button>
            </div>
        </div>
    `;
}

// Extract filename from path
function getFileName(path) {
    if (!path) return 'N/A';
    return path.split(/[\\/]/).pop();
}

// Setup form handlers
function setupFormHandlers() {
    const form = document.getElementById('createJobForm');
    form.addEventListener('submit', handleCreateJob);
    
    // Add event listeners for template changes
    const templateFile = document.getElementById('template_file');
    const templatePath = document.getElementById('template_path');
    
    if (templateFile) {
        templateFile.addEventListener('change', toggleExcelPrintSettings);
    }
    if (templatePath) {
        templatePath.addEventListener('input', toggleExcelPrintSettings);
    }
}

// Handle create job form submission
async function handleCreateJob(e) {
    e.preventDefault();
    
    const form = document.getElementById('createJobForm');
    const editingJobId = form.dataset.editingJobId;
    const isEditing = !!editingJobId;
    
    const formData = new FormData();
    
    // Check if using multiple templates
    const useMultipleTemplates = document.getElementById('use_multiple_templates').checked;
    
    if (useMultipleTemplates) {
        // Handle multiple templates
        if (templatesData.length === 0) {
            showError('Please add at least one template');
            return;
        }
        
        // Upload all template files first
        const uploadedTemplates = [];
        for (const tmpl of templatesData) {
            // Create temporary FormData for file upload
            const tempFormData = new FormData();
            tempFormData.append('template_file', tmpl.file);
            
            try {
                const response = await fetch(`${API_BASE}/upload-temp-file`, {
                    method: 'POST',
                    body: tempFormData
                });
                
                const result = await response.json();
                if (result.success && result.path) {
                    uploadedTemplates.push({
                        path: result.path,
                        priority: tmpl.priority,
                        sheet: tmpl.sheet
                    });
                } else {
                    showError(`Failed to upload template: ${tmpl.fileName}`);
                    return;
                }
            } catch (error) {
                showError(`Error uploading template: ${tmpl.fileName}`);
                return;
            }
        }
        
        // Add templates as JSON
        console.log('Sending templates:', uploadedTemplates);
        formData.append('templates', JSON.stringify(uploadedTemplates));
    } else {
        // Single template mode (legacy)
        const templateSource = document.querySelector('input[name="template_source"]:checked').value;
        if (templateSource === 'file') {
            const templateFile = document.getElementById('template_file').files[0];
            if (!templateFile && !isEditing) {
                showError('Please select a template file');
                return;
            }
            if (templateFile) {
                formData.append('template_file', templateFile);
            }
        } else {
            const templatePath = document.getElementById('template_path').value;
            if (!templatePath) {
                showError('Please enter a template path');
                return;
            }
            formData.append('template_path', templatePath);
        }
    }
    
    // Get data
    const dataSource = document.querySelector('input[name="data_source"]:checked').value;
    if (dataSource === 'file') {
        const dataFile = document.getElementById('data_file').files[0];
        if (!dataFile && !isEditing) {
            showError('Please select a data file');
            return;
        }
        if (dataFile) {
            formData.append('data_file', dataFile);
        }
    } else {
        const dataPath = document.getElementById('data_path').value;
        if (!dataPath) {
            showError('Please enter a data path');
            return;
        }
        formData.append('data_path', dataPath);
    }
    
    // Get output formats
    const formats = Array.from(document.querySelectorAll('input[name="output_formats"]:checked'))
        .map(cb => cb.value);
    
    if (formats.length === 0) {
        showError('Please select at least one output format');
        return;
    }
    
    formData.append('output_formats', formats.join(','));
    formData.append('auto_process', 'true');
    
    // Add filename variable if specified
    const filenameVariable = document.getElementById('filename_variable').value.trim();
    if (filenameVariable) {
        formData.append('filename_variable', filenameVariable);
    }
    
    // Add tabname variable if specified
    const tabnameVariable = document.getElementById('tabname_variable').value.trim();
    if (tabnameVariable) {
        formData.append('tabname_variable', tabnameVariable);
    }
    
    // Add sheet names if specified
    const dataSheet = document.getElementById('data_sheet')?.value.trim();
    if (dataSheet) {
        formData.append('data_sheet', dataSheet);
    }
    
    const templateSheet = document.getElementById('template_sheet')?.value.trim();
    if (templateSheet) {
        formData.append('template_sheet', templateSheet);
    }
    
    // Add Excel print settings if applicable
    const excelPrintSettings = getExcelPrintSettings();
    if (excelPrintSettings) {
        formData.append('excel_print_settings', JSON.stringify(excelPrintSettings));
    }
    
    // Add output directory if specified
    const outputDirectory = document.getElementById('output_directory').value.trim();
    if (outputDirectory) {
        formData.append('output_directory', outputDirectory);
    }
    
    try {
        // Always create a new job - don't try to update existing jobs
        // This allows "editing" completed jobs by creating a new one with the same settings
        const url = `${API_BASE}/jobs`;
        const method = 'POST';
        
        const response = await fetch(url, {
            method: method,
            body: formData
        });
        
        const data = await response.json();
        
        if (data.success) {
            showSuccess('Job created successfully!');
            closeCreateJobModal();
            document.getElementById('createJobForm').reset();
            loadDashboardData();
        } else {
            showError(data.error || 'Failed to create job');
        }
    } catch (error) {
        console.error('Error creating job:', error);
        showError('Failed to create job');
    }
}

// Process a job
async function processJob(jobId) {
    try {
        const response = await fetch(`${API_BASE}/jobs/${jobId}/process`, {
            method: 'POST'
        });
        
        const data = await response.json();
        
        if (data.success) {
            showSuccess('Job processing started');
            loadDashboardData();
        } else {
            showError(data.error || 'Failed to start job');
        }
    } catch (error) {
        console.error('Error processing job:', error);
        showError('Failed to process job');
    }
}

// Download job output
async function downloadJob(jobId) {
    try {
        // First check if the job and file are ready
        const checkResponse = await fetch(`${API_BASE}/jobs/${jobId}`);
        const checkData = await checkResponse.json();
        
        if (!checkData.success) {
            showError(checkData.error || 'Failed to verify job status');
            return;
        }
        
        const job = checkData.job;
        
        if (job.status !== 'completed') {
            showError(`Job is not completed yet (status: ${job.status})`);
            return;
        }
        
        if (job.warnings && job.warnings.length > 0) {
            showError(`Cannot download: ${job.warnings.join(', ')}`);
            return;
        }
        
        // Proceed with download
        const downloadUrl = `${API_BASE}/jobs/${jobId}/download`;
        
        // Use fetch to check for errors before triggering download
        const response = await fetch(downloadUrl);
        
        if (!response.ok) {
            const errorData = await response.json();
            showError(errorData.error || 'Failed to download job output');
            return;
        }
        
        // If successful, trigger the download
        window.location.href = downloadUrl;
        
    } catch (error) {
        console.error('Error downloading job:', error);
        showError('Failed to download job output: ' + error.message);
    }
}

// View job files
async function viewJobFiles(jobId) {
    try {
        const response = await fetch(`${API_BASE}/jobs/${jobId}/files`);
        const data = await response.json();
        
        if (data.success) {
            showJobFilesModal(jobId, data.files);
        } else {
            showError(data.error || 'Failed to load files');
        }
    } catch (error) {
        console.error('Error loading job files:', error);
        showError('Failed to load job files');
    }
}

// Show job files in modal
function showJobFilesModal(jobId, files) {
    const content = document.getElementById('preview-content');
    
    let html = '<div class="space-y-4">';
    
    for (const [format, fileList] of Object.entries(files)) {
        html += `
            <div>
                <h4 class="text-lg font-semibold text-gray-900 mb-2 capitalize">${format}</h4>
                <div class="space-y-2">
        `;
        
        fileList.forEach(file => {
            html += `
                <div class="flex items-center justify-between p-3 bg-gray-50 rounded-lg">
                    <div class="flex items-center space-x-3">
                        <i class="fas fa-file text-blue-600"></i>
                        <span class="text-sm text-gray-900">${file.name}</span>
                        <span class="text-xs text-gray-500">(${formatFileSize(file.size)})</span>
                    </div>
                    <button onclick="previewFile('${jobId}', '${file.path}')" class="text-blue-600 hover:text-blue-800 text-sm">
                        <i class="fas fa-eye mr-1"></i> Preview
                    </button>
                </div>
            `;
        });
        
        html += '</div></div>';
    }
    
    html += '</div>';
    content.innerHTML = html;
    openPreviewModal();
}

// Preview a file
async function previewFile(jobId, filePath) {
    const fileName = getFileName(filePath);
    const ext = fileName.split('.').pop().toLowerCase();
    
    // filePath is already a relative path from the backend (e.g., "outputs/pdf/file.pdf")
    if (ext === 'pdf') {
        // Open PDF in new tab
        window.open(`${API_BASE}/jobs/${jobId}/preview/${filePath}`, '_blank');
    } else if (ext === 'docx' || ext === 'xlsx' || ext === 'msg') {
        // Download these files
        window.location.href = `${API_BASE}/jobs/${jobId}/preview/${filePath}`;
    }
}

// Edit a job
async function editJob(jobId) {
    try {
        const response = await fetch(`${API_BASE}/jobs/${jobId}`);
        const data = await response.json();
        
        console.log('Edit job response:', data);
        
        if (data.success && data.job) {
            const job = data.job;
            console.log('Job data:', job);
            
            // Open modal
            openCreateJobModal();
            
            // Store job ID for updating
            document.getElementById('createJobForm').dataset.editingJobId = jobId;
            document.querySelector('#createJobModal h3').textContent = 'Edit Job';
            
            // Check if job has multiple templates
            if (job.templates && job.templates.length > 0) {
                // Enable multiple templates mode
                document.getElementById('use_multiple_templates').checked = true;
                toggleMultipleTemplates();
                
                // Note: Can't easily repopulate file inputs for security reasons
                // Show a message instead
                const multiSection = document.getElementById('multiple-templates-section');
                const notice = document.createElement('div');
                notice.className = 'bg-yellow-50 border border-yellow-200 rounded p-3 mb-3';
                notice.innerHTML = `
                    <p class="text-sm text-yellow-800">
                        <i class="fas fa-info-circle mr-2"></i>
                        This job uses ${job.templates.length} templates. You'll need to re-add them to edit.
                    </p>
                `;
                multiSection.insertBefore(notice, multiSection.firstChild);
            } else {
                // Populate single template
                if (job.template_path) {
                    document.querySelector('input[name="template_source"][value="path"]').checked = true;
                    toggleTemplateInput();
                    document.getElementById('template_path').value = job.template_path;
                }
            }
            
            // Populate data
            if (job.data_path) {
                document.querySelector('input[name="data_source"][value="path"]').checked = true;
                toggleDataInput();
                document.getElementById('data_path').value = job.data_path;
            }
            
            // Populate output formats
            document.querySelectorAll('input[name="output_formats"]').forEach(cb => {
                cb.checked = job.output_formats.includes(cb.value);
            });
            
            // Populate filename variable
            if (job.metadata && job.metadata.filename_variable) {
                document.getElementById('filename_variable').value = job.metadata.filename_variable;
            }
            
            // Populate tabname variable
            if (job.metadata && job.metadata.tabname_variable) {
                document.getElementById('tabname_variable').value = job.metadata.tabname_variable;
            }
            
            // Toggle tabname variable visibility based on output formats
            toggleTabnameVariable();
            
            // Load and populate sheet selections if Excel files
            if (job.template_path && (job.template_path.toLowerCase().endsWith('.xlsx') || job.template_path.toLowerCase().endsWith('.xls'))) {
                loadTemplateSheets();
                if (job.metadata && job.metadata.template_sheet) {
                    setTimeout(() => {
                        document.getElementById('template_sheet').value = job.metadata.template_sheet;
                    }, 500);
                }
            }
            
            if (job.data_path && (job.data_path.toLowerCase().endsWith('.xlsx') || job.data_path.toLowerCase().endsWith('.xls'))) {
                loadDataSheets();
                if (job.metadata && job.metadata.data_sheet) {
                    setTimeout(() => {
                        document.getElementById('data_sheet').value = job.metadata.data_sheet;
                    }, 500);
                }
            }
            
            // Populate output directory
            if (job.output_directory) {
                document.getElementById('output_directory').value = job.output_directory;
            }
            
            // Populate Excel print settings if available
            if (job.excel_print_settings) {
                // Trigger display of Excel settings
                toggleExcelPrintSettings();
                
                const settings = job.excel_print_settings;
                if (settings.orientation) {
                    document.querySelector(`input[name="orientation"][value="${settings.orientation}"]`).checked = true;
                }
                if (settings.paper_size) {
                    document.getElementById('paper_size').value = settings.paper_size;
                }
                if (settings.margins) {
                    document.getElementById('margin_left').value = settings.margins.left || 0.75;
                    document.getElementById('margin_right').value = settings.margins.right || 0.75;
                    document.getElementById('margin_top').value = settings.margins.top || 1.0;
                    document.getElementById('margin_bottom').value = settings.margins.bottom || 1.0;
                }
                if (settings.scaling) {
                    const scalingType = settings.scaling.type || 'percent';
                    const scalingRadio = document.querySelector(`input[name="scaling_type"][value="${scalingType}"]`);
                    if (scalingRadio) {
                        scalingRadio.checked = true;
                    }
                    toggleScalingInputs();
                    
                    if (scalingType === 'percent') {
                        document.getElementById('scale_percent').value = settings.scaling.value || 100;
                    } else if (scalingType === 'fit_to') {
                        document.getElementById('fit_width').value = settings.scaling.width || 1;
                        document.getElementById('fit_height').value = settings.scaling.height || 1;
                    }
                    // Other scaling types (no_scaling, fit_sheet_on_one_page, etc.) don't need input values
                }
                if (settings.center_horizontally) {
                    document.getElementById('center_h').checked = true;
                }
                if (settings.center_vertically) {
                    document.getElementById('center_v').checked = true;
                }
            }
        } else {
            showError('Failed to load job data');
        }
    } catch (error) {
        console.error('Error loading job:', error);
        showError('Failed to load job data');
    }
}

// Open job details modal for editing
async function openJobDetails(jobId) {
    try {
        const response = await fetch(`${API_BASE}/jobs/${jobId}`);
        const data = await response.json();
        
        if (data.success && data.job) {
            const job = data.job;
            
            // Show a modal with job details and edit option
            const modalContent = `
                <div class="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50" onclick="closeJobDetailsModal()">
                    <div class="bg-white rounded-lg shadow-xl max-w-2xl w-full m-4 max-h-[90vh] overflow-y-auto" onclick="event.stopPropagation()">
                        <div class="p-6">
                            <div class="flex justify-between items-start mb-4">
                                <h2 class="text-2xl font-bold text-gray-900">Job Details</h2>
                                <button onclick="closeJobDetailsModal()" class="text-gray-400 hover:text-gray-600">
                                    <i class="fas fa-times text-xl"></i>
                                </button>
                            </div>
                            
                            <div class="space-y-4">
                                <div>
                                    <label class="block text-sm font-medium text-gray-700 mb-1">Job ID</label>
                                    <p class="text-gray-900">${job.id}</p>
                                </div>
                                
                                <div>
                                    <label class="block text-sm font-medium text-gray-700 mb-1">Status</label>
                                    <span class="status-badge status-${job.status}">
                                        <i class="fas ${job.status === 'pending' ? 'fa-clock' : job.status === 'processing' ? 'fa-spinner fa-spin' : job.status === 'completed' ? 'fa-check-circle' : 'fa-exclamation-circle'} mr-1"></i>
                                        ${job.status}
                                    </span>
                                </div>
                                
                                <div>
                                    <label class="block text-sm font-medium text-gray-700 mb-1">Template File</label>
                                    <p class="text-gray-900 break-all">${job.template_path || 'N/A'}</p>
                                </div>
                                
                                <div>
                                    <label class="block text-sm font-medium text-gray-700 mb-1">Data File</label>
                                    <p class="text-gray-900 break-all">${job.data_path || 'N/A'}</p>
                                </div>
                                
                                <div>
                                    <label class="block text-sm font-medium text-gray-700 mb-1">Output Formats</label>
                                    <p class="text-gray-900">${job.output_formats.join(', ')}</p>
                                </div>
                                
                                ${job.output_directory ? `
                                    <div>
                                        <label class="block text-sm font-medium text-gray-700 mb-1">Output Directory</label>
                                        <p class="text-gray-900 break-all">${job.output_directory}</p>
                                    </div>
                                ` : ''}
                                
                                <div>
                                    <label class="block text-sm font-medium text-gray-700 mb-1">Progress</label>
                                    <p class="text-gray-900">${job.processed_records} / ${job.total_records} records processed</p>
                                    ${job.failed_records > 0 ? `<p class="text-red-600">${job.failed_records} records failed</p>` : ''}
                                </div>
                                
                                <div>
                                    <label class="block text-sm font-medium text-gray-700 mb-1">Created</label>
                                    <p class="text-gray-900">${new Date(job.created_at).toLocaleString()}</p>
                                </div>
                                
                                ${job.error_message ? `
                                    <div>
                                        <label class="block text-sm font-medium text-gray-700 mb-1">Error Message</label>
                                        <p class="text-red-600 break-words">${job.error_message}</p>
                                    </div>
                                ` : ''}
                                
                                <div class="flex gap-3 pt-4 border-t">
                                    <button onclick="editJobFromDetails('${job.id}')" class="flex-1 bg-purple-600 hover:bg-purple-700 text-white px-4 py-2 rounded-lg transition">
                                        <i class="fas fa-edit mr-2"></i>Edit & Rerun
                                    </button>
                                    <button onclick="closeJobDetailsModal()" class="flex-1 bg-gray-500 hover:bg-gray-600 text-white px-4 py-2 rounded-lg transition">
                                        Close
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            `;
            
            // Add modal to body
            const modalDiv = document.createElement('div');
            modalDiv.id = 'jobDetailsModal';
            modalDiv.innerHTML = modalContent;
            document.body.appendChild(modalDiv);
        } else {
            showError('Failed to load job details');
        }
    } catch (error) {
        console.error('Error loading job details:', error);
        showError('Failed to load job details');
    }
}

// Close job details modal
function closeJobDetailsModal() {
    const modal = document.getElementById('jobDetailsModal');
    if (modal) {
        modal.remove();
    }
}

// Edit job from details modal
function editJobFromDetails(jobId) {
    closeJobDetailsModal();
    editJob(jobId);
}

// Rerun a job with same settings
async function rerunJob(jobId) {
    if (!confirm('This will create a new job with the same settings and reprocess all data. Continue?')) {
        return;
    }
    
    try {
        const response = await fetch(`${API_BASE}/jobs/${jobId}/rerun`, {
            method: 'POST'
        });
        
        const data = await response.json();
        
        if (data.success) {
            showSuccess('Job rerun started successfully!');
            loadDashboardData();
        } else {
            showError(data.error || 'Failed to rerun job');
        }
    } catch (error) {
        console.error('Error rerunning job:', error);
        showError('Failed to rerun job');
    }
}

// Delete a job
async function deleteJob(jobId) {
    if (!confirm('Are you sure you want to delete this job? This action cannot be undone.')) {
        return;
    }
    
    try {
        const response = await fetch(`${API_BASE}/jobs/${jobId}`, {
            method: 'DELETE'
        });
        
        const data = await response.json();
        
        if (data.success) {
            showSuccess('Job deleted successfully');
            loadDashboardData();
        } else {
            showError(data.error || 'Failed to delete job');
        }
    } catch (error) {
        console.error('Error deleting job:', error);
        showError('Failed to delete job');
    }
}

// Modal functions
function openCreateJobModal() {
    document.getElementById('createJobModal').classList.add('active');
    // Delay settings check slightly to allow DOM to settle
    setTimeout(() => {
        toggleExcelPrintSettings();
        toggleTabnameVariable();
    }, 100);
}

function closeCreateJobModal() {
    document.getElementById('createJobModal').classList.remove('active');
    document.getElementById('createJobForm').reset();
    delete document.getElementById('createJobForm').dataset.editingJobId;
    document.querySelector('#createJobModal h3').textContent = 'Create New Job';
    
    // Reset multi-template state
    templatesData = [];
    renderTemplatesTable();
    document.getElementById('use_multiple_templates').checked = false;
    toggleMultipleTemplates();
    
    // Remove any notice messages from edit mode
    const notices = document.querySelectorAll('#multiple-templates-section .bg-yellow-50');
    notices.forEach(notice => notice.remove());
}

function openPreviewModal() {
    document.getElementById('previewModal').classList.add('active');
}

function closePreviewModal() {
    document.getElementById('previewModal').classList.remove('active');
}

// Toggle input visibility
function toggleTemplateInput() {
    const source = document.querySelector('input[name="template_source"]:checked').value;
    document.getElementById('template-file-input').style.display = source === 'file' ? 'block' : 'none';
    document.getElementById('template-path-input').style.display = source === 'path' ? 'block' : 'none';
    toggleExcelPrintSettings();
}

function toggleDataInput() {
    const source = document.querySelector('input[name="data_source"]:checked').value;
    document.getElementById('data-file-input').style.display = source === 'file' ? 'block' : 'none';
    document.getElementById('data-path-input').style.display = source === 'path' ? 'block' : 'none';
}

// Handle template file change
function handleTemplateFileChange(input) {
    toggleExcelPrintSettings();
}

// Browse for template file path
async function browseTemplatePath() {
    try {
        const response = await fetch('/api/browse-file', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ type: 'template' })
        });
        
        const result = await response.json();
        
        if (result.success && result.path) {
            document.getElementById('template_path').value = result.path;
        } else if (result.error) {
            showError(result.error);
        }
    } catch (error) {
        showError('Failed to open file browser: ' + error.message);
    }
}

// Browse for data file path
async function browseDataPath() {
    try {
        const response = await fetch('/api/browse-file', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ type: 'data' })
        });
        
        const result = await response.json();
        
        if (result.success && result.path) {
            document.getElementById('data_path').value = result.path;
            loadDataSheets();
        } else if (result.error) {
            showError(result.error);
        }
    } catch (error) {
        showError('Failed to open file browser: ' + error.message);
    }
}

// Handle template file change
function handleTemplateFileChange(input) {
    if (input.files && input.files[0]) {
        const file = input.files[0];
        if (file.name.toLowerCase().endsWith('.xlsx') || file.name.toLowerCase().endsWith('.xls')) {
            loadTemplateSheetsFromFile(file);
        } else {
            document.getElementById('template-sheet-selection').style.display = 'none';
        }
    }
    toggleExcelPrintSettings();
}

// Handle data file change
function handleDataFileChange(input) {
    if (input.files && input.files[0]) {
        loadDataSheetsFromFile(input.files[0]);
    }
}

// Load template sheets from path
async function loadTemplateSheets() {
    const pathInput = document.getElementById('template_path');
    if (!pathInput || !pathInput.value) return;
    
    const path = pathInput.value.trim();
    if (!path.toLowerCase().endsWith('.xlsx') && !path.toLowerCase().endsWith('.xls')) {
        document.getElementById('template-sheet-selection').style.display = 'none';
        return;
    }
    
    try {
        const formData = new FormData();
        formData.append('file_path', path);
        
        const response = await fetch(`${API_BASE}/excel/sheets`, {
            method: 'POST',
            body: formData
        });
        
        const result = await response.json();
        
        if (result.success && result.sheets) {
            populateTemplateSheets(result.sheets);
        }
    } catch (error) {
        console.error('Error loading template sheets:', error);
    }
}

// Load template sheets from uploaded file
async function loadTemplateSheetsFromFile(file) {
    try {
        const formData = new FormData();
        formData.append('file', file);
        
        const response = await fetch(`${API_BASE}/excel/sheets`, {
            method: 'POST',
            body: formData
        });
        
        const result = await response.json();
        
        if (result.success && result.sheets) {
            populateTemplateSheets(result.sheets);
        }
    } catch (error) {
        console.error('Error loading template sheets:', error);
    }
}

// Populate template sheet dropdown
function populateTemplateSheets(sheets) {
    const select = document.getElementById('template_sheet');
    const container = document.getElementById('template-sheet-selection');
    
    if (!select || !container) return;
    
    // Clear existing options except first
    select.innerHTML = '<option value="">-- All Sheets / Auto-detect --</option>';
    
    if (sheets && sheets.length > 1) {
        sheets.forEach(sheet => {
            const option = document.createElement('option');
            option.value = sheet;
            option.textContent = sheet;
            select.appendChild(option);
        });
        container.style.display = 'block';
    } else {
        container.style.display = 'none';
    }
}

// Load data sheets from path
async function loadDataSheets() {
    const pathInput = document.getElementById('data_path');
    if (!pathInput || !pathInput.value) return;
    
    const path = pathInput.value.trim();
    if (!path.toLowerCase().endsWith('.xlsx') && !path.toLowerCase().endsWith('.xls')) {
        document.getElementById('data-sheet-selection').style.display = 'none';
        return;
    }
    
    try {
        const formData = new FormData();
        formData.append('file_path', path);
        
        const response = await fetch(`${API_BASE}/excel/sheets`, {
            method: 'POST',
            body: formData
        });
        
        const result = await response.json();
        
        if (result.success && result.sheets) {
            populateDataSheets(result.sheets, result.detected_sheet);
        }
    } catch (error) {
        console.error('Error loading data sheets:', error);
    }
}

// Load data sheets from uploaded file
async function loadDataSheetsFromFile(file) {
    try {
        const formData = new FormData();
        formData.append('file', file);
        
        const response = await fetch(`${API_BASE}/excel/sheets`, {
            method: 'POST',
            body: formData
        });
        
        const result = await response.json();
        
        if (result.success && result.sheets) {
            populateDataSheets(result.sheets, result.detected_sheet);
        }
    } catch (error) {
        console.error('Error loading data sheets:', error);
    }
}

// Populate data sheet dropdown
function populateDataSheets(sheets, detectedSheet) {
    const select = document.getElementById('data_sheet');
    const container = document.getElementById('data-sheet-selection');
    const info = document.getElementById('detected-sheet-info');
    
    if (!select || !container) return;
    
    // Clear existing options except first
    select.innerHTML = '<option value="">-- Auto-detect --</option>';
    
    if (sheets && sheets.length > 0) {
        sheets.forEach(sheet => {
            const option = document.createElement('option');
            option.value = sheet;
            option.textContent = sheet;
            if (sheet === detectedSheet) {
                option.textContent += ' ✓';
            }
            select.appendChild(option);
        });
        
        if (sheets.length > 1) {
            container.style.display = 'block';
            
            if (detectedSheet && info) {
                info.textContent = `Auto-detected: "${detectedSheet}" (contains ##variable## headers)`;
                info.style.color = '#059669'; // green
            } else if (info) {
                info.textContent = 'No ##variable## headers detected. First sheet will be used.';
                info.style.color = '#d97706'; // orange
            }
        } else {
            container.style.display = 'none';
        }
    } else {
        container.style.display = 'none';
    }
}

// Browse for output directory
async function browseOutputDirectory() {
    try {
        const response = await fetch('/api/browse-directory', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({})
        });
        
        const result = await response.json();
        
        if (result.success && result.path) {
            document.getElementById('output_directory').value = result.path;
        } else if (result.error) {
            showError(result.error);
        }
    } catch (error) {
        showError('Failed to open directory browser: ' + error.message);
    }
}

// Toggle Excel print settings visibility
function toggleExcelPrintSettings() {
    const settingsDiv = document.getElementById('excel-print-settings');
    if (!settingsDiv) return;
    
    // Check if template is Excel and PDF is selected in output formats
    const pdfChecked = document.querySelector('input[name="output_formats"][value="pdf"]')?.checked;
    const pdfMergedChecked = document.querySelector('input[name="output_formats"][value="pdf_merged"]')?.checked;
    const isExcelTemplate = isExcelTemplateSelected();
    
    // Show if Excel template AND (individual PDF OR merged PDF is selected)
    settingsDiv.style.display = ((pdfChecked || pdfMergedChecked) && isExcelTemplate) ? 'block' : 'none';
}

// Check if selected template is Excel
function isExcelTemplateSelected() {
    const templateSource = document.querySelector('input[name="template_source"]:checked')?.value;
    
    if (templateSource === 'file') {
        const fileInput = document.getElementById('template_file');
        if (fileInput && fileInput.files.length > 0) {
            const fileName = fileInput.files[0].name.toLowerCase();
            return fileName.endsWith('.xlsx') || fileName.endsWith('.xls');
        }
    } else if (templateSource === 'path') {
        const pathInput = document.getElementById('template_path');
        if (pathInput && pathInput.value) {
            const path = pathInput.value.toLowerCase().trim();
            return path.endsWith('.xlsx') || path.endsWith('.xls');
        }
    }
    
    return false;
}

// Toggle tab name variable input visibility
function toggleTabnameVariable() {
    const container = document.getElementById('tabnameVariableContainer');
    if (!container) return;
    
    // Check if excel_workbook is selected
    const excelWorkbookChecked = document.querySelector('input[name="output_formats"][value="excel_workbook"]')?.checked;
    
    // Show if excel_workbook is selected
    container.style.display = excelWorkbookChecked ? 'block' : 'none';
}

// Toggle scaling inputs
function toggleScalingInputs() {
    const scalingType = document.querySelector('input[name="scaling_type"]:checked')?.value;
    const percentInput = document.getElementById('scale_percent');
    const widthInput = document.getElementById('fit_width');
    const heightInput = document.getElementById('fit_height');
    
    // Enable/disable inputs based on scaling type
    if (scalingType === 'percent') {
        percentInput.disabled = false;
        widthInput.disabled = true;
        heightInput.disabled = true;
    } else if (scalingType === 'fit_to') {
        percentInput.disabled = true;
        widthInput.disabled = false;
        heightInput.disabled = false;
    } else {
        // For no_scaling, fit_sheet_on_one_page, fit_all_columns_on_one_page, fit_all_rows_on_one_page
        percentInput.disabled = true;
        widthInput.disabled = true;
        heightInput.disabled = true;
    }
}

// Get Excel print settings from form
function getExcelPrintSettings() {
    const settingsDiv = document.getElementById('excel-print-settings');
    if (!settingsDiv || settingsDiv.style.display === 'none') {
        return null;
    }
    
    const scalingType = document.querySelector('input[name="scaling_type"]:checked')?.value;
    const settings = {
        page_range: {
            from: parseInt(document.getElementById('page_from')?.value || 1),
            to: parseInt(document.getElementById('page_to')?.value || 0)
        },
        orientation: document.querySelector('input[name="orientation"]:checked')?.value || 'portrait',
        paper_size: document.getElementById('paper_size')?.value || 'a4',
        margins: {
            left: parseFloat(document.getElementById('margin_left')?.value || 0.75),
            right: parseFloat(document.getElementById('margin_right')?.value || 0.75),
            top: parseFloat(document.getElementById('margin_top')?.value || 1.0),
            bottom: parseFloat(document.getElementById('margin_bottom')?.value || 1.0)
        },
        scaling: {
            type: scalingType,
            value: scalingType === 'percent' ? parseInt(document.getElementById('scale_percent')?.value || 100) : null,
            width: scalingType === 'fit_to' ? parseInt(document.getElementById('fit_width')?.value || 1) : null,
            height: scalingType === 'fit_to' ? parseInt(document.getElementById('fit_height')?.value || 1) : null
        },
        center_horizontally: document.getElementById('center_horizontally')?.checked || false,
        center_vertically: document.getElementById('center_vertically')?.checked || false,
        ignore_print_areas: document.getElementById('ignore_print_areas')?.checked || false
    };
    
    return settings;
}

// Utility functions
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

function showError(message) {
    alert('Error: ' + message);
}

function showSuccess(message) {
    alert(message);
}

// Multi-Template Management
function toggleMultipleTemplates() {
    const isChecked = document.getElementById('use_multiple_templates').checked;
    const singleSection = document.getElementById('single-template-section');
    const multipleSection = document.getElementById('multiple-templates-section');
    
    if (isChecked) {
        singleSection.style.display = 'none';
        multipleSection.style.display = 'block';
    } else {
        singleSection.style.display = 'block';
        multipleSection.style.display = 'none';
        templatesData = [];
        renderTemplatesTable();
    }
}

async function addTemplate() {
    const fileInput = document.getElementById('multi_template_file');
    
    if (!fileInput.files || fileInput.files.length === 0) {
        showError('Please select a template file first');
        return;
    }
    
    const file = fileInput.files[0];
    const fileName = file.name;
    const fileExt = fileName.toLowerCase().substring(fileName.lastIndexOf('.'));
    
    if (!['.docx', '.xlsx', '.msg'].includes(fileExt)) {
        showError('Invalid file format. Supported: .docx, .xlsx, .msg');
        return;
    }
    
    // Detect sheet for Excel files
    let detectedSheet = null;
    if (fileExt === '.xlsx' || fileExt === '.xls') {
        try {
            const formData = new FormData();
            formData.append('file', file);
            formData.append('is_template', 'true');
            
            const response = await fetch(`${API_BASE}/excel/sheets`, {
                method: 'POST',
                body: formData
            });
            
            const result = await response.json();
            if (result.success && result.detected_sheet) {
                detectedSheet = result.detected_sheet;
            }
        } catch (error) {
            console.error('Error detecting sheet:', error);
        }
    }
    
    // Add to templates list
    const priority = templatesData.length + 1;
    templatesData.push({
        file: file,
        fileName: fileName,
        sheet: detectedSheet,
        priority: priority
    });
    
    renderTemplatesTable();
    fileInput.value = '';  // Clear file input
}

function renderTemplatesTable() {
    const container = document.getElementById('templates-table-container');
    const tbody = document.getElementById('templates-table-body');
    
    if (templatesData.length === 0) {
        container.style.display = 'none';
        return;
    }
    
    container.style.display = 'block';
    tbody.innerHTML = '';
    
    templatesData.forEach((template, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td class="px-3 py-2">
                <div class="flex items-center space-x-1">
                    <button type="button" onclick="movePriority(${index}, -1)" class="text-gray-500 hover:text-gray-700" ${index === 0 ? 'disabled' : ''}>
                        <i class="fas fa-arrow-up"></i>
                    </button>
                    <button type="button" onclick="movePriority(${index}, 1)" class="text-gray-500 hover:text-gray-700" ${index === templatesData.length - 1 ? 'disabled' : ''}>
                        <i class="fas fa-arrow-down"></i>
                    </button>
                    <span class="ml-2 font-medium">${template.priority}</span>
                </div>
            </td>
            <td class="px-3 py-2 text-sm">${template.fileName}</td>
            <td class="px-3 py-2 text-sm">
                ${template.sheet ? `
                    <select onchange="updateTemplateSheet(${index}, this.value)" class="text-xs border rounded px-2 py-1">
                        <option value="">Auto-detect</option>
                        <option value="${template.sheet}" selected>${template.sheet} ✓</option>
                    </select>
                ` : '<span class="text-gray-400">N/A</span>'}
            </td>
            <td class="px-3 py-2">
                <button type="button" onclick="removeTemplate(${index})" class="text-red-600 hover:text-red-800">
                    <i class="fas fa-trash"></i>
                </button>
            </td>
        `;
        tbody.appendChild(row);
    });
}

function movePriority(index, direction) {
    const newIndex = index + direction;
    if (newIndex < 0 || newIndex >= templatesData.length) return;
    
    // Swap elements
    [templatesData[index], templatesData[newIndex]] = [templatesData[newIndex], templatesData[index]];
    
    // Update priorities
    templatesData.forEach((tmpl, idx) => {
        tmpl.priority = idx + 1;
    });
    
    renderTemplatesTable();
}

function updateTemplateSheet(index, sheet) {
    templatesData[index].sheet = sheet || null;
}

function removeTemplate(index) {
    templatesData.splice(index, 1);
    
    // Update priorities
    templatesData.forEach((tmpl, idx) => {
        tmpl.priority = idx + 1;
    });
    
    renderTemplatesTable();
}
