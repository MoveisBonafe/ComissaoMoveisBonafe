// JavaScript for enhanced user experience

document.addEventListener('DOMContentLoaded', function() {
    const form = document.getElementById('uploadForm');
    const submitBtn = document.getElementById('submitBtn');
    const excelInput = document.getElementById('excel_files');
    
    // Create progress indicator
    const progressIndicator = document.createElement('div');
    progressIndicator.className = 'progress-indicator';
    document.body.appendChild(progressIndicator);
    
    // File validation
    function validateFile(input, allowedTypes, maxSize = 16 * 1024 * 1024) {
        const file = input.files[0];
        if (!file) return true;
        
        // Check file size
        if (file.size > maxSize) {
            showAlert('Arquivo muito grande. Tamanho máximo: 16MB', 'error');
            input.value = '';
            return false;
        }
        
        // Check file type
        const fileExtension = file.name.split('.').pop().toLowerCase();
        if (!allowedTypes.includes(fileExtension)) {
            showAlert(`Tipo de arquivo não permitido. Use: ${allowedTypes.join(', ')}`, 'error');
            input.value = '';
            return false;
        }
        
        return true;
    }
    
    // File input change handlers
    if (excelInput) {
        excelInput.addEventListener('change', function() {
            validateFile(this, ['xlsx']);
            updateFileInfo(this, 'excel-info');
        });
    }
    
    // Update file information display
    function updateFileInfo(input, infoId) {
        const file = input.files[0];
        let infoElement = document.getElementById(infoId);
        
        if (!infoElement) {
            infoElement = document.createElement('div');
            infoElement.id = infoId;
            infoElement.className = 'file-info mt-2';
            input.parentNode.appendChild(infoElement);
        }
        
        if (file) {
            const fileSize = (file.size / 1024 / 1024).toFixed(2);
            infoElement.innerHTML = `
                <small class="text-success">
                    <i data-feather="check-circle" class="me-1"></i>
                    ${file.name} (${fileSize} MB)
                </small>
            `;
            feather.replace();
        } else {
            infoElement.innerHTML = '';
        }
    }
    
    // Form submission handler
    form.addEventListener('submit', function(e) {
        // Validate form
        if (!excelInput.files || excelInput.files.length === 0) {
            e.preventDefault();
            showAlert('Por favor, selecione pelo menos um arquivo Excel antes de continuar.', 'error');
            return;
        }
        
        // Show loading state
        setLoadingState(true);
        
        // Update button text
        const originalText = submitBtn.innerHTML;
        submitBtn.innerHTML = '<i data-feather="loader" class="me-2"></i>Processando...';
        submitBtn.disabled = true;
        
        // Show progress indicator
        progressIndicator.classList.add('active');
        
        // Re-enable form after timeout (in case of error)
        setTimeout(() => {
            setLoadingState(false);
            submitBtn.innerHTML = originalText;
            submitBtn.disabled = false;
            progressIndicator.classList.remove('active');
            feather.replace();
        }, 30000); // 30 seconds timeout
    });
    
    // Set loading state
    function setLoadingState(loading) {
        if (loading) {
            submitBtn.classList.add('loading');
            document.body.style.cursor = 'wait';
        } else {
            submitBtn.classList.remove('loading');
            document.body.style.cursor = 'default';
        }
    }
    
    // Show alert message
    function showAlert(message, type = 'info') {
        const alertContainer = document.querySelector('.container .row').nextElementSibling || 
                              document.querySelector('.container .row');
        
        const alertDiv = document.createElement('div');
        alertDiv.className = `alert alert-${type === 'error' ? 'danger' : type} alert-dismissible fade show`;
        alertDiv.innerHTML = `
            <i data-feather="${type === 'error' ? 'alert-circle' : 'info'}" class="me-2"></i>
            ${message}
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        `;
        
        alertContainer.insertBefore(alertDiv, alertContainer.firstChild);
        feather.replace();
        
        // Auto-dismiss after 5 seconds
        setTimeout(() => {
            if (alertDiv.parentNode) {
                alertDiv.remove();
            }
        }, 5000);
    }
    
    // Drag and drop functionality
    function setupDragAndDrop(input) {
        const parent = input.parentNode;
        
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            parent.addEventListener(eventName, preventDefaults, false);
        });
        
        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }
        
        ['dragenter', 'dragover'].forEach(eventName => {
            parent.addEventListener(eventName, highlight, false);
        });
        
        ['dragleave', 'drop'].forEach(eventName => {
            parent.addEventListener(eventName, unhighlight, false);
        });
        
        function highlight(e) {
            parent.classList.add('dragover');
        }
        
        function unhighlight(e) {
            parent.classList.remove('dragover');
        }
        
        parent.addEventListener('drop', handleDrop, false);
        
        function handleDrop(e) {
            const dt = e.dataTransfer;
            const files = dt.files;
            
            if (files.length > 0) {
                input.files = files;
                input.dispatchEvent(new Event('change'));
            }
        }
    }
    
    // Setup drag and drop for file inputs
    if (excelInput) {
        setupDragAndDrop(excelInput);
    }
    
    // Auto-dismiss alerts after page load
    setTimeout(() => {
        const alerts = document.querySelectorAll('.alert');
        alerts.forEach(alert => {
            if (alert.querySelector('.btn-close')) {
                setTimeout(() => {
                    if (alert.parentNode) {
                        alert.remove();
                    }
                }, 3000);
            }
        });
    }, 100);
});

// Utility function for formatting file sizes
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Handle window beforeunload to warn about unsaved changes
window.addEventListener('beforeunload', function(e) {
    const form = document.getElementById('uploadForm');
    const excelFiles = document.getElementById('excel_files');
    const hasFiles = excelFiles && excelFiles.files.length > 0;
    
    if (hasFiles && !form.submitted) {
        e.preventDefault();
        e.returnValue = '';
    }
});
