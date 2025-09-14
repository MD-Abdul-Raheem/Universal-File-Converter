let selectedFile = null;
let selectedFormat = null;
let downloadPath = null;

// DOM elements
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const formatPills = document.querySelectorAll('.format-pill');
const convertBtn = document.getElementById('convertBtn');
const spinner = document.getElementById('spinner');
const btnText = convertBtn.querySelector('.btn-text');
const progressSection = document.getElementById('progressSection');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const statusMessage = document.getElementById('statusMessage');
const resultSection = document.getElementById('resultSection');
const resultFileName = document.getElementById('resultFileName');
const textPreview = document.getElementById('textPreview');
const previewText = document.getElementById('previewText');

// Event listeners
fileInput.addEventListener('change', handleFileSelect);
formatPills.forEach(pill => {
    pill.addEventListener('click', () => selectFormat(pill.dataset.format, pill));
});

// Drag and drop
uploadArea.addEventListener('dragover', handleDragOver);
uploadArea.addEventListener('dragleave', handleDragLeave);
uploadArea.addEventListener('drop', handleDrop);
uploadArea.addEventListener('click', () => fileInput.click());

function handleDragOver(e) {
    e.preventDefault();
    uploadArea.classList.add('dragover');
}

function handleDragLeave(e) {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
}

function handleDrop(e) {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
}

function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        handleFile(file);
    }
}

function handleFile(file) {
    selectedFile = file;
    fileName.textContent = file.name;
    fileSize.textContent = formatFileSize(file.size);
    
    uploadArea.style.display = 'none';
    fileInfo.style.display = 'flex';
    
    showNotification('📁', 'File uploaded successfully!', 'success');
    updateConvertButton();
}

function removeFile() {
    selectedFile = null;
    fileInput.value = '';
    uploadArea.style.display = 'block';
    fileInfo.style.display = 'none';
    resultSection.style.display = 'none';
    progressSection.style.display = 'none';
    
    updateConvertButton();
}

function selectFormat(format, pillElement) {
    selectedFormat = format;
    
    // Remove active class from all pills
    formatPills.forEach(pill => pill.classList.remove('active'));
    
    // Add active class to selected pill
    pillElement.classList.add('active');
    
    updateConvertButton();
}

function updateConvertButton() {
    convertBtn.disabled = !selectedFile || !selectedFormat;
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

async function convertFile() {
    if (!selectedFile || !selectedFormat) return;
    
    // Show loading state
    btnText.textContent = 'Converting...';
    spinner.style.display = 'block';
    convertBtn.disabled = true;
    
    // Show progress bar
    progressSection.style.display = 'block';
    resultSection.style.display = 'none';
    
    // Animate progress bar
    animateProgress();
    
    const formData = new FormData();
    formData.append('file', selectedFile);
    formData.append('format', selectedFormat);
    
    try {
        const response = await fetch('/convert', {
            method: 'POST',
            body: formData
        });
        
        const result = await response.json();
        
        if (result.success) {
            downloadPath = result.download_path;
            showResult(result);
            showNotification('✅', 'File converted successfully!', 'success');
        } else {
            throw new Error(result.error);
        }
    } catch (error) {
        alert('Conversion failed: ' + error.message);
        hideProgress();
    } finally {
        resetButton();
    }
}

function animateProgress() {
    let progress = 0;
    const interval = setInterval(() => {
        progress += Math.random() * 15;
        if (progress > 90) progress = 90;
        
        progressFill.style.width = progress + '%';
        progressText.textContent = Math.round(progress) + '%';
        
        if (progress >= 90) {
            clearInterval(interval);
        }
    }, 200);
    
    // Complete progress when conversion is done
    setTimeout(() => {
        clearInterval(interval);
        progressFill.style.width = '100%';
        progressText.textContent = '100%';
        statusMessage.textContent = 'Conversion complete!';
    }, 2000);
}

function showResult(result) {
    progressSection.style.display = 'none';
    resultSection.style.display = 'block';
    
    // Set result file name
    const fileExt = selectedFormat.toUpperCase();
    const baseName = selectedFile.name.split('.')[0];
    resultFileName.textContent = `${baseName}.${selectedFormat}`;
    
    // Show text preview if available
    if (result.text_content) {
        textPreview.style.display = 'block';
        previewText.value = result.text_content;
    } else {
        textPreview.style.display = 'none';
    }
}

function hideProgress() {
    progressSection.style.display = 'none';
    statusMessage.textContent = 'Converting...';
    progressFill.style.width = '0%';
    progressText.textContent = '0%';
}

function resetButton() {
    btnText.textContent = 'Convert File';
    spinner.style.display = 'none';
    convertBtn.disabled = !selectedFile || !selectedFormat;
}

function downloadFile() {
    if (downloadPath) {
        const link = document.createElement('a');
        link.href = `/download/${encodeURIComponent(downloadPath)}`;
        link.download = '';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        showNotification('💾', 'Download started!', 'download');
    }
}

function showNotification(icon, message, type) {
    const container = document.getElementById('notificationContainer');
    
    const notification = document.createElement('div');
    notification.className = `notification ${type}`;
    
    notification.innerHTML = `
        <div class="notification-icon">${icon}</div>
        <div class="notification-text">${message}</div>
    `;
    
    container.appendChild(notification);
    
    // Auto remove after 3 seconds
    setTimeout(() => {
        notification.style.animation = 'slideOut 0.4s ease-out forwards';
        setTimeout(() => {
            if (notification.parentNode) {
                notification.parentNode.removeChild(notification);
            }
        }, 400);
    }, 3000);
}

// PWA Install Popup
let deferredPrompt;

window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredPrompt = e;
    document.getElementById('installPopup').style.display = 'flex';
});

document.getElementById('installBtn').addEventListener('click', () => {
    if (deferredPrompt) {
        deferredPrompt.prompt();
        deferredPrompt.userChoice.then(() => {
            deferredPrompt = null;
            closePopup();
        });
    }
});

function closePopup() {
    document.getElementById('installPopup').style.display = 'none';
}