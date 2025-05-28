let selectedFiles = [];
let convertedFiles = []; // 変換完了ファイルを管理

const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const folderInput = document.getElementById('folderInput');
const fileList = document.getElementById('fileList');
const convertBtn = document.getElementById('convertBtn');
const progress = document.getElementById('progress');
const progressBar = document.getElementById('progressBar');
const progressText = document.getElementById('progressText');
const results = document.getElementById('results');

const allowedTypes = ['.doc', '.docx', '.ppt', '.pptx', '.xls', '.xlsx'];

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const files = Array.from(e.dataTransfer.files);
    handleFiles(files);
});

fileInput.addEventListener('change', (e) => {
    const files = Array.from(e.target.files);
    handleFiles(files);
});

folderInput.addEventListener('change', (e) => {
    const files = Array.from(e.target.files);
    handleFiles(files);
});

function handleFiles(files) {
    const validFiles = files.filter(file => {
        const ext = getFileExtension(file.name);
        return allowedTypes.includes(ext);
    });

    if (validFiles.length === 0) {
        showError('サポートされているファイル形式（.doc, .docx, .ppt, .pptx, .xls, .xlsx）を選択してください。');
        return;
    }

    validFiles.forEach(file => {
        if (!selectedFiles.find(f => f.name === file.name && f.size === file.size)) {
            selectedFiles.push(file);
        }
    });

    updateFileList();
    updateConvertButton();
}

function getFileExtension(filename) {
    return filename.toLowerCase().substring(filename.lastIndexOf('.'));
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function updateFileList() {
    fileList.innerHTML = '';
    
    selectedFiles.forEach((file, index) => {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        fileItem.innerHTML = `
            <div>
                <div class="file-name">${escapeHtml(file.name)}</div>
                <div class="file-size">${formatFileSize(file.size)}</div>
            </div>
            <button class="remove-btn" onclick="removeFile(${index})">×</button>
        `;
        fileList.appendChild(fileItem);
    });
}

function removeFile(index) {
    selectedFiles.splice(index, 1);
    updateFileList();
    updateConvertButton();
}

function updateConvertButton() {
    convertBtn.disabled = selectedFiles.length === 0;
}

async function convertFiles() {
    if (selectedFiles.length === 0) return;

    const formData = new FormData();
    selectedFiles.forEach(file => {
        formData.append('files', file);
    });

    convertBtn.disabled = true;
    progress.style.display = 'block';
    results.innerHTML = '';
    
    let progressValue = 0;
    const progressInterval = setInterval(() => {
        if (progressValue < 90) {
            progressValue += Math.random() * 10;
            progressBar.style.width = Math.min(progressValue, 90) + '%';
        }
    }, 200);

    try {
        const response = await fetch('/convert', {
            method: 'POST',
            body: formData
        });

        clearInterval(progressInterval);
        progressBar.style.width = '100%';
        progressText.textContent = '変換完了！';

        const result = await response.json();

        if (result.success) {
            convertedFiles = convertedFiles.concat(result.files); // 変換済みファイルを追加
            showResults(result.files);
            updateDownloadAllButton(); // まとめてダウンロードボタンの状態更新
            selectedFiles = [];
            updateFileList();
        } else {
            showError(result.message);
        }
    } catch (error) {
        clearInterval(progressInterval);
        showError('変換中にエラーが発生しました: ' + error.message);
    } finally {
        setTimeout(() => {
            progress.style.display = 'none';
            progressBar.style.width = '0%';
            progressText.textContent = '変換中...';
            convertBtn.disabled = false;
        }, 2000);
    }
}

function showResults(files) {
    results.innerHTML = '<h3>✅ 変換完了</h3>';
    
    files.forEach(file => {
        const resultItem = document.createElement('div');
        resultItem.className = 'result-item';
        const displayText = file.sheetCount 
            ? `${escapeHtml(file.originalName)} (${file.sheetCount}シート結合) → ${escapeHtml(file.pdfFileName)}`
            : `${escapeHtml(file.originalName)} → ${escapeHtml(file.pdfFileName)}`;
            
        resultItem.innerHTML = `
            <div>
                <div class="file-name">${displayText}</div>
            </div>
            <button class="download-btn" onclick="downloadFile('${escapeHtml(file.downloadUrl)}', '${escapeHtml(file.pdfFileName)}')">
                ダウンロード
            </button>
        `;
        results.appendChild(resultItem);
    });
}

function downloadFile(url, filename) {
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

function updateDownloadAllButton() {
    const downloadAllBtn = document.getElementById('downloadAllBtn');
    if (convertedFiles.length > 0) {
        downloadAllBtn.style.display = 'block';
        downloadAllBtn.textContent = `全てダウンロード (${convertedFiles.length}ファイル)`;
    } else {
        downloadAllBtn.style.display = 'none';
    }
}

async function downloadAllFiles() {
    if (convertedFiles.length === 0) {
        showError('ダウンロードできるファイルがありません。');
        return;
    }

    const downloadAllBtn = document.getElementById('downloadAllBtn');
    const originalText = downloadAllBtn.textContent;
    downloadAllBtn.textContent = 'ZIP作成中...';
    downloadAllBtn.disabled = true;

    try {
        const fileNames = convertedFiles.map(file => file.pdfFileName);
        
        const response = await fetch('/download-zip', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ fileNames })
        });

        if (response.ok) {
            // ZIPファイルのダウンロード
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `converted_files_${Date.now()}.zip`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
        } else {
            const error = await response.json();
            showError(error.error || 'ZIP作成に失敗しました');
        }
    } catch (error) {
        showError('ZIPダウンロード中にエラーが発生しました: ' + error.message);
    } finally {
        downloadAllBtn.textContent = originalText;
        downloadAllBtn.disabled = false;
    }
}

function clearAllFiles() {
    convertedFiles = [];
    results.innerHTML = '';
    updateDownloadAllButton();
}

function showError(message) {
    const errorDiv = document.createElement('div');
    errorDiv.className = 'error-message';
    errorDiv.textContent = message;
    results.innerHTML = '';
    results.appendChild(errorDiv);
    
    setTimeout(() => {
        errorDiv.remove();
    }, 5000);
}