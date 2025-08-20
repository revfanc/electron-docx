const { ipcRenderer } = require('electron');
const path = require('path');
const fs = require('fs-extra');

// 全局变量
let selectedFiles = [];
let outputDirectory = '';

// DOM元素
const fileInputArea = document.getElementById('fileInputArea');
const selectFilesBtn = document.getElementById('selectFilesBtn');
const fileListSection = document.getElementById('fileListSection');
const fileList = document.getElementById('fileList');
const clearFilesBtn = document.getElementById('clearFilesBtn');
const startConvertBtn = document.getElementById('startConvertBtn');
const selectOutputBtn = document.getElementById('selectOutputBtn');
const outputPath = document.getElementById('outputPath');
const progressSection = document.getElementById('progressSection');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const resultsSection = document.getElementById('resultsSection');
const results = document.getElementById('results');
const successCount = document.getElementById('successCount');
const errorCount = document.getElementById('errorCount');
const statusText = document.getElementById('statusText');

// 事件监听器
selectFilesBtn.addEventListener('click', selectFiles);
clearFilesBtn.addEventListener('click', clearFiles);
startConvertBtn.addEventListener('click', startConversion);
selectOutputBtn.addEventListener('click', selectOutputDirectory);

// 拖拽功能
fileInputArea.addEventListener('dragover', handleDragOver);
fileInputArea.addEventListener('dragleave', handleDragLeave);
fileInputArea.addEventListener('drop', handleDrop);

// 文件选择
async function selectFiles() {
    try {
        const filePaths = await ipcRenderer.invoke('select-files');
        if (filePaths && filePaths.length > 0) {
            addFiles(filePaths);
        }
    } catch (error) {
        showError('选择文件时出错: ' + error.message);
    }
}

// 添加文件到列表
function addFiles(filePaths) {
    const newFiles = filePaths.filter(filePath => {
        const ext = path.extname(filePath).toLowerCase();
        return ext === '.docx' || ext === '.doc';
    });

    if (newFiles.length === 0) {
        showError('没有选择有效的Word文档文件');
        return;
    }

    // 去重
    const existingPaths = selectedFiles.map(f => f.path);
    const uniqueNewFiles = newFiles.filter(filePath => !existingPaths.includes(filePath));

    if (uniqueNewFiles.length === 0) {
        showError('所有文件都已存在');
        return;
    }

    // 获取文件信息
    uniqueNewFiles.forEach(filePath => {
        try {
            const stats = fs.statSync(filePath);
            selectedFiles.push({
                path: filePath,
                name: path.basename(filePath),
                size: formatFileSize(stats.size)
            });
        } catch (error) {
            console.error('获取文件信息失败:', filePath, error);
        }
    });

    updateFileList();
    showFileListSection();
}

// 更新文件列表显示
function updateFileList() {
    fileList.innerHTML = '';
    
    selectedFiles.forEach((file, index) => {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        fileItem.innerHTML = `
            <div class="file-name">${file.name}</div>
            <div class="file-size">${file.size}</div>
            <button class="remove-file" onclick="removeFile(${index})">×</button>
        `;
        fileList.appendChild(fileItem);
    });
}

// 移除文件
function removeFile(index) {
    selectedFiles.splice(index, 1);
    updateFileList();
    
    if (selectedFiles.length === 0) {
        hideFileListSection();
    }
}

// 清空文件列表
function clearFiles() {
    selectedFiles = [];
    updateFileList();
    hideFileListSection();
}

// 显示文件列表区域
function showFileListSection() {
    fileListSection.classList.remove('hidden');
}

// 隐藏文件列表区域
function hideFileListSection() {
    fileListSection.classList.add('hidden');
}

// 选择输出目录
async function selectOutputDirectory() {
    try {
        const dirPath = await ipcRenderer.invoke('select-output-directory');
        if (dirPath) {
            outputDirectory = dirPath;
            outputPath.value = dirPath;
        }
    } catch (error) {
        showError('选择输出目录时出错: ' + error.message);
    }
}

// 开始转换
async function startConversion() {
    if (selectedFiles.length === 0) {
        showError('请先选择要转换的文件');
        return;
    }

    if (!outputDirectory) {
        showError('请先选择输出目录');
        return;
    }

    // 禁用按钮
    startConvertBtn.disabled = true;
    selectFilesBtn.disabled = true;
    clearFilesBtn.disabled = true;

    // 显示进度区域
    progressSection.classList.remove('hidden');
    resultsSection.classList.add('hidden');

    try {
        const filePaths = selectedFiles.map(f => f.path);
        const results = await ipcRenderer.invoke('batch-convert', filePaths, outputDirectory);
        
        showResults(results);
    } catch (error) {
        showError('转换过程中出错: ' + error.message);
    } finally {
        // 恢复按钮
        startConvertBtn.disabled = false;
        selectFilesBtn.disabled = false;
        clearFilesBtn.disabled = false;
        
        // 隐藏进度区域
        progressSection.classList.add('hidden');
    }
}

// 显示转换结果
function showResults(conversionResults) {
    results.innerHTML = '';
    
    let successCount = 0;
    let errorCount = 0;

    conversionResults.forEach(result => {
        const resultItem = document.createElement('div');
        resultItem.className = `result-item ${result.success ? 'result-success' : 'result-error'}`;
        
        const icon = result.success ? '✅' : '❌';
        resultItem.innerHTML = `
            <div class="result-icon">${icon}</div>
            <div>
                <div><strong>${result.file}</strong></div>
                <div>${result.message}</div>
                ${result.outputPath ? `<div style="font-size: 0.9rem; color: #7f8c8d;">输出: ${result.outputPath}</div>` : ''}
            </div>
        `;
        
        results.appendChild(resultItem);

        if (result.success) {
            successCount++;
        } else {
            errorCount++;
        }
    });

    // 更新统计
    document.getElementById('successCount').textContent = successCount;
    document.getElementById('errorCount').textContent = errorCount;
    
    if (errorCount === 0) {
        statusText.textContent = '所有文件转换成功！';
    } else if (successCount === 0) {
        statusText.textContent = '所有文件转换失败';
    } else {
        statusText.textContent = `转换完成，${successCount}个成功，${errorCount}个失败`;
    }

    resultsSection.classList.remove('hidden');
}

// 拖拽处理
function handleDragOver(e) {
    e.preventDefault();
    fileInputArea.classList.add('dragover');
}

function handleDragLeave(e) {
    e.preventDefault();
    fileInputArea.classList.remove('dragover');
}

function handleDrop(e) {
    e.preventDefault();
    fileInputArea.classList.remove('dragover');
    
    const files = Array.from(e.dataTransfer.files);
    const filePaths = files.map(file => file.path);
    
    addFiles(filePaths);
}

// 工具函数
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function showError(message) {
    // 创建错误提示
    const errorDiv = document.createElement('div');
    errorDiv.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: #e74c3c;
        color: white;
        padding: 15px 20px;
        border-radius: 5px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
        z-index: 1000;
        max-width: 300px;
        word-wrap: break-word;
    `;
    errorDiv.textContent = message;
    
    document.body.appendChild(errorDiv);
    
    // 3秒后自动移除
    setTimeout(() => {
        if (errorDiv.parentNode) {
            errorDiv.parentNode.removeChild(errorDiv);
        }
    }, 3000);
}

// 初始化
document.addEventListener('DOMContentLoaded', () => {
    // 设置默认输出目录为桌面
    const os = require('os');
    const desktopPath = path.join(os.homedir(), 'Desktop');
    outputDirectory = desktopPath;
    outputPath.value = desktopPath;
});
