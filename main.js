const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs-extra');
const mammoth = require('mammoth');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true
    },
    icon: path.join(__dirname, 'assets', 'icon.png'),
    title: 'Word to HTML Converter',
    show: false
  });

  mainWindow.loadFile('index.html');

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

// IPC处理程序
ipcMain.handle('select-files', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: 'Word Documents', extensions: ['docx', 'doc'] }
    ]
  });
  return result.filePaths;
});

ipcMain.handle('select-output-directory', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory']
  });
  return result.filePaths[0];
});

ipcMain.handle('convert-file', async (event, filePath, outputPath) => {
  try {
    const result = await mammoth.convertToHtml({ path: filePath });
    const html = result.value;
    
    // 创建基本的HTML结构
    const fullHtml = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${path.basename(filePath, path.extname(filePath))}</title>
    <style>
        body {
            font-family: 'Microsoft YaHei', Arial, sans-serif;
            line-height: 1.6;
            margin: 40px;
            color: #333;
        }
        h1, h2, h3, h4, h5, h6 {
            color: #2c3e50;
            margin-top: 30px;
            margin-bottom: 15px;
        }
        p {
            margin-bottom: 15px;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin: 20px 0;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
        }
        img {
            max-width: 100%;
            height: auto;
        }
        .page-break {
            page-break-before: always;
        }
    </style>
</head>
<body>
    ${html}
</body>
</html>`;
    
    await fs.writeFile(outputPath, fullHtml, 'utf8');
    return { success: true, message: '转换成功' };
  } catch (error) {
    return { success: false, message: `转换失败: ${error.message}` };
  }
});

ipcMain.handle('batch-convert', async (event, files, outputDir) => {
  const results = [];
  
  for (const filePath of files) {
    try {
      const fileName = path.basename(filePath, path.extname(filePath));
      const outputPath = path.join(outputDir, `${fileName}.html`);
      
      const result = await mammoth.convertToHtml({ path: filePath });
      const html = result.value;
      
      const fullHtml = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${fileName}</title>
    <style>
        body {
            font-family: 'Microsoft YaHei', Arial, sans-serif;
            line-height: 1.6;
            margin: 40px;
            color: #333;
        }
        h1, h2, h3, h4, h5, h6 {
            color: #2c3e50;
            margin-top: 30px;
            margin-bottom: 15px;
        }
        p {
            margin-bottom: 15px;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin: 20px 0;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
        }
        img {
            max-width: 100%;
            height: auto;
        }
        .page-break {
            page-break-before: always;
        }
    </style>
</head>
<body>
    ${html}
</body>
</html>`;
      
      await fs.writeFile(outputPath, fullHtml, 'utf8');
      results.push({
        file: path.basename(filePath),
        success: true,
        message: '转换成功',
        outputPath: outputPath
      });
    } catch (error) {
      results.push({
        file: path.basename(filePath),
        success: false,
        message: `转换失败: ${error.message}`
      });
    }
  }
  
  return results;
});
