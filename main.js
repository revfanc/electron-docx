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
      contextIsolation: false
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

// 生成CSS样式
function generateCSS(options) {
  let css = `
    body {
        font-family: ${options.fontFamily};
        font-size: ${options.fontSize}px;
        line-height: ${options.lineHeight};
        margin: ${options.pageMargin}px;
        color: ${options.textColor};
    }
    
    h1, h2, h3, h4, h5, h6 {
        color: ${options.headingColor};
        margin-top: ${options.headingMargin}px;
        margin-bottom: ${options.headingMargin / 2}px;
        ${options.headingBold ? 'font-weight: bold;' : ''}
        ${options.headingUnderline ? 'text-decoration: underline;' : ''}
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
        border: 1px solid ${options.tableBorderColor};
        padding: ${options.tablePadding}px;
        text-align: left;
    }
    
    th {
        background-color: ${options.tableHeaderBg};
    }
    
    ${options.tableStriped ? `
    tr:nth-child(even) {
        background-color: rgba(0,0,0,0.02);
    }
    ` : ''}
    
    img {
        max-width: ${options.imageMaxWidth}%;
        height: auto;
        ${options.imageResponsive ? '' : 'width: auto;'}
        ${options.imageCenter ? 'display: block; margin: 0 auto;' : ''}
    }
    
    ${options.addPageBreaks ? `
    .page-break {
        page-break-before: always;
    }
    ` : ''}
    
    ${options.customCSS}
  `;
  
  return css;
}

// 创建mammoth转换选项
function createMammothOptions(userOptions) {
  const options = {};

  // 如果不保留原始样式，添加基本样式映射
  if (!userOptions.preserveStyles) {
    options.styleMap = [
      "p[style-name='Heading 1'] => h1:fresh",
      "p[style-name='Heading 2'] => h2:fresh",
      "p[style-name='Heading 3'] => h3:fresh",
      "p[style-name='Heading 4'] => h4:fresh",
      "p[style-name='Heading 5'] => h5:fresh",
      "p[style-name='Heading 6'] => h6:fresh",
      "p[style-name='Normal'] => p:fresh"
    ];
  } else {
    options.styleMap = [];
  }

  // 如果不保留列表格式
  if (!userOptions.preserveLists) {
    options.styleMap.push(
      "p[style-name='List Paragraph'] => p:fresh"
    );
  }

  // 如果不保留超链接
  if (!userOptions.preserveLinks) {
    options.transformDocument = (document) => {
      const elements = document.getElementsByTagName("a");
      for (let i = elements.length - 1; i >= 0; i--) {
        const element = elements[i];
        const text = element.textContent;
        element.parentNode.replaceChild(document.createTextNode(text), element);
      }
      return document;
    };
  }

  // 图片处理选项
  if (userOptions.preserveImages) {
    options.convertImage = mammoth.images.imgElement((image) => {
      return image.read().then((imageBuffer) => {
        const base64 = imageBuffer.toString('base64');
        const mimeType = image.contentType;
        return {
          src: `data:${mimeType};base64,${base64}`
        };
      });
    });
  }

  return options;
}

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

ipcMain.handle('convert-file', async (event, filePath, outputPath, options = {}) => {
  try {
    const mammothOptions = createMammothOptions(options);
    const result = await mammoth.convertToHtml({ path: filePath, ...mammothOptions });
    const html = result.value;
    
    // 生成CSS样式
    const css = generateCSS(options);
    
    // 创建完整的HTML结构
    const fullHtml = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${path.basename(filePath, path.extname(filePath))}</title>
    <style>
        ${css}
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

ipcMain.handle('batch-convert', async (event, files, outputDir, options = {}) => {
  const results = [];
  
  for (const filePath of files) {
    try {
      const fileName = path.basename(filePath, path.extname(filePath));
      const outputPath = path.join(outputDir, `${fileName}.html`);
      
      const mammothOptions = createMammothOptions(options);
      const result = await mammoth.convertToHtml({ path: filePath, ...mammothOptions });
      const html = result.value;
      
      // 生成CSS样式
      const css = generateCSS(options);
      
      const fullHtml = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${fileName}</title>
    <style>
        ${css}
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
