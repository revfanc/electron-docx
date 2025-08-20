# Word to HTML Converter

一个基于Electron的Word文档批量转换为HTML的桌面应用程序。

## 功能特性

- 🚀 **批量转换**: 支持同时转换多个Word文档
- 📁 **拖拽支持**: 可以直接拖拽文件到应用窗口
- 🎨 **现代化UI**: 美观的用户界面，支持响应式设计
- 📊 **实时进度**: 显示转换进度和结果统计
- 🔧 **自定义输出**: 可选择输出目录
- 📱 **跨平台**: 支持Windows、macOS和Linux

## 支持的文件格式

- `.docx` - Microsoft Word 2007及以上版本
- `.doc` - Microsoft Word 97-2003文档

## 安装和运行

### 开发环境

1. 确保已安装Node.js (版本14或更高)

2. 克隆项目并安装依赖：
```bash
git clone <repository-url>
cd electron-docx
npm install
```

3. 启动开发模式：
```bash
npm start
```

### 构建可执行文件

```bash
# 构建所有平台
npm run build

# 仅构建当前平台
npm run dist
```

## 使用说明

1. **选择文件**: 点击"选择文件"按钮或直接拖拽Word文档到应用窗口
2. **设置输出目录**: 点击"选择目录"按钮选择HTML文件的保存位置
3. **开始转换**: 点击"开始转换"按钮开始批量转换
4. **查看结果**: 转换完成后会显示详细的转换结果和统计信息

## 技术栈

- **Electron**: 跨平台桌面应用框架
- **Mammoth.js**: Word文档解析和转换库
- **HTML/CSS/JavaScript**: 前端界面
- **Node.js**: 后端逻辑

## 项目结构

```
electron-docx/
├── main.js          # Electron主进程
├── index.html       # 主界面HTML
├── renderer.js      # 渲染进程脚本
├── package.json     # 项目配置
└── README.md        # 项目说明
```

## 开发说明

### 主进程 (main.js)
- 负责创建应用窗口
- 处理文件选择和目录选择对话框
- 执行Word到HTML的转换逻辑

### 渲染进程 (renderer.js)
- 处理用户界面交互
- 管理文件列表和转换状态
- 显示转换进度和结果

## 许可证

MIT License

## 贡献

欢迎提交Issue和Pull Request来改进这个项目。
