# Office Preview - Obsidian 插件

在Obsidian中本地预览Office文件（Word、Excel、PowerPoint）。

## 功能特性

- 📄 支持 .docx, .doc (Word文档)
- 📊 支持 .xlsx, .xls (Excel表格)  
- 📽️ 支持 .pptx, .ppt (PowerPoint演示文稿)
- 🎨 统一的预览界面风格
- 🌙 支持暗色模式

## 安装方式

### 方式一：BRAT插件安装（推荐）
1. 安装BRAT插件
2. 在BRAT中添加此仓库: `https://github.com/Paul-H1992/obsidian-office-preview`
3. 点击"Check for updates"然后"Install"

### 方式二：手动安装
1. 下载最新发布版本
2. 解压到 `.obsidian/plugins/office-preview/` 目录
3. 重启Obsidian
4. 在设置中启用插件

## 使用方法

### 方法1：右键菜单预览
1. 在文件列表中右键点击Office文件
2. 选择 "Preview Office File"

### 方法2：代码块嵌入
```
```office-preview
path/to/your/file.docx
```
```

## 版本历史

### v1.0.0 (2026-04-20)
- 初始版本发布
- 支持Word、Excel、PowerPoint文件预览
- 统一的预览界面
- 暗色模式支持

## 技术栈

- TypeScript
- Obsidian API

## License

MIT
