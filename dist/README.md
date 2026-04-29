# Office Preview - Obsidian 插件

在 Obsidian 中直接预览 Word、Excel、PowerPoint 文件。

## 功能

- ✅ **Word (.docx)** - 完整预览，支持格式
- ✅ **Excel (.xlsx, .xls, .csv)** - 表格预览，前100行
- ⚠️ **PowerPoint (.pptx)** - 提示外部打开（技术限制）

## 安装

1. 下载插件文件夹到你的 Obsidian 仓库：
   ```
   你的Vault/.obsidian/plugins/office-preview/
   ```

2. 在 Obsidian 中启用社区插件：
   - 设置 → 社区插件 → 开启

3. 找到并启用 **Office Preview** 插件

## 使用方法

### 方法1：命令面板
- 按 `Ctrl/Cmd + P` 打开命令面板
- 输入 "Preview Office File"
- 回车执行

### 方法2：右键菜单
- 在文件上右键
- 选择 "📄 预览 Office 文件"

### 方法3：点击文件后使用命令

## 支持格式

| 格式 | 扩展名 | 预览效果 |
|------|--------|---------|
| Word | .docx | 完整富文本预览 |
| Excel | .xlsx, .xls, .csv | 表格预览（仅前100行）|
| PowerPoint | .pptx | 提示外部打开 |

## 技术说明

- 使用 [mammoth.js](https://github.com/mwilliamson/mammoth.js) 解析 DOCX
- 使用 [SheetJS (xlsx)](https://github.com/SheetJS/sheetjs) 解析 Excel
- 库通过 CDN 自动加载，无需额外安装

## 文件结构

```
office-preview/
├── manifest.json   # 插件配置
├── main.js        # 插件代码（CDN加载依赖）
└── README.md      # 说明文档
```

## 问题排查

**Q: 提示"库加载中"？**
A: 等待几秒后重试，需要网络连接加载CDN库。

**Q: Excel 显示乱码？**
A: 部分特殊编码可能不完全兼容。

**Q: 无法打开 PowerPoint？**
A: 这是技术限制，PPTX 需要外部程序打开。

---

🦞 Made with love by 代码如峰
