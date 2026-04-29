const { App, Plugin, PluginSettingTab, Setting, Notice, Modal } = require('obsidian');

const DOCX_EXTENSIONS = ['docx'];
const EXCEL_EXTENSIONS = ['xlsx', 'xls', 'csv'];
const PPTX_EXTENSIONS = ['pptx', 'ppt'];

module.exports = class OfficePreviewPlugin extends Plugin {
  constructor(app) {
    super(app);
    this.settings = { defaultMode: 'right' };
    this.loadedLibs = { mammoth: false, xlsx: false };
  }

  async onload() {
    await this.loadSettings();
    
    this.addCommand({
      id: 'preview-office-file',
      name: 'Preview Office File',
      callback: () => this.previewActiveFile()
    });

    this.addSettingTab(new OfficePreviewSettingTab(this.app, this));
    
    // Register right-click menu
    this.registerEvent(this.app.workspace.on('file-menu', (menu, file) => {
      if (this.isOfficeFile(file.path)) {
        menu.addItem((item) => {
          item.setTitle('📄 预览 Office 文件').onClick(() => this.previewFile(file));
        });
      }
    }));

    // Load external libraries from CDN
    await this.loadLibraries();
    
    console.log('Office Preview plugin loaded');
  }

  onunload() {
    console.log('Office Preview plugin unloaded');
  }

  async loadLibraries() {
    return new Promise((resolve) => {
      // Load mammoth for DOCX
      if (!window.mammoth) {
        const script1 = document.createElement('script');
        script1.src = 'https://cdn.jsdelivr.net/npm/mammoth@1.6.0/mammoth.browser.min.js';
        script1.onload = () => { this.loadedLibs.mammoth = true; };
        document.head.appendChild(script1);
      } else {
        this.loadedLibs.mammoth = true;
      }
      
      // Load xlsx for Excel  
      if (!window.XLSX) {
        const script2 = document.createElement('script');
        script2.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
        script2.onload = () => { this.loadedLibs.xlsx = true; };
        document.head.appendChild(script2);
      } else {
        this.loadedLibs.xlsx = true;
      }
      
      // Wait for libs to load
      let attempts = 0;
      const check = () => {
        attempts++;
        if (this.loadedLibs.mammoth && this.loadedLibs.xlsx) {
          resolve(true);
        } else if (attempts < 50) {
          setTimeout(check, 100);
        } else {
          resolve(false);
        }
      };
      check();
    });
  }

  isOfficeFile(path) {
    const ext = path.split('.').pop().toLowerCase();
    return [...DOCX_EXTENSIONS, ...EXCEL_EXTENSIONS, ...PPTX_EXTENSIONS].includes(ext);
  }

  async loadSettings() {
    const data = await this.loadData();
    this.settings = { ...{ defaultMode: 'right' }, ...data };
  }

  async saveSettings() {
    await this.saveData(this.settings);
  }

  async previewActiveFile() {
    const file = this.app.workspace.getActiveFile();
    if (!file) {
      new Notice('没有打开的文件');
      return;
    }
    await this.previewFile(file);
  }

  async previewFile(file) {
    const ext = file.extension?.toLowerCase();
    try {
      if (DOCX_EXTENSIONS.includes(ext)) {
        await this.previewDocx(file);
      } else if (EXCEL_EXTENSIONS.includes(ext)) {
        await this.previewExcel(file);
      } else if (PPTX_EXTENSIONS.includes(ext)) {
        await this.previewPptx(file);
      } else {
        new Notice('不支持的文件格式');
      }
    } catch (e) {
      console.error(e);
      new Notice('预览失败: ' + e.message);
    }
  }

  async previewDocx(file) {
    if (!window.mammoth) {
      new Notice('库加载中，请稍后再试');
      return;
    }
    
    const buffer = await file.read();
    const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
    const result = await window.mammoth.convertToHtml({ arrayBuffer });
    
    const modal = new OfficePreviewModal(this.app, '📝 ' + file.name, result.value, 'docx');
    modal.open();
  }

  async previewExcel(file) {
    if (!window.XLSX) {
      new Notice('库加载中，请稍后再试');
      return;
    }
    
    const buffer = await file.read();
    const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
    const workbook = window.XLSX.read(arrayBuffer, { type: 'array' });
    
    let html = `<div class="excel-preview">
      <style>
        .excel-preview { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; padding: 20px; }
        .excel-preview h3 { color: #333; margin: 15px 0 10px; font-size: 16px; }
        .excel-preview table { border-collapse: collapse; width: 100%; margin-bottom: 20px; font-size: 13px; }
        .excel-preview th { background: #f5f5f5; font-weight: 600; text-align: left; }
        .excel-preview td, .excel-preview th { border: 1px solid #ddd; padding: 6px 10px; }
        .excel-preview tr:nth-child(even) { background: #fafafa; }
      </style>
      <h2>📊 ${this.escapeHtml(file.name)}</h2>`;
    
    for (let i = 0; i < Math.min(workbook.SheetNames.length, 5); i++) {
      const sheetName = workbook.SheetNames[i];
      const sheet = workbook.Sheets[sheetName];
      const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      
      html += `<h3>📄 ${this.escapeHtml(sheetName)}</h3><table><thead><tr>`;
      if (rows[0]) {
        rows[0].forEach(cell => { html += `<th>${this.escapeHtml(cell)}</th>`; });
      }
      html += '</tr></thead><tbody>';
      
      rows.slice(1, 101).forEach(row => {
        html += '<tr>';
        row.forEach(cell => { html += `<td>${this.escapeHtml(cell)}</td>`; });
        html += '</tr>';
      });
      html += '</tbody></table>';
    }
    html += '<p style="color:#888;font-size:12px;margin-top:10px">* 仅显示前100行</p></div>';
    
    const modal = new OfficePreviewModal(this.app, '📊 ' + file.name, html, 'excel');
    modal.open();
  }

  async previewPptx(file) {
    const html = `
      <div class="pptx-notice" style="text-align:center;padding:40px;font-family:-apple-system,sans-serif">
        <div style="font-size:64px;margin-bottom:20px">📊</div>
        <h2 style="color:#333;margin-bottom:15px">PowerPoint 文件</h2>
        <p style="color:#666;margin-bottom:20px">${this.escapeHtml(file.name)}</p>
        <p style="color:#888;font-size:14px">由于技术限制，PPTX 文件需要在外部程序中打开</p>
        <p style="color:#666;margin-top:15px">路径: <code style="background:#f5f5f5;padding:2px 6px;border-radius:3px">${this.escapeHtml(file.path)}</code></p>
        <button onclick="require('electron').shell.openPath('${this.escapeHtml(file.path)}')" 
          style="margin-top:20px;padding:10px 24px;background:#4CAF50;color:white;border:none;border-radius:6px;cursor:pointer;font-size:14px">
          在外部程序中打开
        </button>
      </div>
    `;
    const modal = new OfficePreviewModal(this.app, '📊 ' + file.name, html, 'pptx');
    modal.open();
  }

  escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = String(text);
    return div.innerHTML;
  }
};

class OfficePreviewModal extends Modal {
  constructor(app, title, content, type) {
    super(app);
    this.title = title;
    this.content = content;
    this.type = type;
  }

  onOpen() {
    const { contentEl } = this;
    contentEl.innerHTML = `
      <div class="office-preview-modal" style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif">
        <div style="display:flex;justify-content:space-between;align-items:center;padding:15px 20px;border-bottom:1px solid #eee">
          <h2 style="margin:0;font-size:18px;color:#333">${this.title}</h2>
          <button class="close-btn" style="background:none;border:none;font-size:24px;cursor:pointer;color:#999;padding:0;line-height:1">×</button>
        </div>
        <div class="preview-content" style="max-height:70vh;overflow-y:auto;padding:20px;background:#fafafa">
          ${this.content}
        </div>
      </div>
    `;
    contentEl.querySelector('.close-btn').onclick = () => this.close();
  }

  onClose() {
    const { contentEl } = this;
    contentEl.empty();
  }
}

class OfficePreviewSettingTab extends PluginSettingTab {
  constructor(app, plugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display() {
    const { containerEl } = this;
    containerEl.empty();
    containerEl.createEl('h2', { text: 'Office 预览设置' });

    new Setting(containerEl)
      .setName('预览模式')
      .setDesc('选择文件预览的显示方式')
      .addDropdown(dropdown => dropdown
        .addOption('right', '右侧预览')
        .addOption('split', '分屏预览')
        .addOption('tab', '新标签页')
        .setValue(this.plugin.settings.defaultMode)
        .onChange(async (value) => {
          this.plugin.settings.defaultMode = value;
          await this.plugin.saveSettings();
        }));
  }
}
