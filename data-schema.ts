import { App, Plugin } from 'obsidian';

export default class OfficePreviewPlugin extends Plugin {
  private static instance: OfficePreviewPlugin | null = null;

  async onload() {
    OfficePreviewPlugin.instance = this;
    
    // Register markdown code block processor for office-preview
    this.registerMarkdownCodeBlockProcessor('office-preview', async (source, el, ctx) => {
      el.innerHTML = '<div class="office-preview-loading">⏳ 加载Office预览...</div>';
      
      const filePath = source.trim();
      if (!filePath) {
        el.innerHTML = '<div class="office-preview-error">❌ 未提供文件路径</div>';
        return;
      }

      try {
        const preview = await this.generatePreview(filePath);
        el.innerHTML = preview;
      } catch (error: any) {
        el.innerHTML = `<div class="office-preview-error">❌ 错误: ${error.message}</div>`;
      }
    });

    // Register file menu handler for Office files
    this.registerEvent(
      this.app.workspace.on('file-menu', (menu, file) => {
        if (this.isOfficeFile(file.path)) {
          menu.addItem((item) => {
            item
              .setTitle('预览Office文件')
              .setIcon('file-text')
              .onClick(() => {
                new Notice('📄 Office预览: 使用代码块嵌入 ' + file.path);
              });
          });
        }
      })
    );

    console.log('Office Preview插件已加载 | Office Preview plugin loaded');
  }

  private isOfficeFile(path: string): boolean {
    const ext = path.toLowerCase().substring(path.lastIndexOf('.'));
    const officeExts = ['.docx', '.xlsx', '.pptx', '.doc', '.xls', '.ppt'];
    return officeExts.includes(ext);
  }

  private getExtension(path: string): string {
    const lastDot = path.lastIndexOf('.');
    return lastDot > 0 ? path.substring(lastDot).toLowerCase() : '';
  }

  private async readFile(file: any): Promise<ArrayBuffer> {
    try {
      return await this.app.vault.readBinary(file);
    } catch {
      const text = await this.app.vault.read(file);
      return new TextEncoder().encode(text).buffer;
    }
  }

  private async generatePreview(filePath: string): Promise<string> {
    const file = this.app.vault.getAbstractFileByPath(filePath);
    if (!file || !('extension' in file)) {
      throw new Error('文件未找到: ' + filePath);
    }

    const ext = this.getExtension(filePath);
    const arrayBuffer = await this.readFile(file);

    switch (ext) {
      case 'docx':
        return await this.previewDocx(filePath, arrayBuffer);
      case 'xlsx':
        return this.previewXlsx(filePath, arrayBuffer);
      case 'pptx':
        return await this.previewPptx(filePath, arrayBuffer);
      case 'doc':
        return await this.previewDoc(filePath, arrayBuffer);
      case 'xls':
        return this.previewXls(filePath, arrayBuffer);
      case 'ppt':
        return this.previewPpt(filePath, arrayBuffer);
      default:
        throw new Error('不支持的文件类型: ' + ext);
    }
  }

  private async previewDocx(filePath: string, arrayBuffer: ArrayBuffer): Promise<string> {
    const fileName = filePath.split('/').pop() || 'Document';
    return `
      <div class="office-preview docx-preview">
        <div class="preview-header">
          <span class="file-icon">📄</span>
          <span class="file-type">Word 文档</span>
        </div>
        <div class="preview-content">
          <p class="preview-note">📝 ${fileName}</p>
          <p class="preview-hint">💡 提示: 在外部编辑器中打开以查看完整内容</p>
        </div>
      </div>
    `;
  }

  private previewXlsx(filePath: string, arrayBuffer: ArrayBuffer): string {
    const fileName = filePath.split('/').pop() || 'Spreadsheet';
    return `
      <div class="office-preview xlsx-preview">
        <div class="preview-header">
          <span class="file-icon">📊</span>
          <span class="file-type">Excel 表格</span>
        </div>
        <div class="preview-content">
          <p class="preview-note">📊 ${fileName}</p>
          <p class="preview-hint">💡 提示: 在外部编辑器中打开以查看完整内容</p>
        </div>
      </div>
    `;
  }

  private async previewPptx(filePath: string, arrayBuffer: ArrayBuffer): Promise<string> {
    const fileName = filePath.split('/').pop() || 'Presentation';
    return `
      <div class="office-preview pptx-preview">
        <div class="preview-header">
          <span class="file-icon">📽️</span>
          <span class="file-type">PowerPoint 演示文稿</span>
        </div>
        <div class="preview-content">
          <p class="preview-note">📽️ ${fileName}</p>
          <p class="preview-hint">💡 提示: 在外部编辑器中打开以查看完整内容</p>
        </div>
      </div>
    `;
  }

  private async previewDoc(filePath: string, arrayBuffer: ArrayBuffer): Promise<string> {
    const fileName = filePath.split('/').pop() || 'Document';
    return `
      <div class="office-preview doc-preview">
        <div class="preview-header">
          <span class="file-icon">📄</span>
          <span class="file-type">Word 文档 (旧版)</span>
        </div>
        <div class="preview-content">
          <p class="preview-note">📝 ${fileName}</p>
          <p class="preview-hint">⚠️ 提示: 建议转换为.docx格式以获得更好的兼容性</p>
        </div>
      </div>
    `;
  }

  private previewXls(filePath: string, arrayBuffer: ArrayBuffer): string {
    const fileName = filePath.split('/').pop() || 'Spreadsheet';
    return `
      <div class="office-preview xls-preview">
        <div class="preview-header">
          <span class="file-icon">📊</span>
          <span class="file-type">Excel 表格 (旧版)</span>
        </div>
        <div class="preview-content">
          <p class="preview-note">📊 ${fileName}</p>
          <p class="preview-hint">⚠️ 提示: 建议转换为.xlsx格式以获得更好的兼容性</p>
        </div>
      </div>
    `;
  }

  private previewPpt(filePath: string, arrayBuffer: ArrayBuffer): string {
    const fileName = filePath.split('/').pop() || 'Presentation';
    return `
      <div class="office-preview ppt-preview">
        <div class="preview-header">
          <span class="file-icon">📽️</span>
          <span class="file-type">PowerPoint (旧版)</span>
        </div>
        <div class="preview-content">
          <p class="preview-note">📽️ ${fileName}</p>
          <p class="preview-hint">⚠️ 提示: 建议转换为.pptx格式以获得更好的兼容性</p>
        </div>
      </div>
    `;
  }

  onunload() {
    console.log('Office Preview插件已卸载 | Office Preview plugin unloaded');
  }
}

// Re-export for external use
export { OfficePreviewPlugin };
