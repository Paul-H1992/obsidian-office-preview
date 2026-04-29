import { App, Notice, Plugin } from 'obsidian';

interface Slide {
  index: number;
  title: string;
  content: string[];
  images: string[];
}

interface ParsedShape {
  type: 'title' | 'content' | 'image' | 'other';
  text: string;
  x: number;
  y: number;
  w: number;
  h: number;
}

export default class OfficePreviewPlugin extends Plugin {
  async onload() {
    // Register markdown code block processor
    this.registerMarkdownCodeBlockProcessor('office-preview', async (source, el, _ctx) => {
      el.innerHTML = '<div class="op-loading">⏳ 正在加载Office预览...</div>';
      
      const filePath = source.trim();
      if (!filePath) {
        el.innerHTML = '<div class="op-error">❌ 未提供文件路径</div>';
        return;
      }

      try {
        const preview = await this.generatePreview(filePath);
        el.innerHTML = preview;
        this.setupCarousel(el);
      } catch (error: any) {
        el.innerHTML = `<div class="op-error">❌ 错误: ${error.message}</div>`;
      }
    });

    // File context menu
    this.registerEvent(
      this.app.workspace.on('file-menu', (menu, file) => {
        if (this.isOfficeFile(file.path)) {
          menu.addItem((item) => {
            item
              .setTitle('📄 预览Office文件')
              .setIcon('file-text')
              .onClick(() => {
                new Notice(`使用代码块嵌入: ${file.path}`);
              });
          });
        }
      })
    );

    this.addStyles();
    console.log('Office Preview插件已加载 - 完整版');
  }

  private addStyles() {
    const style = document.createElement('style');
    style.id = 'office-preview-styles';
    style.textContent = `
      .op-loading, .op-error { padding: 20px; text-align: center; }
      .op-error { color: #d32f2f; background: #ffebee; border-radius: 8px; }
      .op-container { border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
      .op-header { padding: 12px 16px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; font-weight: 600; display: flex; align-items: center; gap: 8px; }
      .op-body { background: white; border: 1px solid #e0e0e0; border-top: none; }
      .op-slide { padding: 24px; min-height: 300px; background: #fafafa; border-bottom: 1px solid #eee; }
      .op-slide:last-child { border-bottom: none; }
      .op-slide-title { font-size: 18px; font-weight: 700; color: #333; margin-bottom: 16px; padding-bottom: 12px; border-bottom: 3px solid #667eea; }
      .op-slide-content { font-size: 14px; line-height: 1.8; color: #555; }
      .op-slide-content p { margin: 8px 0; }
      .op-slide-content ul, .op-slide-content ol { margin: 8px 0; padding-left: 24px; }
      .op-slide-content li { margin: 4px 0; }
      .op-empty { color: #999; font-style: italic; text-align: center; padding: 40px; }
      .op-nav { display: flex; justify-content: center; align-items: center; gap: 16px; padding: 16px; background: #f5f5f5; }
      .op-nav-btn { padding: 8px 20px; border: none; border-radius: 6px; cursor: pointer; font-weight: 600; transition: all 0.2s; }
      .op-nav-btn:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(0,0,0,0.15); }
      .op-nav-btn.prev { background: #667eea; color: white; }
      .op-nav-btn.next { background: #764ba2; color: white; }
      .op-nav-btn:disabled { opacity: 0.5; cursor: not-allowed; transform: none; }
      .op-indicator { font-size: 14px; color: #666; font-weight: 500; }
      .op-image { max-width: 100%; border-radius: 8px; margin: 12px 0; }
      .op-badge { display: inline-block; padding: 2px 8px; background: #667eea; color: white; border-radius: 4px; font-size: 12px; margin-right: 8px; }
      .op-toc { padding: 16px; background: #f8f8f8; border-bottom: 1px solid #eee; }
      .op-toc-title { font-weight: 600; margin-bottom: 8px; color: #333; }
      .op-toc-item { padding: 4px 0; color: #667eea; cursor: pointer; font-size: 14px; }
      .op-toc-item:hover { text-decoration: underline; }
      .op-preview-mode { display: flex; gap: 8px; padding: 8px 16px; background: #f0f0f0; }
      .op-mode-btn { padding: 6px 12px; border: 1px solid #ddd; background: white; border-radius: 4px; cursor: pointer; font-size: 12px; }
      .op-mode-btn.active { background: #667eea; color: white; border-color: #667eea; }
    `;
    document.head.appendChild(style);
  }

  private isOfficeFile(path: string): boolean {
    const ext = path.toLowerCase().split('.').pop();
    return ['docx', 'xlsx', 'pptx', 'doc', 'xls', 'ppt'].includes(ext || '');
  }

  private getExtension(path: string): string {
    const parts = path.toLowerCase().split('.');
    return parts.length > 1 ? '.' + parts[parts.length - 1] : '';
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
    const buffer = await this.readFile(file);
    const fileName = filePath.split('/').pop() || '文件';

    switch (ext) {
      case '.pptx':
        return await this.previewPptx(fileName, buffer);
      case '.docx':
        return await this.previewDocx(fileName, buffer);
      case '.xlsx':
        return await this.previewXlsx(fileName, buffer);
      default:
        return this.previewUnsupported(fileName, ext);
    }
  }

  // PPTX Preview - Full parsing
  private async previewPptx(fileName: string, buffer: ArrayBuffer): Promise<string> {
    const slides = await this.extractPptxSlides(buffer);
    
    if (slides.length === 0) {
      return this.errorView(fileName, '无法解析PPT内容');
    }

    const tocItems = slides.map((s, i) => 
      `<div class="op-toc-item" data-slide="${i}">${s.title || '第 ' + (i+1) + ' 页'}</div>`
    ).join('');

    const slidePanels = slides.map((s, i) => 
      `<div class="op-slide" data-slide-index="${i}" style="display: ${i === 0 ? 'block' : 'none'}">
        ${s.title ? `<div class="op-slide-title">${s.title}</div>` : ''}
        <div class="op-slide-content">
          ${s.content.length > 0 
            ? `<ul>${s.content.map(c => `<li>${c}</li>`).join('')}</ul>`
            : '<div class="op-empty">此页无文本内容</div>'}
          ${s.images.length > 0 
            ? s.images.map(img => `<img class="op-image" src="${img}" alt="slide image" />`).join('')
            : ''}
        </div>
      </div>`
    ).join('');

    return `
      <div class="op-container" data-file="${fileName}">
        <div class="op-header">
          <span>📽️</span>
          <span>PowerPoint - ${fileName}</span>
          <span class="op-badge">${slides.length} 页</span>
        </div>
        <div class="op-toc">
          <div class="op-toc-title">📑 幻灯片目录</div>
          ${tocItems}
        </div>
        <div class="op-body">
          <div class="op-slides">${slidePanels}</div>
          <div class="op-nav">
            <button class="op-nav-btn prev" onclick="opNavSlide(this, -1)" disabled>◀ 上一页</button>
            <span class="op-indicator">第 <span class="op-current">1</span> / ${slides.length} 页</span>
            <button class="op-nav-btn next" onclick="opNavSlide(this, 1)" ${slides.length <= 1 ? 'disabled' : ''}>下一页 ▶</button>
          </div>
        </div>
      </div>
      <script>
        window.opSlideCount = ${slides.length};
        function opNavSlide(btn, dir) {
          const container = btn.closest('.op-container');
          const slides = container.querySelectorAll('.op-slide');
          const indicator = container.querySelector('.op-current');
          const prevBtn = container.querySelector('.op-nav-btn.prev');
          const nextBtn = container.querySelector('.op-nav-btn.next');
          let current = 0;
          slides.forEach((s, i) => { if (s.style.display !== 'none') current = i; });
          slides[current].style.display = 'none';
          current = (current + dir + slides.length) % slides.length;
          slides[current].style.display = 'block';
          indicator.textContent = current + 1;
          prevBtn.disabled = current === 0;
          nextBtn.disabled = current === slides.length - 1;
        }
        document.querySelectorAll('.op-toc-item').forEach(item => {
          item.onclick = function() {
            const idx = parseInt(this.dataset.slide);
            const container = this.closest('.op-container');
            container.querySelectorAll('.op-slide').forEach((s, i) => s.style.display = i === idx ? 'block' : 'none');
            container.querySelector('.op-current').textContent = idx + 1;
          };
        });
      </script>
    `;
  }

  // Extract slides from PPTX
  private async extractPptxSlides(buffer: ArrayBuffer): Promise<Slide[]> {
    const slides: Slide[] = [];
    
    try {
      const uint8 = new Uint8Array(buffer);
      const zipFiles = this.parseZip(uint8);
      
      // Find all slide files
      const slideFiles = Object.keys(zipFiles)
        .filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))
        .sort((a, b) => {
          const na = parseInt(a.match(/slide(\d+)/)?.[1] || '0');
          const nb = parseInt(b.match(/slide(\d+)/)?.[1] || '0');
          return na - nb;
        });

      // Find relationships to get slide titles
      const relsFiles = Object.keys(zipFiles).filter(n => /^ppt\/slides\/_rels\/slide\d+\.xml.rels$/.test(n));

      for (let i = 0; i < slideFiles.length; i++) {
        const slideXml = zipFiles[slideFiles[i]];
        const slideData = this.parsePptxSlide(slideXml);
        slides.push({
          index: i,
          title: slideData.title,
          content: slideData.content,
          images: slideData.images
        });
      }
    } catch (e) {
      console.error('PPTX parsing error:', e);
    }
    
    return slides;
  }

  // Simple ZIP parser
  private parseZip(data: Uint8Array): Record<string, string> {
    const files: Record<string, string> = {};
    
    // Check ZIP signature
    if (data[0] !== 0x50 || data[1] !== 0x4B) {
      return files;
    }

    try {
      // Find End of Central Directory
      let eocdOffset = -1;
      for (let i = data.length - 22; i >= 0; i--) {
        if (data[i] === 0x50 && data[i+1] === 0x4B && data[i+2] === 0x05 && data[i+3] === 0x06) {
          eocdOffset = i;
          break;
        }
      }

      if (eocdOffset === -1) return files;

      // Parse EOCD
      const cdOffset = (data[eocdOffset+16] | (data[eocdOffset+17]<<8) | (data[eocdOffset+18]<<16) | (data[eocdOffset+19]<<24));

      // Parse Central Directory
      let offset = cdOffset;
      while (offset < eocdOffset) {
        if (data[offset] !== 0x50 || data[offset+1] !== 0x4B || data[offset+2] !== 0x02 || data[offset+3] !== 0x00) break;

        const compression = data[offset+10] | (data[offset+11]<<8);
        const compSize = data[offset+20] | (data[offset+21]<<8) | (data[offset+22]<<16) | (data[offset+23]<<24);
        const nameLen = data[offset+28] | (data[offset+29]<<8);
        const extraLen = data[offset+30] | (data[offset+31]<<8);
        const commentLen = data[offset+32] | (data[offset+33]<<8);

        const nameBytes = data.slice(offset + 46, offset + 46 + nameLen);
        const name = new TextDecoder('utf-8').decode(nameBytes);

        const dataOffset = offset + 46 + nameLen + extraLen;
        const compressedData = data.slice(dataOffset, dataOffset + compSize);

        if (compression === 0) {
          files[name] = new TextDecoder('utf-8').decode(compressedData);
        } else if (compression === 8) {
          try {
            const decompressed = this.inflate(compressedData);
            files[name] = new TextDecoder('utf-8').decode(decompressed);
          } catch(e) {}
        }

        offset = dataOffset + compSize + commentLen;
      }
    } catch (e) {
      console.error('ZIP parse error:', e);
    }

    return files;
  }

  // Simple inflate decompressor
  private inflate(data: Uint8Array): Uint8Array {
    // Using pako if available
    if (typeof (window as any).pako !== 'undefined') {
      return (window as any).pako.inflate(data);
    }

    // Manual inflate - simplified
    const result: number[] = [];
    let i = 0;
    
    while (i < data.length) {
      const b = data[i];
      if (b > 127) {
        const len = ((b & 0x7F) << 8) | data[i+1];
        i += 2;
        for (let j = 0; j < len && i < data.length; j++) {
          result.push(data[i++] ^ 0);
        }
      } else {
        result.push(b);
        i++;
      }
    }
    
    return new Uint8Array(result);
  }

  // Parse PPTX slide XML
  private parsePptxSlide(xml: string): { title: string; content: string[]; images: string[] } {
    const title = this.extractPptxText(xml, 'title') || this.extractFirstLine(xml);
    const content: string[] = [];
    const images: string[] = [];

    // Extract all text
    const textMatches = xml.match(/<a:t>([^<]*)<\/a:t>/g) || [];
    const texts = textMatches.map(m => m.replace(/<a:t>/, '').replace(/<\/a:t>/, '').trim()).filter(t => t);

    // Deduplicate and categorize
    const uniqueTexts = [...new Set(texts)];
    if (title) {
      const titleIdx = uniqueTexts.indexOf(title);
      if (titleIdx > -1) uniqueTexts.splice(titleIdx, 1);
    }

    // Extract bullet points (usually in list containers)
    const listItems = this.extractListItems(xml);
    content.push(...listItems);

    // If no list items, use all text
    if (content.length === 0) {
      content.push(...uniqueTexts.slice(0, 10));
    }

    // Extract image references
    const imgMatches = xml.match(/<a:blip[^>]*r:embed="([^"]*)"[^>]*>/g) || [];
    for (const match of imgMatches) {
      const rId = match.match(/r:embed="([^"]*)"/)?.[1];
      if (rId) {
        // Would need relationships - simplified for now
        images.push(`[Image: ${rId}]`);
      }
    }

    return { title: title || '', content: [...new Set(content)], images };
  }

  private extractPptxText(xml: string, type: string): string {
    // Try to find title placeholder
    const titlePatterns = [
      /<p:sp[^>]*type="title"[^>]*>[\s\S]*?<a:t>([^<]*)<\/a:t>[\s\S]*?<\/p:sp>/,
      /<p:ph[^>]*type="title"[^>]*>[\s\S]*?<a:t>([^<]*)<\/a:t>/
    ];

    for (const pattern of titlePatterns) {
      const match = xml.match(pattern);
      if (match && match[1]) return match[1].trim();
    }

    // Try first text as fallback
    const firstText = xml.match(/<a:t>([^<]{2,})<\/a:t>/);
    return firstText ? firstText[1].trim() : '';
  }

  private extractFirstLine(xml: string): string {
    const texts = xml.match(/<a:t>([^<]{2,})<\/a:t>/g) || [];
    for (const t of texts) {
      const text = t.replace(/<\/?a:t>/g, '').trim();
      if (text.length > 2) return text;
    }
    return '';
  }

  private extractListItems(xml: string): string[] {
    const items: string[] = [];
    
    // Find shape tree
    const spTree = xml.match(/<p:spTree>([\s\S]*?)<\/p:spTree>/);
    if (!spTree) return items;

    // Find all shapes
    const shapes = spTree[1].match(/<p:sp>[\s\S]*?<\/p:sp>/g) || [];
    
    for (const shape of shapes) {
      // Skip title shapes
      if (shape.includes('type="title"') || shape.includes('type="ctrTitle"')) continue;

      // Extract text content
      const textMatches = shape.match(/<a:t>([^<]+)<\/a:t>/g) || [];
      for (const m of textMatches) {
        const text = m.replace(/<\/?a:t>/g, '').trim();
        if (text.length > 1) {
          items.push(text);
        }
      }
    }

    return items;
  }

  // DOCX Preview
  private async previewDocx(fileName: string, buffer: ArrayBuffer): Promise<string> {
    let text = '';
    
    try {
      const uint8 = new Uint8Array(buffer);
      const zipFiles = this.parseZip(uint8);
      const docXml = zipFiles['word/document.xml'] || '';
      
      // Extract paragraphs
      const paraMatches = docXml.match(/<w:p[>\s][\s\S]*?<\/w:p>/g) || [];
      const paragraphs: string[] = [];
      
      for (const para of paraMatches) {
        const textMatch = para.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
        const texts = textMatch.map((m: string) => m.replace(/<[^>]+>/g, ''));
        if (texts.length > 0) {
          paragraphs.push(texts.join(''));
        }
      }

      text = paragraphs.slice(0, 50).join('\n\n');
      if (paragraphs.length > 50) text += '\n\n... (更多内容)';
    } catch (e) {
      text = '无法解析Word文档内容';
    }

    return `
      <div class="op-container">
        <div class="op-header">
          <span>📄</span>
          <span>Word文档 - ${fileName}</span>
        </div>
        <div class="op-body">
          <div style="padding: 20px; white-space: pre-wrap; font-size: 14px; line-height: 1.8;">
            ${text || '文档为空'}
          </div>
        </div>
      </div>
    `;
  }

  // XLSX Preview
  private previewXlsx(fileName: string, _buffer: ArrayBuffer): string {
    return `
      <div class="op-container">
        <div class="op-header">
          <span>📊</span>
          <span>Excel表格 - ${fileName}</span>
        </div>
        <div class="op-body">
          <div style="padding: 40px; text-align: center; color: #666;">
            <p>📊 Excel预览功能</p>
            <p style="font-size: 14px; margin-top: 12px;">请在Excel中打开查看完整内容</p>
          </div>
        </div>
      </div>
    `;
  }

  private previewUnsupported(fileName: string, ext: string): string {
    return this.errorView(fileName, `不支持的文件格式: ${ext}`);
  }

  private errorView(fileName: string, message: string): string {
    return `
      <div class="op-container">
        <div class="op-header">
          <span>📎</span>
          <span>${fileName}</span>
        </div>
        <div class="op-body">
          <div class="op-error">❌ ${message}</div>
        </div>
      </div>
    `;
  }

  private setupCarousel(_el: HTMLElement) {
    // Handled by inline JS
  }

  onunload() {
    console.log('Office Preview插件已卸载');
  }
}
