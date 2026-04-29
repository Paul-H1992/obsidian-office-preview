"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const obsidian_1 = require("obsidian");
class OfficePreviewPlugin extends obsidian_1.Plugin {
    async onload() {
        // Register markdown code block processor for office-preview
        this.registerMarkdownCodeBlockProcessor('office-preview', async (source, el, _ctx) => {
            el.innerHTML = '<div class="office-preview-loading">⏳ 加载Office预览...</div>';
            const filePath = source.trim();
            if (!filePath) {
                el.innerHTML = '<div class="office-preview-error">❌ 未提供文件路径</div>';
                return;
            }
            try {
                const preview = await this.generatePreview(filePath);
                el.innerHTML = preview;
                // Setup carousel if PPTX
                this.setupCarouselControls(el, filePath);
            }
            catch (error) {
                el.innerHTML = `<div class="office-preview-error">❌ 错误: ${error.message}</div>`;
            }
        });
        // Register file menu handler
        this.registerEvent(this.app.workspace.on('file-menu', (menu, file) => {
            if (this.isOfficeFile(file.path)) {
                menu.addItem((item) => {
                    item
                        .setTitle('预览Office文件')
                        .setIcon('file-text')
                        .onClick(() => {
                        new obsidian_1.Notice('📄 Office预览: 使用代码块嵌入 ' + file.path);
                    });
                });
            }
        }));
        // Add styles
        this.addOfficePreviewStyles();
        console.log('Office Preview插件已加载 - 支持PPT预览');
    }
    addOfficePreviewStyles() {
        const styleEl = document.createElement('style');
        styleEl.textContent = `
      .office-preview { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; }
      .office-preview-header { display: flex; align-items: center; gap: 8px; padding: 12px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border-radius: 8px 8px 0 0; }
      .office-preview-body { border: 1px solid #e0e0e0; border-top: none; border-radius: 0 0 8px 8px; padding: 16px; background: #fff; }
      .office-preview-loading { text-align: center; padding: 40px; color: #666; }
      .office-preview-error { color: #d32f2f; padding: 16px; background: #ffebee; border-radius: 8px; }
      .ppt-carousel { position: relative; }
      .ppt-slide { padding: 20px; background: #fafafa; border-radius: 8px; margin-bottom: 12px; }
      .ppt-slide-title { font-size: 16px; font-weight: 600; color: #333; margin-bottom: 12px; padding-bottom: 8px; border-bottom: 2px solid #667eea; }
      .ppt-slide-content { font-size: 14px; line-height: 1.8; color: #555; }
      .ppt-slide-content li { margin: 4px 0; }
      .ppt-controls { display: flex; justify-content: center; gap: 12px; padding: 12px; }
      .ppt-btn { padding: 8px 16px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; transition: all 0.2s; }
      .ppt-btn:hover { transform: translateY(-1px); box-shadow: 0 2px 8px rgba(0,0,0,0.15); }
      .ppt-btn-prev { background: #667eea; color: white; }
      .ppt-btn-next { background: #764ba2; color: white; }
      .ppt-indicator { text-align: center; padding: 8px; color: #666; font-size: 14px; }
      .ppt-preview-placeholder { text-align: center; padding: 40px; background: #f5f5f5; border-radius: 8px; color: #666; }
    `;
        document.head.appendChild(styleEl);
    }
    isOfficeFile(path) {
        const ext = path.toLowerCase().substring(path.lastIndexOf('.'));
        return ['.docx', '.xlsx', '.pptx', '.doc', '.xls', '.ppt'].includes(ext);
    }
    getExtension(path) {
        const lastDot = path.lastIndexOf('.');
        return lastDot > 0 ? path.substring(lastDot).toLowerCase() : '';
    }
    async readFile(file) {
        try {
            return await this.app.vault.readBinary(file);
        }
        catch {
            const text = await this.app.vault.read(file);
            return new TextEncoder().encode(text).buffer;
        }
    }
    async generatePreview(filePath) {
        const file = this.app.vault.getAbstractFileByPath(filePath);
        if (!file || !('extension' in file)) {
            throw new Error('文件未找到: ' + filePath);
        }
        const ext = this.getExtension(filePath);
        const arrayBuffer = await this.readFile(file);
        switch (ext) {
            case 'pptx':
                return await this.previewPptx(filePath, arrayBuffer);
            case 'docx':
                return await this.previewDocx(filePath, arrayBuffer);
            case 'xlsx':
                return this.previewXlsx(filePath, arrayBuffer);
            default:
                return await this.previewGeneric(ext, filePath);
        }
    }
    async previewPptx(filePath, arrayBuffer) {
        const fileName = filePath.split('/').pop() || '演示文稿';
        // Extract slides from PPTX (which is a ZIP file)
        const slides = await this.extractSlidesFromPptx(arrayBuffer);
        if (slides.length === 0) {
            return `
        <div class="office-preview">
          <div class="office-preview-header">📽️ PowerPoint 演示文稿</div>
          <div class="office-preview-body">
            <div class="ppt-preview-placeholder">
              <p>📽️ ${fileName}</p>
              <p>无法提取幻灯片内容</p>
            </div>
          </div>
        </div>
      `;
        }
        return `
      <div class="office-preview" data-pptx-file="${filePath}">
        <div class="office-preview-header">📽️ PowerPoint 演示文稿 - ${fileName}</div>
        <div class="office-preview-body">
          <div class="ppt-carousel" data-current-slide="0">
            ${slides.map((slide, i) => `
              <div class="ppt-slide" data-slide-index="${i}" style="display: ${i === 0 ? 'block' : 'none'}">
                <div class="ppt-slide-title">📑 第 ${i + 1} 页${slide.title ? ': ' + slide.title : ''}</div>
                <div class="ppt-slide-content">
                  ${slide.content.length > 0
            ? '<ul>' + slide.content.map(c => `<li>${c}</li>`).join('') + '</ul>'
            : '<p style="color:#999;font-style:italic;">此页无文本内容</p>'}
                </div>
              </div>
            `).join('')}
          </div>
          <div class="ppt-indicator">第 <span class="current-slide">1</span> / ${slides.length} 页</div>
          <div class="ppt-controls">
            <button class="ppt-btn ppt-btn-prev" onclick="window.officePreviewPrevSlide(this)">◀ 上一页</button>
            <button class="ppt-btn ppt-btn-next" onclick="window.officePreviewNextSlide(this)">下一页 ▶</button>
          </div>
        </div>
      </div>
      <script>
        window.officePreviewSlides = ${JSON.stringify(slides.length)};
        window.officePreviewNextSlide = function(btn) {
          const carousel = btn.closest('.ppt-carousel');
          const slides = carousel.querySelectorAll('.ppt-slide');
          const indicator = btn.closest('.office-preview-body').querySelector('.current-slide');
          let current = parseInt(carousel.dataset.currentSlide);
          slides[current].style.display = 'none';
          current = (current + 1) % slides.length;
          slides[current].style.display = 'block';
          carousel.dataset.currentSlide = current;
          indicator.textContent = current + 1;
        };
        window.officePreviewPrevSlide = function(btn) {
          const carousel = btn.closest('.ppt-carousel');
          const slides = carousel.querySelectorAll('.ppt-slide');
          const indicator = btn.closest('.office-preview-body').querySelector('.current-slide');
          let current = parseInt(carousel.dataset.currentSlide);
          slides[current].style.display = 'none';
          current = (current - 1 + slides.length) % slides.length;
          slides[current].style.display = 'block';
          carousel.dataset.currentSlide = current;
          indicator.textContent = current + 1;
        };
      </script>
    `;
    }
    async extractSlidesFromPptx(arrayBuffer) {
        const slides = [];
        try {
            // Convert ArrayBuffer to Uint8Array
            const uint8Array = new Uint8Array(arrayBuffer);
            // Simple ZIP parser - find files inside the ZIP
            // PPTX is a ZIP archive containing slide XML files
            const zipEntries = this.parseZip(uint8Array);
            // Find all slide files (ppt/slides/slide*.xml)
            const slideFiles = Object.keys(zipEntries)
                .filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name))
                .sort((a, b) => {
                const numA = parseInt(a.match(/slide(\d+)/)?.[1] || '0');
                const numB = parseInt(b.match(/slide(\d+)/)?.[1] || '0');
                return numA - numB;
            });
            for (let i = 0; i < slideFiles.length; i++) {
                const slideXml = zipEntries[slideFiles[i]];
                const slideData = this.parseSlideXml(slideXml);
                slides.push({
                    index: i,
                    title: slideData.title,
                    content: slideData.content
                });
            }
        }
        catch (error) {
            console.error('Error extracting slides:', error);
        }
        return slides;
    }
    parseZip(uint8Array) {
        const entries = {};
        // Check for ZIP signature
        if (uint8Array[0] !== 0x50 || uint8Array[1] !== 0x4B) {
            // Not a valid ZIP
            return entries;
        }
        let offset = 0;
        // Find central directory and local file headers
        // This is a simplified parser - handles most PPTX files
        try {
            // Scan for local file headers (PK\x03\x04)
            let i = 0;
            while (i < uint8Array.length - 30) {
                if (uint8Array[i] === 0x50 && uint8Array[i + 1] === 0x4B &&
                    uint8Array[i + 2] === 0x03 && uint8Array[i + 3] === 0x04) {
                    // Local file header found
                    const compression = uint8Array[i + 8] | (uint8Array[i + 9] << 8);
                    const compressedSize = uint8Array[i + 18] | (uint8Array[i + 19] << 8) |
                        (uint8Array[i + 20] << 16) | (uint8Array[i + 21] << 24);
                    const nameLen = uint8Array[i + 26] | (uint8Array[i + 27] << 8);
                    const extraLen = uint8Array[i + 28] | (uint8Array[i + 29] << 8);
                    const nameBytes = uint8Array.slice(i + 30, i + 30 + nameLen);
                    const fileName = new TextDecoder('utf-8').decode(nameBytes);
                    const dataStart = i + 30 + nameLen + extraLen;
                    const dataEnd = dataStart + compressedSize;
                    if (compression === 0) {
                        // Stored (no compression)
                        const data = uint8Array.slice(dataStart, dataEnd);
                        entries[fileName] = new TextDecoder('utf-8').decode(data);
                    }
                    else if (compression === 8) {
                        // Deflate - need to decompress
                        const compressedData = uint8Array.slice(dataStart, dataEnd);
                        try {
                            const decompressed = this.inflateDeflate(compressedData);
                            entries[fileName] = new TextDecoder('utf-8').decode(decompressed);
                        }
                        catch (e) {
                            // Skip failed decompressions
                        }
                    }
                    i = dataEnd;
                }
                else {
                    i++;
                }
            }
        }
        catch (error) {
            console.error('ZIP parsing error:', error);
        }
        return entries;
    }
    inflateDeflate(compressedData) {
        // Simple deflate decompressor
        // Using pako-style inflate if available, otherwise return empty
        try {
            // Check if pako is available
            if (typeof window.pako !== 'undefined') {
                return window.pako.inflate(compressedData);
            }
            // Fallback: return empty - the content won't be parsed
            return new Uint8Array(0);
        }
        catch {
            return new Uint8Array(0);
        }
    }
    parseSlideXml(xml) {
        const title = this.extractTextBetween(xml, '<p:sp>', '<p:ph type="title"', '</p:sp>') ||
            this.extractTextBetween(xml, '<p:sp>', 'type="title"', '</p:sp>');
        // Extract all text runs
        const content = [];
        const textPattern = /<a:t>([^<]*)<\/a:t>/g;
        let match;
        while ((match = textPattern.exec(xml)) !== null) {
            const text = match[1].trim();
            if (text && text.length > 0) {
                content.push(text);
            }
        }
        return { title: title || '', content };
    }
    extractTextBetween(xml, startTag, middleTag, endTag) {
        const startIdx = xml.indexOf(startTag);
        if (startIdx === -1)
            return '';
        const middleIdx = xml.indexOf(middleTag, startIdx);
        if (middleIdx === -1)
            return '';
        const endIdx = xml.indexOf(endTag, middleIdx);
        if (endIdx === -1)
            return '';
        const textBlock = xml.substring(middleIdx, endIdx);
        const textPattern = /<a:t>([^<]*)<\/a:t>/g;
        const matches = [];
        let match;
        while ((match = textPattern.exec(textBlock)) !== null) {
            matches.push(match[1]);
        }
        return matches.join(' ');
    }
    setupCarouselControls(_el, _filePath) {
        // Carousel is handled by inline JavaScript
    }
    async previewDocx(filePath, _arrayBuffer) {
        const fileName = filePath.split('/').pop() || 'Document';
        return `
      <div class="office-preview">
        <div class="office-preview-header">📄 Word 文档</div>
        <div class="office-preview-body">
          <p>📝 <strong>${fileName}</strong></p>
          <p style="color:#666;">Word 文档预览功能开发中...</p>
        </div>
      </div>
    `;
    }
    previewXlsx(filePath, _arrayBuffer) {
        const fileName = filePath.split('/').pop() || 'Spreadsheet';
        return `
      <div class="office-preview">
        <div class="office-preview-header">📊 Excel 表格</div>
        <div class="office-preview-body">
          <p>📊 <strong>${fileName}</strong></p>
          <p style="color:#666;">Excel 预览功能开发中...</p>
        </div>
      </div>
    `;
    }
    async previewGeneric(ext, filePath) {
        const fileName = filePath.split('/').pop() || '文件';
        const icon = ext === '.ppt' || ext === '.pptx' ? '📽️' : ext === '.doc' || ext === '.docx' ? '📄' : '📊';
        return `
      <div class="office-preview">
        <div class="office-preview-header">${icon} Office 文件</div>
        <div class="office-preview-body">
          <p>${icon} <strong>${fileName}</strong></p>
          <p style="color:#666;">预览功能开发中...</p>
        </div>
      </div>
    `;
    }
    onunload() {
        console.log('Office Preview插件已卸载');
    }
}
exports.default = OfficePreviewPlugin;
