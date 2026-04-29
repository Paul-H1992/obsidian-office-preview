"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const obsidian_1 = require("obsidian");

class OfficePreviewPlugin extends obsidian_1.Plugin {
  constructor() {
    super(...arguments);
    this.currentLeaf = null;
  }

  async onload() {
    // Register event: when file is opened, if PPTX, show custom preview
    this.registerEvent(
      this.app.workspace.on('file-open', (file) => {
        if (file && (file.path.endsWith('.pptx') || file.path.endsWith('.PPTX'))) {
          this.openPptxPreview(file);
        }
      })
    );

    // File context menu - "Open as PPT Preview"
    this.registerEvent(
      this.app.workspace.on('file-menu', (menu, file) => {
        if (file.path.endsWith('.pptx') || file.path.endsWith('.PPTX')) {
          menu.addItem((item) => {
            item
              .setTitle('📽️ PPTX预览')
              .setIcon('file-text')
              .onClick(() => this.openPptxPreview(file));
          });
        }
      })
    );

    // Ribbon icon
    this.addRibbonIcon('file-text', 'PPTX预览', () => {
      const activeFile = this.app.workspace.getActiveFile();
      if (activeFile && (activeFile.path.endsWith('.pptx') || activeFile.path.endsWith('.PPTX'))) {
        this.openPptxPreview(activeFile);
      } else {
        new obsidian_1.Notice('请先打开一个 .pptx 文件');
      }
    });

    this.addStyles();
    console.log('Office Preview PPTX插件已加载');
  }

  addStyles() {
    if (document.getElementById('op-pptx-styles')) return;
    const style = document.createElement('style');
    style.id = 'op-pptx-styles';
    style.textContent = `
      .op-pptx-preview { height: 100%; display: flex; flex-direction: column; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }
      .op-pptx-header { padding: 16px 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; font-weight: 600; display: flex; justify-content: space-between; align-items: center; flex-shrink: 0; }
      .op-pptx-header-left { display: flex; align-items: center; gap: 10px; font-size: 16px; }
      .op-pptx-header-right { font-size: 13px; opacity: 0.9; }
      .op-pptx-body { flex: 1; display: flex; overflow: hidden; }
      .op-pptx-toc { width: 220px; background: #f8f8f8; border-right: 1px solid #e0e0e0; overflow-y: auto; flex-shrink: 0; padding: 12px 0; }
      .op-pptx-toc-title { padding: 8px 16px; font-weight: 600; font-size: 13px; color: #666; border-bottom: 1px solid #eee; margin-bottom: 8px; }
      .op-pptx-toc-item { padding: 10px 16px; cursor: pointer; font-size: 13px; color: #333; transition: all 0.15s; border-left: 3px solid transparent; }
      .op-pptx-toc-item:hover { background: #eee; }
      .op-pptx-toc-item.active { background: #e8ecff; border-left-color: #667eea; color: #667eea; font-weight: 500; }
      .op-pptx-content { flex: 1; padding: 32px 40px; overflow-y: auto; background: #fafafa; }
      .op-pptx-slide { display: none; }
      .op-pptx-slide.active { display: block; }
      .op-pptx-slide-title { font-size: 24px; font-weight: 700; color: #222; margin-bottom: 20px; padding-bottom: 16px; border-bottom: 3px solid #667eea; }
      .op-pptx-slide-body { font-size: 15px; line-height: 1.9; color: #444; }
      .op-pptx-slide-body ul { margin: 12px 0; padding-left: 28px; }
      .op-pptx-slide-body li { margin: 6px 0; }
      .op-pptx-empty { color: #999; font-style: italic; text-align: center; padding: 60px 20px; }
      .op-pptx-nav { padding: 14px 20px; background: white; border-top: 1px solid #e0e0e0; display: flex; justify-content: center; align-items: center; gap: 20px; flex-shrink: 0; }
      .op-pptx-nav-btn { padding: 8px 24px; border: none; border-radius: 6px; cursor: pointer; font-weight: 600; font-size: 14px; transition: all 0.2s; }
      .op-pptx-nav-btn:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(0,0,0,0.15); }
      .op-pptx-nav-btn.prev { background: #667eea; color: white; }
      .op-pptx-nav-btn.next { background: #764ba2; color: white; }
      .op-pptx-nav-btn:disabled { opacity: 0.45; cursor: not-allowed; transform: none; box-shadow: none; }
      .op-pptx-indicator { font-size: 14px; color: #666; font-weight: 500; }
      .op-pptx-loading { display: flex; align-items: center; justify-content: center; height: 100%; font-size: 16px; color: #666; }
      .op-pptx-error { padding: 40px; text-align: center; color: #d32f2f; }
    `;
    document.head.appendChild(style);
  }

  async openPptxPreview(file) {
    const leaf = this.app.workspace.createLeafInGroup(this.app.workspace.rootSplit, 'vertical');
    try {
      const buffer = await file.arrayBuffer();
      const slides = await this.extractPptxSlides(buffer);
      if (slides.length === 0) {
        leaf.view.containerEl.innerHTML = '<div class="op-pptx-error">❌ 无法解析PPT内容</div>';
        return;
      }

      const tocItems = slides.map((s, i) => 
        `<div class="op-pptx-toc-item ${i === 0 ? 'active' : ''}" data-index="${i}">${this.escapeHtml(s.title || '第 ' + (i+1) + ' 页')}</div>`
      ).join('');

      const slideContents = slides.map((s, i) => 
        `<div class="op-pptx-slide ${i === 0 ? 'active' : ''}" data-slide="${i}">
          ${s.title ? `<div class="op-pptx-slide-title">${this.escapeHtml(s.title)}</div>` : ''}
          <div class="op-pptx-slide-body">
            ${s.content.length > 0 
              ? `<ul>${s.content.map(c => `<li>${this.escapeHtml(c)}</li>`).join('')}</ul>`
              : '<div class="op-pptx-empty">此页无文本内容</div>'}
          </div>
        </div>`
      ).join('');

      leaf.view.containerEl.innerHTML = `
        <div class="op-pptx-preview">
          <div class="op-pptx-header">
            <div class="op-pptx-header-left"><span>📽️</span><span>${this.escapeHtml(file.name)}</span></div>
            <div class="op-pptx-header-right">${slides.length} 页</div>
          </div>
          <div class="op-pptx-body">
            <div class="op-pptx-toc"><div class="op-pptx-toc-title">📑 幻灯片</div>${tocItems}</div>
            <div class="op-pptx-content">${slideContents}</div>
          </div>
          <div class="op-pptx-nav">
            <button class="op-pptx-nav-btn prev" onclick="opPptxNav(-1)" ${slides.length <= 1 ? 'disabled' : ''}>◀ 上一页</button>
            <span class="op-pptx-indicator">第 <span class="op-cur">1</span> / ${slides.length} 页</span>
            <button class="op-pptx-nav-btn next" onclick="opPptxNav(1)" ${slides.length <= 1 ? 'disabled' : ''}>下一页 ▶</button>
          </div>
        </div>
        <script>
        function opPptxNav(dir) {
          var preview = document.querySelector('.op-pptx-preview');
          var cur = 0;
          preview.querySelectorAll('.op-pptx-slide').forEach(function(s, i) { if (s.classList.contains('active')) cur = i; });
          preview.querySelectorAll('.op-pptx-slide').forEach(function(s) { s.classList.remove('active'); });
          preview.querySelectorAll('.op-pptx-toc-item').forEach(function(s) { s.classList.remove('active'); });
          cur = (cur + dir + ${slides.length}) % ${slides.length};
          preview.querySelectorAll('.op-pptx-slide')[cur].classList.add('active');
          preview.querySelectorAll('.op-pptx-toc-item')[cur].classList.add('active');
          preview.querySelector('.op-cur').textContent = cur + 1;
          preview.querySelector('.op-pptx-nav-btn.prev').disabled = cur === 0;
          preview.querySelector('.op-pptx-nav-btn.next').disabled = cur === ${slides.length - 1};
        }
        document.querySelectorAll('.op-pptx-toc-item').forEach(function(it) {
          it.onclick = function() {
            var idx = parseInt(this.dataset.index);
            var preview = document.querySelector('.op-pptx-preview');
            preview.querySelectorAll('.op-pptx-slide').forEach(function(s, i) { s.classList.toggle('active', i === idx); });
            preview.querySelectorAll('.op-pptx-toc-item').forEach(function(s, i) { s.classList.toggle('active', i === idx); });
            preview.querySelector('.op-cur').textContent = idx + 1;
            preview.querySelector('.op-pptx-nav-btn.prev').disabled = idx === 0;
            preview.querySelector('.op-pptx-nav-btn.next').disabled = idx === ${slides.length - 1};
          };
        });
        </script>
      `;

      new obsidian_1.Notice(`已打开: ${file.name}`);
    } catch (e) {
      leaf.view.containerEl.innerHTML = `<div class="op-pptx-error">❌ ${e.message || e}</div>`;
    }
  }

  escapeHtml(text) {
    return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  async extractPptxSlides(buffer) {
    const slides = [];
    try {
      const uint8 = new Uint8Array(buffer);
      const zipFiles = this.parseZip(uint8);
      const slideFiles = Object.keys(zipFiles)
        .filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))
        .sort((a, b) => {
          const na = parseInt(a.match(/slide(\d+)/)?.[1] || '0');
          const nb = parseInt(b.match(/slide(\d+)/)?.[1] || '0');
          return na - nb;
        });
      for (const f of slideFiles) {
        const xml = zipFiles[f];
        const title = this.extractTitle(xml);
        const content = this.extractContent(xml);
        slides.push({ title, content });
      }
    } catch (e) { console.error('PPTX parse error:', e); }
    return slides;
  }

  parseZip(data) {
    const files = {};
    if (data[0] !== 0x50 || data[1] !== 0x4B) return files;
    try {
      let eocd = -1;
      for (let i = data.length - 22; i >= 0; i--) {
        if (data[i]===0x50 && data[i+1]===0x4B && data[i+2]===0x05 && data[i+3]===0x06) { eocd = i; break; }
      }
      if (eocd === -1) return files;
      const cdOffset = data[eocd+16] | (data[eocd+17]<<8) | (data[eocd+18]<<16) | (data[eocd+19]<<24);
      let offset = cdOffset;
      while (offset < eocd) {
        if (data[offset] !== 0x50 || data[offset+1] !== 0x4B || data[offset+2] !== 0x02) break;
        const comp = data[offset+10] | (data[offset+11]<<8);
        const compSize = data[offset+20] | (data[offset+21]<<8) | (data[offset+22]<<16) | (data[offset+23]<<24);
        const nameLen = data[offset+28] | (data[offset+29]<<8);
        const extraLen = data[offset+30] | (data[offset+31]<<8);
        const nameBytes = data.slice(offset + 46, offset + 46 + nameLen);
        const name = new TextDecoder('utf-8').decode(nameBytes);
        const dataOffset = offset + 46 + nameLen + extraLen;
        const compressedData = data.slice(dataOffset, dataOffset + compSize);
        if (comp === 0) {
          files[name] = new TextDecoder('utf-8').decode(compressedData);
        } else if (comp === 8) {
          try { files[name] = new TextDecoder('utf-8').decode(this.inflate(compressedData)); } catch(e) {}
        }
        offset = dataOffset + compSize;
      }
    } catch(e) {}
    return files;
  }

  inflate(data) {
    if (typeof window !== 'undefined' && window.pako) return window.pako.inflate(data);
    const result = [];
    let i = 0;
    while (i < data.length) {
      const b = data[i];
      if (b > 127) {
        const len = ((b & 0x7F) << 8) | data[i+1];
        i += 2;
        for (let j = 0; j < len && i < data.length; j++) result.push(data[i++] ^ 0);
      } else { result.push(b); i++; }
    }
    return new Uint8Array(result);
  }

  extractTitle(xml) {
    const m = xml.match(/<p:sp[^>]*type="title"[^>]*>[\s\S]*?<a:t>([^<]*)<\/a:t>/);
    if (m && m[1]) return m[1].trim();
    const first = xml.match(/<a:t>([^<]{2,})<\/a:t>/);
    return first ? first[1].trim() : '';
  }

  extractContent(xml) {
    const items = [];
    const seen = new Set();
    const texts = xml.match(/<a:t>([^<]+)<\/a:t>/g) || [];
    for (const t of texts) {
      const text = t.replace(/<\/?a:t>/g, '').trim();
      if (text.length > 1 && !seen.has(text)) { seen.add(text); items.push(text); }
    }
    return items;
  }

  onunload() {
    console.log('Office Preview PPTX插件已卸载');
  }
}
exports.default = OfficePreviewPlugin;