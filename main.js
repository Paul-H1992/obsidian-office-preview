"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const obsidian_1 = require("obsidian");

class OfficePreviewPlugin extends obsidian_1.Plugin {
    async onload() {
        this.registerMarkdownCodeBlockProcessor('office-preview', async (source, el) => {
            el.innerHTML = '<div class="op-loading">⏳ 正在加载Office预览...</div>';
            const filePath = source.trim();
            if (!filePath) {
                el.innerHTML = '<div class="op-error">❌ 未提供文件路径</div>';
                return;
            }
            try {
                const preview = await this.generatePreview(filePath);
                el.innerHTML = preview;
            } catch (error) {
                el.innerHTML = `<div class="op-error">❌ 错误: ${error.message}</div>`;
            }
        });

        this.registerEvent(this.app.workspace.on('file-menu', (menu, file) => {
            if (this.isOfficeFile(file.path)) {
                menu.addItem((item) => {
                    item.setTitle('📄 预览Office文件').setIcon('file-text').onClick(() => {
                        new obsidian_1.Notice(`使用代码块嵌入: ${file.path}`);
                    });
                });
            }
        }));

        this.addStyles();
        console.log('Office Preview插件已加载 - 完整版');
    }

    addStyles() {
        if (document.getElementById('op-styles')) return;
        const style = document.createElement('style');
        style.id = 'op-styles';
        style.textContent = `
            .op-loading, .op-error { padding: 20px; text-align: center; }
            .op-error { color: #d32f2f; background: #ffebee; border-radius: 8px; }
            .op-container { border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1); margin: 16px 0; }
            .op-header { padding: 12px 16px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; font-weight: 600; display: flex; align-items: center; gap: 8px; }
            .op-body { background: white; border: 1px solid #e0e0e0; border-top: none; }
            .op-slide { padding: 24px; min-height: 200px; background: #fafafa; border-bottom: 1px solid #eee; }
            .op-slide:last-child { border-bottom: none; }
            .op-slide-title { font-size: 18px; font-weight: 700; color: #333; margin-bottom: 16px; padding-bottom: 12px; border-bottom: 3px solid #667eea; }
            .op-slide-content { font-size: 14px; line-height: 1.8; color: #555; }
            .op-slide-content ul { margin: 8px 0; padding-left: 24px; }
            .op-slide-content li { margin: 4px 0; }
            .op-empty { color: #999; font-style: italic; text-align: center; padding: 40px; }
            .op-nav { display: flex; justify-content: center; align-items: center; gap: 16px; padding: 16px; background: #f5f5f5; }
            .op-nav-btn { padding: 8px 20px; border: none; border-radius: 6px; cursor: pointer; font-weight: 600; transition: all 0.2s; }
            .op-nav-btn:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(0,0,0,0.15); }
            .op-nav-btn.prev { background: #667eea; color: white; }
            .op-nav-btn.next { background: #764ba2; color: white; }
            .op-nav-btn:disabled { opacity: 0.5; cursor: not-allowed; transform: none; }
            .op-indicator { font-size: 14px; color: #666; font-weight: 500; }
            .op-toc { padding: 16px; background: #f8f8f8; border-bottom: 1px solid #eee; }
            .op-toc-title { font-weight: 600; margin-bottom: 8px; color: #333; }
            .op-toc-item { padding: 4px 0; color: #667eea; cursor: pointer; font-size: 14px; }
            .op-toc-item:hover { text-decoration: underline; }
        `;
        document.head.appendChild(style);
    }

    isOfficeFile(path) {
        const ext = path.toLowerCase().split('.').pop();
        return ['docx', 'xlsx', 'pptx', 'doc', 'xls', 'ppt'].includes(ext);
    }

    getExtension(path) {
        const parts = path.toLowerCase().split('.');
        return parts.length > 1 ? '.' + parts[parts.length - 1] : '';
    }

    async readFile(file) {
        try {
            return await this.app.vault.readBinary(file);
        } catch {
            const text = await this.app.vault.read(file);
            return new TextEncoder().encode(text).buffer;
        }
    }

    async generatePreview(filePath) {
        const file = this.app.vault.getAbstractFileByPath(filePath);
        if (!file || !('extension' in file)) throw new Error('文件未找到: ' + filePath);
        const ext = this.getExtension(filePath);
        const buffer = await this.readFile(file);
        const fileName = filePath.split('/').pop() || '文件';
        if (ext === '.pptx') return this.previewPptx(fileName, buffer);
        if (ext === '.docx') return this.previewDocx(fileName, buffer);
        if (ext === '.xlsx') return this.previewXlsx(fileName);
        return this.errorView(fileName, `不支持: ${ext}`);
    }

    async previewPptx(fileName, buffer) {
        const slides = await this.extractPptxSlides(buffer);
        if (slides.length === 0) return this.errorView(fileName, '无法解析PPT');
        const tocItems = slides.map((s, i) => 
            `<div class="op-toc-item" data-slide="${i}">${s.title || '第 ' + (i+1) + ' 页'}</div>`
        ).join('');
        const slidePanels = slides.map((s, i) => 
            `<div class="op-slide" data-slide-index="${i}" style="display: ${i === 0 ? 'block' : 'none'}">
                ${s.title ? `<div class="op-slide-title">${s.title}</div>` : ''}
                <div class="op-slide-content">
                    ${s.content.length > 0 ? `<ul>${s.content.map(c => `<li>${c}</li>`).join('')}</ul>` : '<div class="op-empty">此页无文本内容</div>'}
                </div>
            </div>`
        ).join('');
        return `
            <div class="op-container" data-file="${fileName}">
                <div class="op-header">
                    <span>📽️</span><span>PowerPoint - ${fileName}</span>
                    <span style="background:rgba(255,255,255,0.2);padding:2px 8px;border-radius:4px;font-size:12px">${slides.length}页</span>
                </div>
                <div class="op-toc">
                    <div class="op-toc-title">📑 幻灯片目录</div>
                    ${tocItems}
                </div>
                <div class="op-body">
                    <div class="op-slides">${slidePanels}</div>
                    <div class="op-nav">
                        <button class="op-nav-btn prev" onclick="opNav(this,-1)" disabled>◀</button>
                        <span class="op-indicator">第 <span class="op-cur">1</span> / ${slides.length}</span>
                        <button class="op-nav-btn next" onclick="opNav(this,1)" ${slides.length <= 1 ? 'disabled' : ''}>▶</button>
                    </div>
                </div>
            </div>
            <script>
            function opNav(b,d){
                var c=b.closest('.op-container'),sl=c.querySelectorAll('.op-slide'),cur=0;
                sl.forEach(function(s,i){if(s.style.display!='none')cur=i;});
                sl[cur].style.display='none';cur=(cur+d+sl.length)%sl.length;
                sl[cur].style.display='block';
                c.querySelector('.op-cur').textContent=cur+1;
                c.querySelector('.op-nav-btn.prev').disabled=cur===0;
                c.querySelector('.op-nav-btn.next').disabled=cur===sl.length-1;
            }
            document.querySelectorAll('.op-toc-item').forEach(function(it){
                it.onclick=function(){
                    var idx=parseInt(this.dataset.slide),c=this.closest('.op-container');
                    c.querySelectorAll('.op-slide').forEach(function(s,i){s.style.display=i===idx?'block':'none';});
                    c.querySelector('.op-cur').textContent=idx+1;
                };
            });
            </script>
        `;
    }

    async extractPptxSlides(buffer) {
        const slides = [];
        try {
            const uint8 = new Uint8Array(buffer);
            const zipFiles = this.parseZip(uint8);
            const slideFiles = Object.keys(zipFiles).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort((a,b) => {
                const na = parseInt(a.match(/slide(\d+)/)[1]);
                const nb = parseInt(b.match(/slide(\d+)/)[1]);
                return na - nb;
            });
            for (const f of slideFiles) {
                const xml = zipFiles[f];
                const title = this.extractPptxTitle(xml);
                const content = this.extractPptxContent(xml);
                slides.push({ title, content });
            }
        } catch(e) { console.error(e); }
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
                if (comp === 0) files[name] = new TextDecoder('utf-8').decode(compressedData);
                else if (comp === 8) {
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

    extractPptxTitle(xml) {
        const m = xml.match(/<p:sp[^>]*type="title"[^>]*>[\s\S]*?<a:t>([^<]*)<\/a:t>/);
        if (m && m[1]) return m[1].trim();
        const first = xml.match(/<a:t>([^<]{2,})<\/a:t>/);
        return first ? first[1].trim() : '';
    }

    extractPptxContent(xml) {
        const items = [];
        const texts = xml.match(/<a:t>([^<]+)<\/a:t>/g) || [];
        const seen = new Set();
        for (const t of texts) {
            const text = t.replace(/<\/?a:t>/g, '').trim();
            if (text.length > 1 && !seen.has(text)) { seen.add(text); items.push(text); }
        }
        return items;
    }

    async previewDocx(fileName, buffer) {
        let text = '无法解析Word文档内容';
        try {
            const uint8 = new Uint8Array(buffer);
            const zipFiles = this.parseZip(uint8);
            const docXml = zipFiles['word/document.xml'] || '';
            const paras = docXml.match(/<w:p[>\s][\s\S]*?<\/w:p>/g) || [];
            const paragraphs = [];
            for (const p of paras) {
                const tm = p.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [];
                if (tm.length) paragraphs.push(tm.map(m => m.replace(/<[^>]+>/g, '')).join(''));
            }
            text = paragraphs.slice(0, 30).join('\n\n');
            if (paragraphs.length > 30) text += '\n\n... (更多内容)';
        } catch(e) {}
        return `
            <div class="op-container">
                <div class="op-header"><span>📄</span><span>Word文档 - ${fileName}</span></div>
                <div class="op-body"><div style="padding:20px;white-space:pre-wrap;font-size:14px;line-height:1.8">${text || '(空文档)'}</div></div>
            </div>
        `;
    }

    previewXlsx(fileName) {
        return `
            <div class="op-container">
                <div class="op-header"><span>📊</span><span>Excel表格 - ${fileName}</span></div>
                <div class="op-body"><div style="padding:40px;text-align:center;color:#666"><p>📊 Excel预览</p><p style="font-size:14px;margin-top:12px">请在Excel中打开查看完整内容</p></div></div>
            </div>
        `;
    }

    errorView(fileName, msg) {
        return `<div class="op-container"><div class="op-header"><span>📎</span><span>${fileName}</span></div><div class="op-body"><div class="op-error">❌ ${msg}</div></div></div>`;
    }

    onunload() { console.log('Office Preview插件已卸载'); }
}
exports.default = OfficePreviewPlugin;
