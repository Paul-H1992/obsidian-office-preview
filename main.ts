import { ItemView, Notice, Plugin, WorkspaceLeaf, setIcon } from "obsidian";

const VIEW_TYPE = "pptx-preview";
const SLIDE_WIDTH = 960;
const MIN_SCALE = 0.25;
const MAX_SCALE = 2.0;

interface PptxSlide {
  index: number;
  canvas: HTMLCanvasElement;
  rendered: boolean;
}

export default class PptxPreviewPlugin extends Plugin {
  private statusBarEl: HTMLElement | null = null;

  async onload() {
    this.statusBarEl = this.addStatusBarItem();
    this.statusBarEl.setText("PPTX -");

    // Register the view with .pptx extension
    this.registerView(VIEW_TYPE, (leaf) => new PptxView(leaf, this));
    this.registerExtensions(["pptx"], VIEW_TYPE);

    // Command: open current file in viewer
    this.addCommand({
      id: "open-in-pptx-viewer",
      name: "Open in PPTX Viewer",
      checkCallback: (checking) => {
        const file = this.app.workspace.getActiveFile();
        if (file?.extension === "pptx") {
          if (!checking) this.openFile(file.path);
          return true;
        }
        return false;
      }
    });

    // File context menu
    this.registerEvent(
      this.app.workspace.on("file-menu", (menu, file) => {
        if (file.extension === "pptx") {
          menu.addItem((item) => {
            item
              .setTitle("📽️ Open in PPTX Viewer")
              .setIcon("monitor-play")
              .onClick(() => this.openFile(file.path));
          });
        }
      })
    );

    // Ribbon icon
    this.addRibbonIcon("monitor-play", "PPTX Viewer", () => {
      const file = this.app.workspace.getActiveFile();
      if (file?.extension === "pptx") {
        this.openFile(file.path);
      } else {
        new Notice("Please open a .pptx file first");
      }
    });

    console.log("PPTX Preview plugin loaded");
  }

  onunload() {
    this.statusBarEl?.setText("");
  }

  async openFile(path: string) {
    const leaf = this.app.workspace.getLeaf(false);
    await leaf.setViewState({
      type: VIEW_TYPE,
      active: true,
      state: { file: path }
    });
    this.app.workspace.revealLeaf(leaf);
  }

  setStatus(text: string) {
    this.statusBarEl?.setText(text);
  }
}

class PptxView extends ItemView {
  private plugin: PptxPreviewPlugin;
  private container: HTMLElement | null = null;
  private scrollEl: HTMLElement | null = null;
  private slidesContainer: HTMLElement | null = null;
  private slides: PptxSlide[] = [];
  private currentSlide = 0;
  private totalSlides = 0;
  private scale = 1;
  private viewer: any = null;
  private loadingEl: HTMLElement | null = null;
  private currentFilePath: string | null = null;

  constructor(leaf: WorkspaceLeaf, plugin: PptxPreviewPlugin) {
    super(leaf);
    this.plugin = plugin;
  }

  getViewType() { return VIEW_TYPE; }
  getDisplayText() { return "PPTX Viewer"; }
  getIcon() { return "monitor-play"; }

  async onOpen() {
    const content = this.contentEl;
    content.empty();
    content.addClass("pptx-preview-view");

    // Build UI
    this.buildStyles();

    const wrapper = content.createDiv({ cls: "pptx-wrapper" });
    
    // Toolbar
    const toolbar = wrapper.createDiv({ cls: "pptx-toolbar" });
    const navControls = toolbar.createDiv({ cls: "pptx-nav-controls" });
    
    const prevBtn = navControls.createEl("button", { cls: "pptx-nav-btn" });
    setIcon(prevBtn, "chevron-left");
    prevBtn.onclick = () => this.navigate(-1);

    const nextBtn = navControls.createEl("button", { cls: "pptx-nav-btn" });
    setIcon(nextBtn, "chevron-right");
    nextBtn.onclick = () => this.navigate(1);

    this.container = wrapper.createDiv({ cls: "pptx-container" });
    this.scrollEl = this.container.createDiv({ cls: "pptx-scroll" });
    this.slidesContainer = this.scrollEl.createDiv({ cls: "pptx-slides" });

    // Loading indicator
    this.loadingEl = wrapper.createDiv({ cls: "pptx-loading", text: "Loading presentation..." });

    // Navigation bar
    const navBar = wrapper.createDiv({ cls: "pptx-nav-bar" });
    const prevBottom = navBar.createEl("button", { cls: "pptx-nav-btn" });
    setIcon(prevBottom, "chevron-left");
    prevBottom.onclick = () => this.navigate(-1);

    const slideInfo = navBar.createDiv({ cls: "pptx-slide-info" });
    slideInfo.id = "pptx-slide-info";

    const nextBottom = navBar.createEl("button", { cls: "pptx-nav-btn" });
    setIcon(nextBottom, "chevron-right");
    nextBottom.onclick = () => this.navigate(1);

    const zoomControls = navBar.createDiv({ cls: "pptx-zoom-controls" });
    const zoomOut = zoomControls.createEl("button", { cls: "pptx-nav-btn small" });
    setIcon(zoomOut, "minus");
    zoomOut.onclick = () => this.zoom(-0.1);

    const zoomIn = zoomControls.createEl("button", { cls: "pptx-nav-btn small" });
    setIcon(zoomIn, "plus");
    zoomIn.onclick = () => this.zoom(0.1);

    this.updateSlideInfo();
  }

  async onClose() {
    if (this.viewer) {
      try { await this.viewer.destroy?.(); } catch(e) {}
      this.viewer = null;
    }
    this.slides = [];
  }

  private buildStyles() {
    if (document.getElementById("pptx-preview-styles")) return;
    const style = document.createElement("style");
    style.id = "pptx-preview-styles";
    style.textContent = `
      .pptx-preview-view { height: 100%; background: #1e1e1e; display: flex; flex-direction: column; }
      .pptx-wrapper { flex: 1; display: flex; flex-direction: column; overflow: hidden; }
      .pptx-toolbar { display: flex; align-items: center; justify-content: center; padding: 8px; background: #252526; border-bottom: 1px solid #3c3c3c; gap: 8px; }
      .pptx-nav-controls { display: flex; gap: 4px; }
      .pptx-nav-btn { display: flex; align-items: center; justify-content: center; width: 36px; height: 36px; border: none; border-radius: 6px; background: #3c3c3c; color: #ccc; cursor: pointer; transition: all 0.15s; }
      .pptx-nav-btn:hover { background: #4a4a4a; color: #fff; }
      .pptx-nav-btn.small { width: 28px; height: 28px; }
      .pptx-container { flex: 1; overflow: hidden; display: flex; justify-content: center; background: #2d2d2d; }
      .pptx-scroll { width: 100%; height: 100%; overflow: auto; display: flex; justify-content: center; padding: 24px; }
      .pptx-slides { display: flex; flex-direction: column; align-items: center; gap: 20px; }
      .pptx-slide-canvas { border-radius: 4px; box-shadow: 0 4px 20px rgba(0,0,0,0.4); background: white; }
      .pptx-loading { position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); color: #888; font-size: 16px; }
      .pptx-nav-bar { display: flex; align-items: center; justify-content: center; padding: 8px; background: #252526; border-top: 1px solid #3c3c3c; gap: 16px; }
      .pptx-slide-info { color: #aaa; font-size: 14px; min-width: 100px; text-align: center; }
      .pptx-zoom-controls { display: flex; gap: 4px; }
    `;
    document.head.appendChild(style);
  }

  private updateSlideInfo() {
    const el = document.getElementById("pptx-slide-info");
    if (el) el.textContent = `${this.currentSlide + 1} / ${this.totalSlides}`;
    this.plugin.setStatus(`PPTX ${this.currentSlide + 1}/${this.totalSlides}`);
  }

  private navigate(dir: number) {
    const newSlide = this.currentSlide + dir;
    if (newSlide >= 0 && newSlide < this.totalSlides) {
      this.currentSlide = newSlide;
      this.scrollToSlide(newSlide);
      this.updateSlideInfo();
      this.renderSlideIfNeeded(newSlide);
    }
  }

  private zoom(delta: number) {
    this.scale = Math.min(MAX_SCALE, Math.max(MIN_SCALE, this.scale + delta));
    this.updateAllCanvases();
  }

  private scrollToSlide(index: number) {
    const slide = this.slidesContainer?.children[index] as HTMLElement;
    slide?.scrollIntoView({ behavior: "smooth", block: "center" });
  }

  private renderSlideIfNeeded(index: number) {
    const slide = this.slides[index];
    if (!slide || slide.rendered || !this.viewer) return;
    this.renderSlide(index);
  }

  private async renderSlide(index: number) {
    const slide = this.slides[index];
    if (!slide || !this.viewer) return;
    try {
      await this.viewer.renderSlide(index, slide.canvas);
      slide.rendered = true;
    } catch(e) {
      console.error("Failed to render slide", index, e);
    }
  }

  private updateAllCanvases() {
    if (!this.viewer) return;
    for (let i = 0; i < this.slides.length; i++) {
      if (this.slides[i].rendered) {
        this.renderSlide(i);
      }
    }
  }

  async onLoadFile(file: any) {
    if (!this.slidesContainer) return;
    
    // Clear old slides
    this.slides = [];
    this.slidesContainer.innerHTML = "";
    this.currentSlide = 0;
    this.totalSlides = 0;
    this.currentFilePath = file.path;

    if (this.loadingEl) this.loadingEl.removeClass("is-hidden");

    try {
      // Load pptxviewjs dynamically
      const pptxModule = await import("pptxviewjs");
      const PptxViewer = pptxModule.default || pptxModule;
      
      // Create viewer instance
      this.viewer = new PptxViewer({
        canvas: document.createElement("canvas"),
        autoExposeGlobals: true
      });

      // Read file
      const buffer = await this.app.vault.readBinary(file);
      await this.viewer.loadFile(new Uint8Array(buffer));

      // Get slide count
      this.totalSlides = this.viewer.getSlideCount?.() || 0;
      
      if (this.totalSlides === 0) {
        // Fallback: try to parse ourselves
        await this.loadWithFallback(buffer);
        return;
      }

      // Create slide canvases
      for (let i = 0; i < this.totalSlides; i++) {
        const wrapper = this.slidesContainer!.createDiv({ cls: "pptx-slide-wrapper" });
        const canvas = wrapper.createEl("canvas", { cls: "pptx-slide-canvas" });
        
        // Get slide dimensions from viewer
        const dims = this.viewer.getSlideDimensions?.(i) || { width: 960, height: 540 };
        this.sizeCanvas(canvas, dims.width, dims.height);
        
        this.slides.push({ index: i, canvas, rendered: false });
      }

      if (this.loadingEl) this.loadingEl.addClass("is-hidden");
      
      // Render first few slides
      for (let i = 0; i < Math.min(3, this.totalSlides); i++) {
        this.renderSlide(i);
      }

      this.updateSlideInfo();
      
      // Setup intersection observer for lazy loading
      this.setupObserver();

    } catch(e) {
      console.error("PPTX load error:", e);
      if (this.loadingEl) this.loadingEl.addClass("is-hidden");
      
      // Try fallback loading
      try {
        const buffer = await this.app.vault.readBinary(file);
        await this.loadWithFallback(buffer);
      } catch(e2) {
        this.slidesContainer.innerHTML = `<div style="color:#ff6b6b;padding:40px;text-align:center;">❌ 无法加载PPTX: ${e.message || e}</div>`;
      }
    }
  }

  private async loadWithFallback(buffer: ArrayBuffer) {
    // Simple fallback: just show placeholder slides
    const slides = await this.extractSlideCount(buffer);
    this.totalSlides = slides;
    
    for (let i = 0; i < slides; i++) {
      const wrapper = this.slidesContainer!.createDiv({ cls: "pptx-slide-wrapper" });
      const canvas = wrapper.createEl("canvas", { cls: "pptx-slide-canvas" });
      this.sizeCanvas(canvas, 960, 540);
      
      // Draw placeholder
      const ctx = canvas.getContext("2d");
      if (ctx) {
        ctx.fillStyle = "#333";
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.fillStyle = "#fff";
        ctx.font = "24px sans-serif";
        ctx.textAlign = "center";
        ctx.fillText(`Slide ${i + 1}`, canvas.width/2, canvas.height/2);
      }
      
      this.slides.push({ index: i, canvas, rendered: true });
    }
    
    if (this.loadingEl) this.loadingEl.addClass("is-hidden");
    this.updateSlideInfo();
  }

  private async extractSlideCount(buffer: ArrayBuffer): Promise<number> {
    try {
      const JSZip = (await import("jszip")).default;
      const zip = await JSZip.loadAsync(buffer);
      const slideFiles = Object.keys(zip.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n));
      return slideFiles.length || 1;
    } catch {
      return 1;
    }
  }

  private sizeCanvas(canvas: HTMLCanvasElement, width: number, height: number) {
    const dpr = Math.min(window.devicePixelRatio || 1, 2);
    canvas.width = Math.round(width * this.scale * dpr);
    canvas.height = Math.round(height * this.scale * dpr);
    canvas.style.width = `${Math.round(width * this.scale)}px`;
    canvas.style.height = `${Math.round(height * this.scale)}px`;
  }

  private observer: IntersectionObserver | null = null;

  private setupObserver() {
    if (this.observer) {
      this.observer.disconnect();
    }

    this.observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (entry.isIntersecting) {
            const idx = parseInt((entry.target as HTMLElement).dataset.slideIndex || "0");
            this.renderSlideIfNeeded(idx);
          }
        }
      },
      { root: this.scrollEl, threshold: 0.1 }
    );

    // Observe all slides after they're added to DOM
    requestAnimationFrame(() => {
      this.slidesContainer?.querySelectorAll(".pptx-slide-wrapper").forEach((el) => {
        this.observer?.observe(el);
      });
    });
  }

  onUnloadFile() {
    if (this.observer) {
      this.observer.disconnect();
      this.observer = null;
    }
    this.slides = [];
  }
}