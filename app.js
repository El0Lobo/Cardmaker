// Business Card Maker — JS only update:
// - Sidebar horizontal scroll suppressed, compact color pickers
// - Live print preview refresh on all edits (debounced)
// - Back offset (mm) applied to the imposition preview
// - Visual back offset guide overlay in preview
// - Pattern-back still disables back editor and fills page 2

(function(){
  const $ = (sel, root=document)=>root.querySelector(sel);
  const $$ = (sel, root=document)=>Array.from(root.querySelectorAll(sel));

  // Core DOM
  const canvasFront = $("#canvasFront");
  const canvasBack  = $("#canvasBack");
  const overlayEl   = $("#overlay");
  const ctxF = canvasFront.getContext("2d");
  const ctxB = canvasBack.getContext("2d");

  // Controls
  const unitSel = $("#units");
  const widthInput = $("#cardWidth");
  const heightInput = $("#cardHeight");
  const dpiInput = $("#dpi");
  const bleedInput = $("#bleed");
  const cropMarksChk = $("#cropMarks");
  const gutterInput = $("#gutter");
  const pageMarginInput = $("#pageMargin");
  const trimBoxesChk = $("#trimBoxes");
  const showBleedChk = $("#showBleed");
  const showGridChk  = $("#showGrid");
  const gridSpacingInput = $("#gridSpacing");
  const backFlipSel = $("#backFlip");

  const btnAddText = $("#btnAddText");
  const imgInput = $("#imgInput");
  const btnDelete = $("#btnDelete");
  const btnDuplicate = $("#btnDuplicate");
  const btnBringFwd = $("#btnBringFwd");
  const btnSendBack = $("#btnSendBack");
  const btnFlipH = $("#btnFlipH");
  const btnFlipV = $("#btnFlipV");
  const rotRange = $("#rotation");
  const opacityRange = $("#opacity");
  const lockChk = $("#lock");
  const lockedLayerList = $("#lockedLayerList");
  const btnUnlockLayer = $("#btnUnlockLayer");

  const textControls = $("#textControls");
  const textContent = $("#textContent");
  const canvasTextEditor = $("#canvasTextEditor");
  const inlineEditorToolbar = $("#inlineEditorToolbar");
  const fontFamily = $("#fontFamily");
  const fontSize = $("#fontSize");
  const fontWeight = $("#fontWeight");
  const textAlignSel = $("#textAlign");
  const fillStyleInp = $("#fillStyle");
  const strokeStyleInp = $("#strokeStyle");
  const lineWidthInp = $("#lineWidth");

  const bgColorFront = $("#bgColorFront");
  const bgColorBack  = $("#bgColorBack");
  const bgImgFront   = $("#bgImgFront");
  const bgImgBack    = $("#bgImgBack");
  const backOffsetXmm = $("#backOffsetXmm");
  const backOffsetYmm = $("#backOffsetYmm");

  const thumbFront = $("#thumbFront");
  const thumbBack  = $("#thumbBack");

  const sheetPatternMode = $("#sheetPatternMode");
  const sheetPatternBox  = $("#sheetPatternBox");
  const sheetPatternImg  = $("#sheetPatternImg");
  const sheetOffsetX     = $("#sheetOffsetX");
  const sheetOffsetY     = $("#sheetOffsetY");

  const perCardBacks     = $("#perCardBacks");
  const perCardBox       = $("#perCardBox");
  const perCardBackFiles = $("#perCardBackFiles");

  const canvasTabs = $(".canvas-tabs");
  const canvasHolder = $(".canvas-holder");

  const btnExportPDF = $("#btnExportPDF");
  const btnDownloadProject = $("#btnDownloadProject");
  const btnLoadProject = $("#btnLoadProject");
  const btnResetProject = $("#btnResetProject");

  const thicknessInp = $("#thickness");
  const zoom3D = $("#zoom3D");
  const viewport = $("#viewport");
  const card3d = $("#card3d");
  const faceFront = $("#faceFront");
  const faceBack = $("#faceBack");
  const edges = [$(".edge1"), $(".edge2"), $(".edge3"), $(".edge4")];

  // --- Paper type (replaces BG Front/Back) ---
  const PAPER_FILES = {
    white:    "static/img/white.jpg",
    offwhite: "static/img/offwhite.jpg",
    brown:    "static/img/brown.jpg",
  };
  const PAPER_EDGE = {
    white:    "#e9e9e9",
    offwhite: "#e6dfcf",
    brown:    "rgb(198, 149, 99)",
  };
  let paperImages = { white:null, offwhite:null, brown:null };

  // --- Runtime UI fixes (JS-only) ---
  const sidebarInner = $(".sidebar-inner") || $(".sidebar");
  if (sidebarInner){
    sidebarInner.style.overflowX = "hidden"; // no horizontal scroll
  }
  // Compact color pickers (they can force width on some browsers)
  $$('input[type="color"]').forEach(inp=>{
    inp.style.width = "44px";
    inp.style.minWidth = "44px";
    inp.style.padding = "0";
  });

  // Dynamic elements
  let backDisabledOverlay = null;
  let printPreviewWrap = null;
  let printPreviewFrame = null;
  let paperTypeSelect = null;

  // State
  const createEmptyState = () => ({
    side: "front",
    units: "mm",
    cardWidthMM: 85,
    cardHeightMM: 55,
    dpi: 300,
    bleedMM: 3,
    cropMarks: true,
    gutterMM: 5,
    pageMarginMM: 10,
    trimBoxes: false,
    showBleed: false,
    grid: { show: false, spacingMM: 5 },
    backFlip: "none",
    backOffsetMM: {x:0,y:0}, // applied in preview on BACK page
    bg: {
      front: { color: "#ffffff", image: null },
      back:  { color: "#ffffff", image: null }
    },
    front: { layers: [] },
    back:  { layers: [] },
    selection: null,
    sheetPattern: { enabled: false, image: null, offsetMM: {x:0,y:0}, opacity: 1 },
    perCard: { enabled: false, images: [] },
    customFonts: [],
    paperType: "white"
  });
  let state = createEmptyState();
  let selectionOverlayRect = null;
  let inlineEditingLayerId = null;
  let tinyOverlayEditor = null;
  let tinyOverlayInitPromise = null;
  const FONT_SIZE_OPTIONS = [4,6,8,10,12,14,16,18,24,32];
  const inlineEditorIsVisible = ()=>!!(inlineEditingLayerId && canvasTextEditor && canvasTextEditor.classList.contains("active"));
// === Inject Sheet Pattern Opacity control (0..1) ===
(function injectPatternOpacityControl(){
  if (!sheetPatternBox) return;
  const row = document.createElement("label");
  row.style.display = "grid";
  row.style.gridTemplateColumns = "auto 1fr 48px";
  row.style.alignItems = "center";
  row.style.gap = "8px";
  row.innerHTML = `
    <span>Pattern Opacity</span>
    <input id="sheetPatternOpacity" type="range" min="0" max="1" step="0.01" value="${state.sheetPattern.opacity}">
    <input id="sheetPatternOpacityNum" type="number" min="0" max="1" step="0.01" value="${state.sheetPattern.opacity}" style="width:60px">
  `;
  sheetPatternBox.appendChild(row);

  const r = row.querySelector("#sheetPatternOpacity");
  const n = row.querySelector("#sheetPatternOpacityNum");
  function setOpacity(v){
    const val = Math.max(0, Math.min(1, parseFloat(v)||0));
    state.sheetPattern.opacity = val;
    r.value = String(val);
    n.value = String(val);
    refreshPrintPreview();
    update3DTextures(true);
  }
  r.addEventListener("input", ()=> setOpacity(r.value));
  n.addEventListener("input", ()=> setOpacity(n.value));
})();

  // Helpers
  const mmToPx = (mm)=> (state.dpi/25.4) * mm;
  const pxToMm = (px)=> (25.4/state.dpi) * px;
  const clamp = (v, a, b)=>Math.max(a, Math.min(b, v));
  const debounce = (fn, ms=120)=>{
    let t; return (...args)=>{ clearTimeout(t); t=setTimeout(()=>fn(...args), ms); };
  };

  function getActiveCanvas(){ return state.side==="front" ? canvasFront : canvasBack; }
  function getActiveCtx(){ return state.side==="front" ? ctxF : ctxB; }
  function getActivePage(){ return state.side==="front" ? state.front : state.back; }
  function cloneLayer(l){ const c={...l}; if(c.type==="image") c.image=l.image; return c; }

  // ---- Fonts: standard list + helpers (top-level; used by multiple places) ----
  const STANDARD_FONTS = [
    ["'OlivettiLettera22','Space Mono','IBM Plex Mono',monospace","Olivetti Lettera 22"],
    ["Inter,system-ui,Arial,sans-serif","Inter / System Sans"],
    ["Arial,Helvetica,sans-serif","Arial"],
    ["'Helvetica Neue',Helvetica,Arial,sans-serif","Helvetica Neue"],
    ["Verdana,Geneva,sans-serif","Verdana"],
    ["Tahoma,Geneva,sans-serif","Tahoma"],
    ["'Trebuchet MS',Helvetica,sans-serif","Trebuchet MS"],
    ["Georgia,serif","Georgia"],
    ["'Times New Roman',Times,serif","Times New Roman"],
    ["Palatino Linotype,Book Antiqua,Palatino,serif","Palatino"],
    ["'Courier New',Courier,monospace","Courier New"],
    ["'Lucida Console',Monaco,monospace","Lucida Console"],
    ["Impact,Charcoal,sans-serif","Impact"],
    ["'Comic Sans MS',cursive","Comic Sans MS"],
    ["'Gill Sans','Gill Sans MT',Calibri,sans-serif","Gill Sans"],
    ["'Franklin Gothic Medium','Arial Narrow Bold',sans-serif","Franklin Gothic"]
  ];

  const FONT_ALIAS_MAP = {
    zai_OlivettiLettera22Typewriter: "'OlivettiLettera22','Space Mono','IBM Plex Mono',monospace"
  };

  function ensureStandardFontsInSelect(selectEl){
    if (!selectEl) return;
    const existing = new Set([...selectEl.options].map(o=>o.value));
    STANDARD_FONTS.forEach(([val,label])=>{
      if (!existing.has(val)){
        const opt = document.createElement("option");
        opt.value = val; opt.textContent = label;
        selectEl.appendChild(opt);
      }
    });
  }

  function normalizeFontAliases(targetState = state){
    ["front","back"].forEach(side=>{
      const layers = targetState?.[side]?.layers || [];
      layers.forEach(layer=>{
        if (layer?.type !== "text") return;
        const raw = (layer.fontFamily || "").trim();
        const key = raw.replace(/^['"]+|['"]+$/g,"");
        if (FONT_ALIAS_MAP[key]){
          layer.fontFamily = FONT_ALIAS_MAP[key];
        }
      });
    });
  }

  function renderState(clearInputs=false){
    initUIFromState();
    if (sheetPatternMode){
      sheetPatternMode.checked = !!(state.sheetPattern?.enabled);
      if (sheetPatternBox){
        sheetPatternBox.classList.toggle("hidden", !state.sheetPattern.enabled);
      }
    }
    if (typeof sheetOffsetX?.value !== "undefined"){
      sheetOffsetX.value = state.sheetPattern?.offsetMM?.x || 0;
    }
    if (typeof sheetOffsetY?.value !== "undefined"){
      sheetOffsetY.value = state.sheetPattern?.offsetMM?.y || 0;
    }
    if (perCardBacks){
      perCardBacks.checked = !!state.perCard?.enabled;
      if (perCardBox){
        perCardBox.classList.toggle("hidden", !state.perCard.enabled);
      }
    }
    if (clearInputs){
      if (sheetPatternImg) sheetPatternImg.value = "";
      if (perCardBackFiles) perCardBackFiles.value = "";
    }
    const patRange = document.getElementById("sheetPatternOpacity");
    const patNum = document.getElementById("sheetPatternOpacityNum");
    if (patRange) patRange.value = state.sheetPattern?.opacity ?? 1;
    if (patNum) patNum.value = state.sheetPattern?.opacity ?? 1;
    resizeCanvasToCard();
    setSide("front");
    redrawAll();
    drawOverlay();
    refreshPrintPreview();
    update3DTextures(true);
    applyPaperType(state.paperType || "white");
  }

  function updateSelectedTextFont(ff){
    if (fontFamily) fontFamily.value = ff;
    const sel = getSelection();
    if (sel && sel.type === "text"){
      sel.fontFamily = ff;
      updateTextControls(sel);
      redrawAll();
    }
  }

  // sizing
  function resizeCanvasToCard(){
    const wpx = Math.round(mmToPx(state.cardWidthMM));
    const hpx = Math.round(mmToPx(state.cardHeightMM));
    [canvasFront,canvasBack].forEach(cv=>{ cv.width = wpx; cv.height = hpx; });
    syncThumbs();
    redrawAll();
    update3DTextures(false);
    refreshPrintPreview();
  }

  function setSide(side){
    state.side = side;
    canvasFront.classList.toggle("hidden", side!=="front");
    canvasBack.classList.toggle("hidden", side!=="back");
    updateBackEditorDisabled();
    updateInspector();
    drawOverlay();
    redrawAll();
    renderCanvasTabs();
  }

  // Layer factory
  function makeTextLayer(){
    return {
      id: crypto.randomUUID(),
      type: "text",
      text: "Your Name",
      fontFamily: fontFamily.value || "Inter,system-ui,Arial,sans-serif",
      fontWeight: fontWeight.value || "400",
      fontStyle: "normal",
      fontSize: parseFloat(fontSize.value||"10"),
      align: textAlignSel.value || "center",
      fillStyle: fillStyleInp.value || "#111111",
      strokeStyle: strokeStyleInp.value || "#000000",
      lineWidth: parseFloat(lineWidthInp.value||"0"),
      x: getActiveCanvas().width/2,
      y: getActiveCanvas().height/2,
      scaleX: 1, scaleY: 1,
      rotation: 0,
      opacity: 1,
      flipX: false, flipY: false,
      locked: false
    };
  }
  function makeImageLayer(img){
    const cv = getActiveCanvas();
    const scale = Math.min(cv.width*0.6/img.width, cv.height*0.6/img.height);
    return {
      id: crypto.randomUUID(),
      type: "image",
      image: img,
      x: cv.width/2,
      y: cv.height/2,
      scaleX: scale, scaleY: scale,
      rotation: 0,
      opacity: 1,
      flipX: false, flipY: false,
      locked: false
    };
  }
  function preloadPaperTextures(cb){
    let left = Object.keys(PAPER_FILES).length;
    Object.entries(PAPER_FILES).forEach(([key, url])=>{
      const img = new Image();
      img.onload = ()=>{ paperImages[key]=img; if(--left===0 && cb) cb(); };
      img.onerror = ()=>{ paperImages[key]=null; if(--left===0 && cb) cb(); };
      img.src = url;
    });
  }
  function ensurePaperOverlays(){
    [faceFront, faceBack].forEach(face=>{
      let ov = face.querySelector('.paper-ov');
      if (!ov){
        ov = document.createElement('div');
        ov.className = 'paper-ov';
        ov.style.cssText = `
          position:absolute; inset:0;
          background-size:cover; background-position:center;
          opacity:1;          /* adjust texture strength here */
          mix-blend-mode:multiply;
          pointer-events:none;
        `;
        face.appendChild(ov);
      }
    });
  }

  // Inject "Paper Type" selector UI and hide old BG controls
function installPaperSelector(){
  // Put the selector right under "3D Preview" section
  const h2list = Array.from(document.querySelectorAll("h2"));
  const threeDHeader = h2list.find(h => h.textContent.trim().toLowerCase() === "3d preview");
  const host = threeDHeader ? threeDHeader.parentElement : document.querySelector(".sidebar-inner");

  const wrap = document.createElement("div");
  wrap.className = "grid-1";
  wrap.innerHTML = `
    <h3>Paper (3D texture only)</h3>
    <label>Paper Type
      <select id="paperType">
        <option value="white">White</option>
        <option value="offwhite">Off-white</option>
        <option value="brown">Brown</option>
      </select>
    </label>
    <p class="muted">Affects only the 3D preview texture and edge tint. Your canvas backgrounds still come from the Backgrounds section.</p>
  `;
  if (host) host.appendChild(wrap);

  // IMPORTANT: do NOT hide background controls anymore
  // (bgColorFront, bgColorBack, bgImgFront, bgImgBack remain visible & bound)

  paperTypeSelect = wrap.querySelector("#paperType");
  paperTypeSelect.value = state.paperType || "white";
  paperTypeSelect.addEventListener("change", ()=> applyPaperType(paperTypeSelect.value));
}


  function applyPaperType(kind){
    state.paperType = kind;
    if (paperTypeSelect && paperTypeSelect.value !== kind){
      paperTypeSelect.value = kind;
    }

    // Update 3D face overlays
    ensurePaperOverlays();
    const url = PAPER_FILES[kind];
    [faceFront, faceBack].forEach(face=>{
      const ov = face.querySelector('.paper-ov');
      if (ov) ov.style.backgroundImage = `url(${url})`;
    });

    // Tint the 3D edges
    const edgeCol = PAPER_EDGE[kind] || "#e9e9e9";
    edges.forEach(e => { if (e) e.style.background = edgeCol; });

    // Re-apply 3D transforms/sizes so everything stays aligned
    update3DTextures(true);

    // IMPORTANT: do NOT touch state.bg.* or redraw canvases here
    // -> print preview and export remain unaffected
  }

  // Drawing
  function clearAndBg(ctx, side){
    const cv = ctx.canvas;
    ctx.save();
    ctx.clearRect(0,0,cv.width,cv.height);
    ctx.fillStyle = state.bg[side].color || "#ffffff";
    ctx.fillRect(0,0,cv.width,cv.height);
    const bg = state.bg[side].image;
    if (bg){
      ctx.save();
      if (side==="back"){
        ctx.translate(mmToPx(state.backOffsetMM.x), mmToPx(state.backOffsetMM.y)); // design-time preview too
      }
      const scale = Math.max(cv.width/bg.width, cv.height/bg.height);
      const w = bg.width*scale, h = bg.height*scale;
      ctx.drawImage(bg, (cv.width-w)/2, (cv.height-h)/2, w, h);
      ctx.restore();
    }
    ctx.restore();
  }

  function measureLayerBox(layer, ctx){
    if (layer.type === "image"){
      const w = layer.image?.width || 0;
      const h = layer.image?.height || 0;
      return {
        width: Math.max(20, w),
        height: Math.max(20, h),
        lineHeight: 0,
        maxLineWidth: w,
        anchorShiftX: 0,
        fontPx: 0
      };
    }
    const fontPx = Math.round(mmToPx(layer.fontSize));
    const px = Math.max(8, fontPx);
    const fontStyle = layer.fontStyle && layer.fontStyle!=="normal" ? `${layer.fontStyle} ` : "";
    ctx.font = `${fontStyle}${layer.fontWeight} ${px}px ${layer.fontFamily}`;
    const lines = (layer.text ?? "").split(/\r?\n/);
    const lineHeight = px * 1.3;
    const maxLineW = lines.reduce((m, ln)=>Math.max(m, ctx.measureText(ln).width), 0);
    const pad = 20;
    const width = Math.max(20, maxLineW + pad);
    const height = Math.max(20, lineHeight * lines.length);
    let anchorShiftX = 0;
    if (layer.align === "left") anchorShiftX = maxLineW/2;
    else if (layer.align === "right") anchorShiftX = -maxLineW/2;
    return { width, height, lineHeight, maxLineWidth:maxLineW, anchorShiftX, fontPx:px };
  }

  function drawLayer(ctx, layer){
    ctx.save();
  ctx.globalAlpha = layer.opacity ?? 1;
  ctx.translate(layer.x, layer.y);
  ctx.rotate(layer.rotation * Math.PI / 180);
  ctx.scale(layer.flipX ? -layer.scaleX : layer.scaleX, layer.flipY ? -layer.scaleY : layer.scaleY);

  if (layer.type === "image"){
    const img = layer.image;
    ctx.drawImage(img, -img.width/2, -img.height/2);

  } else if (layer.type === "text"){
    const fontPx = Math.round(mmToPx(layer.fontSize));
    const px = Math.max(8, fontPx);
    const fontStyle = layer.fontStyle && layer.fontStyle !== "normal" ? `${layer.fontStyle} ` : "";
    ctx.font = `${fontStyle}${layer.fontWeight} ${px}px ${layer.fontFamily}`;
    ctx.textAlign = layer.align;
    ctx.textBaseline = "middle";

    const lines = (layer.text ?? "").split(/\r?\n/);
    const lineHeight = px * 1.3;
    const totalH = lineHeight * lines.length;

    // Stroke first (if any), then fill, per-line to respect breaks
    if ((layer.lineWidth || 0) > 0){
      ctx.lineWidth = layer.lineWidth;
      ctx.strokeStyle = layer.strokeStyle || "#000";
      for (let i=0;i<lines.length;i++){
        const y = -totalH/2 + i*lineHeight + lineHeight/2;
        ctx.strokeText(lines[i], 0, y);
      }
    }
    ctx.fillStyle = layer.fillStyle || "#111";
    for (let i=0;i<lines.length;i++){
      const y = -totalH/2 + i*lineHeight + lineHeight/2;
      ctx.fillText(lines[i], 0, y);
    }
  }

  ctx.restore();
}


function drawEditorOverlays(ctx, side){
  const cv = ctx.canvas;
  if (state.showBleed){
    const b = Math.round(mmToPx(state.bleedMM));
    ctx.save();
    ctx.strokeStyle = "rgba(255,0,0,0.8)";
    ctx.setLineDash([8,6]);
    ctx.lineWidth = 2;
    ctx.strokeRect(b, b, cv.width-2*b, cv.height-2*b);
    ctx.restore();
  }
  if (state.grid.show){
    const spacingPx = mmToPx(state.grid.spacingMM);
    ctx.save();
    ctx.strokeStyle = "rgba(110,168,254,0.25)";
    ctx.lineWidth = 1;
    ctx.beginPath();
    for (let x=0; x<=cv.width; x+=spacingPx){ ctx.moveTo(x,0); ctx.lineTo(x,cv.height); }
    for (let y=0; y<=cv.height; y+=spacingPx){ ctx.moveTo(0,y); ctx.lineTo(cv.width,y); }
    ctx.stroke();

    // ✅ always draw midlines
    ctx.strokeStyle = "rgba(0,160,255,0.4)";
    ctx.setLineDash([6,4]);
    ctx.beginPath();
    ctx.moveTo(cv.width/2,0); ctx.lineTo(cv.width/2,cv.height);
    ctx.moveTo(0,cv.height/2); ctx.lineTo(cv.width,cv.height/2);
    ctx.stroke();
    ctx.restore();
  }
}

// ---- Utility: convert Pickr color result to CSS string ----
function pickrColorToCss(p) {
  // p.toHEXA().toString() is fine for solid;
  // but to support gradients, we check p.toRGBA() etc.
  // Pickr gives us a color object that can output multiple formats:
  return p.toHEXA().toString();  // fallback; if using gradient plugin, this would be gradient string
}

// ---- FRONT PICKR ----
const pickrFront = Pickr.create({
  el: '#picker-front',
  theme: 'monolith',
  default: state.bg.front.color || '#ffffff',
  swatches: ['#ffffff','#000000','#ff0000','#00ff00','#0000ff','#ffff00'],
  components: {
    preview: true,
    opacity: true,
    hue: true,
    interaction: {
      hex: true,
      rgba: true,
      hsla: true,
      input: true,
      save: true
    }
  }
});

pickrFront.on('save', (color, instance)=>{
  const css = pickrColorToCss(color);
  state.bg.front.color = css;
  state.bg.front.gradient = null;  // if you want to differentiate gradients later
  redrawAll();
  debouncedPreview();
  update3DTextures(false);
});

// ---- BACK PICKR ----
const pickrBack = Pickr.create({
  el: '#picker-back',
  theme: 'monolith',
  default: state.bg.back.color || '#ffffff',
  swatches: ['#ffffff','#000000','#ff0000','#00ff00','#0000ff','#ffff00'],
  components: {
    preview: true,
    opacity: true,
    hue: true,
    interaction: {
      hex: true,
      rgba: true,
      hsla: true,
      input: true,
      save: true
    }
  }
});

pickrBack.on('save', (color, instance)=>{
  const css = pickrColorToCss(color);
  state.bg.back.color = css;
  state.bg.back.gradient = null;
  redrawAll();
  debouncedPreview();
  update3DTextures(false);
});

  function redraw(ctx, side){
    clearAndBg(ctx, side);
    const page = side==="front" ? state.front : state.back;
    page.layers.forEach(ly=>drawLayer(ctx, ly));
    drawEditorOverlays(ctx, side);
    drawSelectionOutline(side);
  }
  function redrawAll(){
    redraw(ctxF,"front");
    redraw(ctxB,"back");
    syncThumbs();
    update3DTextures(false);
    debouncedPreview(); // auto refresh preview on any draw
    refreshLockedLayerList();
  }

  // Selection
  function getSelection(){
    const page = getActivePage();
    return page.layers.find(l => l.id === state.selection);
  }
  function setSelection(id){
    if (inlineEditingLayerId && inlineEditingLayerId !== id){
      closeInlineCanvasEditor(true);
    }
    state.selection = id;
    updateInspector();
    drawOverlay();
    redrawAll();
  }
  function refreshLockedLayerList(){
    if (!lockedLayerList) return;
    const page = getActivePage();
    const locked = page.layers.filter(l=>l.locked);
    lockedLayerList.innerHTML = "";
    if (!locked.length){
      const opt = document.createElement("option");
      opt.value = "";
      opt.textContent = "None";
      lockedLayerList.appendChild(opt);
      if (btnUnlockLayer) btnUnlockLayer.disabled = true;
      return;
    }
    locked.forEach((layer, idx)=>{
      const opt = document.createElement("option");
      opt.value = layer.id;
      const label = layer.type === "text" && layer.text
        ? `Text: ${layer.text.slice(0,18)}${layer.text.length>18?"…":""}`
        : `${layer.type || "Layer"} #${idx+1}`;
      opt.textContent = label;
      lockedLayerList.appendChild(opt);
    });
    lockedLayerList.value = locked[0].id;
    if (btnUnlockLayer) btnUnlockLayer.disabled = false;
  }
  function moveSelectionBy(dx, dy){
    if (state.side === "back" && state.sheetPattern.enabled) return false;
    const layer = getSelection();
    if (!layer || layer.locked) return false;
    const cv = getActiveCanvas();
    let nx = layer.x + dx;
    let ny = layer.y + dy;
    if (state.grid.show){
      const g = mmToPx(state.grid.spacingMM);
      nx = Math.round(nx/g)*g;
      ny = Math.round(ny/g)*g;
      if (Math.abs(nx - cv.width/2) < g/2) nx = cv.width/2;
      if (Math.abs(ny - cv.height/2) < g/2) ny = cv.height/2;
    }
    layer.x = nx;
    layer.y = ny;
    drawOverlay();
    redrawAll();
    return true;
  }

  // Outline + handles (overlay-driven)
  function drawSelectionOutline(side){
    if (state.side!==side) return;
    const layer = getSelection();
    if (!layer || layer.locked) return;
    // Outline is rendered in drawOverlay()
  }

  // Hit testing
  function pointInLayer(l, x, y){
    const cos = Math.cos(l.rotation*Math.PI/180);
    const sin = Math.sin(l.rotation*Math.PI/180);
    const dx = x - l.x, dy = y - l.y;
    let lx =  cos*dx + sin*dy;
    let ly = -sin*dx + cos*dy;
    lx /= (l.flipX?-l.scaleX:l.scaleX);
    ly /= (l.flipY?-l.scaleY:l.scaleY);
    const ctx = getActiveCtx();
    const metrics = measureLayerBox(l, ctx);
    const halfW = metrics.width / 2;
    const halfH = metrics.height / 2;
    const shiftedX = lx + (metrics.anchorShiftX || 0);
    return {
      hit: shiftedX>=-halfW && shiftedX<=halfW && ly>=-halfH && ly<=halfH,
      w: metrics.width,
      h: metrics.height
    };
  }

  // Canvas drag (move)
  let drag = null;
  function canvasToImageCoords(ev, cv){
    const rect = cv.getBoundingClientRect();
    const x = (ev.clientX-rect.left) * (cv.width/rect.width);
    const y = (ev.clientY-rect.top) * (cv.height/rect.height);
    return {x,y,rect};
  }
  function onCanvasDown(ev){
    if (state.side==="back" && state.sheetPattern.enabled) return;
    if (inlineEditingLayerId) closeInlineCanvasEditor(true);
    const cv = getActiveCanvas();
    const {x,y} = canvasToImageCoords(ev, cv);
    const page = getActivePage();
    for (let i=page.layers.length-1;i>=0;i--){
      const l = page.layers[i];
      if (l.locked) continue;
      const ht = pointInLayer(l,x,y);
      if (ht.hit){
        setSelection(l.id);
        drag = {mode:"move", startX:x, startY:y, start:cloneLayer(l)};
        return;
      }
    }
    setSelection(null);
  }
  function onCanvasMove(ev){
    if (!drag) return;
    const cv = getActiveCanvas();
    const {x,y} = canvasToImageCoords(ev, cv);
    const l = getSelection();
    if (!l) return;
    if (drag.mode==="move"){
      let nx = drag.start.x + (x - drag.startX);
      let ny = drag.start.y + (y - drag.startY);
      if (state.grid.show){
      const g = mmToPx(state.grid.spacingMM);
      nx = Math.round(nx/g)*g;
      ny = Math.round(ny/g)*g;

      // ✅ snap to center lines if close
      if (Math.abs(nx - cv.width/2) < g/2) nx = cv.width/2;
      if (Math.abs(ny - cv.height/2) < g/2) ny = cv.height/2;
    }
      l.x = nx; l.y = ny;
      drawOverlay();
      redrawAll();
    }
  }

  function onCanvasUp(){ drag=null; }

  canvasFront.addEventListener("mousedown", onCanvasDown);
  canvasBack .addEventListener("mousedown", onCanvasDown);
  window.addEventListener("mousemove", onCanvasMove);
  window.addEventListener("mouseup", onCanvasUp);
  function onCanvasDoubleClick(ev){
    if (state.side==="back" && state.sheetPattern.enabled) return;
    const cv = getActiveCanvas();
    const {x,y} = canvasToImageCoords(ev, cv);
    const l = getSelection();
    if (!l || l.type!=="text") return;
    const hit = pointInLayer(l,x,y);
    if (!hit.hit) return;
    ev.preventDefault();
    openInlineCanvasEditor();
  }
  canvasFront.addEventListener("dblclick", onCanvasDoubleClick);
  canvasBack.addEventListener("dblclick", onCanvasDoubleClick);

  // Overlay DOM for handles
  const HANDLE_SIZE = 12;
  function drawOverlay(){
    overlayEl.innerHTML = "";
    selectionOverlayRect = null;
    const l = getSelection();
    if (!l || l.locked) return;
    if (state.side === "back" && state.sheetPattern.enabled) return;

    const cv = getActiveCanvas();
    const rect = cv.getBoundingClientRect();
    const ctx = getActiveCtx();

    // ---- Measure the unscaled, unrotated local box of the layer ----
    const metrics = measureLayerBox(l, ctx);
    const { width:w, height:h, anchorShiftX, lineHeight } = metrics;

    // ---- Helpers to convert local -> image -> screen ----
    function localToImage(ix, iy){
      let x = ix * (l.flipX ? -l.scaleX : l.scaleX);
      let y = iy * (l.flipY ? -l.scaleY : l.scaleY);
      const cos = Math.cos(l.rotation * Math.PI/180);
      const sin = Math.sin(l.rotation * Math.PI/180);
      const rx = cos*x - sin*y;
      const ry = sin*x + cos*y;
      return { x: l.x + rx, y: l.y + ry };
    }
    function imageToScreen(p){
      return {
        x: rect.left + (p.x * rect.width / cv.width),
        y: rect.top  + (p.y * rect.height / cv.height)
      };
    }

    const half = { x: w/2, y: h/2 };

    // Shift the selection box horizontally to respect text alignment.
    // (translate the box center by -anchorShiftX in local space)
    const a = anchorShiftX;
    const ptsLocal = {
      tl:{x:-half.x - a, y:-half.y},  tm:{x:0 - a, y:-half.y},  tr:{x:+half.x - a, y:-half.y},
      mr:{x:+half.x - a, y:0},        br:{x:+half.x - a, y:+half.y}, bm:{x:0 - a, y:+half.y},
      bl:{x:-half.x - a, y:+half.y},  ml:{x:-half.x - a, y:0}
    };

    const ptsScreen = {};
    for (const k in ptsLocal) ptsScreen[k] = imageToScreen(localToImage(ptsLocal[k].x, ptsLocal[k].y));
    const rotHandle = imageToScreen(localToImage(0 - a, -half.y - 30));

    // ---- Selection outline (screen-aligned) ----
    const minX = Math.min(ptsScreen.tl.x, ptsScreen.tr.x, ptsScreen.bl.x, ptsScreen.br.x) - rect.left;
    const maxX = Math.max(ptsScreen.tl.x, ptsScreen.tr.x, ptsScreen.bl.x, ptsScreen.br.x) - rect.left;
    const minY = Math.min(ptsScreen.tl.y, ptsScreen.tr.y, ptsScreen.bl.y, ptsScreen.br.y) - rect.top;
    const maxY = Math.max(ptsScreen.tl.y, ptsScreen.tr.y, ptsScreen.bl.y, ptsScreen.br.y) - rect.top;
    selectionOverlayRect = {
      left:minX,
      top:minY,
      width:maxX - minX,
      height:maxY - minY,
      centerX: minX + (maxX - minX)/2,
      centerY: minY + (maxY - minY)/2,
      rotation: l.rotation || 0,
      scaleX: rect.width / cv.width,
      scaleY: rect.height / cv.height,
      layerId: l.id,
      fontPxCss: l.type==="text" ? Math.max(8, Math.round(mmToPx(l.fontSize))) * (rect.height / cv.height) : null,
      lineHeightCss: l.type==="text" ? lineHeight * (rect.height / cv.height) : null
    };

    const outline = document.createElement("div");
    outline.className = "sel-outline";
    outline.style.left   = `${minX}px`;
    outline.style.top    = `${minY}px`;
    outline.style.width  = `${maxX - minX}px`;
    outline.style.height = `${maxY - minY}px`;
    overlayEl.appendChild(outline);

    // ---- Handles ----
    function addHandle(pt, role){
      const hdl = document.createElement("div");
      hdl.className = "handle";
      hdl.dataset.role = role;
      hdl.style.left = (pt.x - rect.left - HANDLE_SIZE/2) + "px";
      hdl.style.top  = (pt.y - rect.top  - HANDLE_SIZE/2) + "px";
      overlayEl.appendChild(hdl);
      return hdl;
    }
    const handles = {
      tl:addHandle(ptsScreen.tl,"tl"),
      tm:addHandle(ptsScreen.tm,"tm"),
      tr:addHandle(ptsScreen.tr,"tr"),
      mr:addHandle(ptsScreen.mr,"mr"),
      br:addHandle(ptsScreen.br,"br"),
      bm:addHandle(ptsScreen.bm,"bm"),
      bl:addHandle(ptsScreen.bl,"bl"),
      ml:addHandle(ptsScreen.ml,"ml"),
      rot:addHandle(rotHandle,"rot")
    };

    // ---- Equal-margin guides (approximate, ignore rotation) ----
    (function drawEqualMarginGuides(){
      const sx = (l.flipX ? -l.scaleX : l.scaleX);
      const sy = (l.flipY ? -l.scaleY : l.scaleY);
      const centerX = l.x + a * sx;          // include alignment shift
      const halfScaledX = half.x * Math.abs(sx);
      const centerY = l.y;                    // text is vertically centered already
      const halfScaledY = half.y * Math.abs(sy);

      const leftDist   = centerX - halfScaledX;
      const rightDist  = cv.width - (centerX + halfScaledX);
      const topDist    = centerY - halfScaledY;
      const bottomDist = cv.height - (centerY + halfScaledY);
      const tol = mmToPx(1);

      if (Math.abs(leftDist - rightDist) < tol){
        const vGuide = document.createElement("div");
        vGuide.style.position = "absolute";
        vGuide.style.left = `${(rect.width/2)|0}px`;
        vGuide.style.top = "0";
        vGuide.style.width = "1px";
        vGuide.style.height = "100%";
        vGuide.style.background = "rgba(0,200,0,0.45)";
        overlayEl.appendChild(vGuide);
      }
      if (Math.abs(topDist - bottomDist) < tol){
        const hGuide = document.createElement("div");
        hGuide.style.position = "absolute";
        hGuide.style.left = "0";
        hGuide.style.top = `${(rect.height/2)|0}px`;
        hGuide.style.width = "100%";
        hGuide.style.height = "1px";
        hGuide.style.background = "rgba(0,200,0,0.45)";
        overlayEl.appendChild(hGuide);
      }
    })();

    // ---- Drag logic (move/scale/rotate) ----
    let hdDrag = null;
    function screenToImage(sxScr, syScr){
      const x = (sxScr - rect.left) * (cv.width / rect.width);
      const y = (syScr - rect.top)  * (cv.height / rect.height);
      return {x,y};
    }

    overlayEl.onmousedown = (e)=>{
      const role = e.target?.dataset?.role;
      if (!role) return;
      e.preventDefault();
      hdDrag = { role, startScr:{x:e.clientX, y:e.clientY}, start: cloneLayer(getSelection()) };
      window.addEventListener("mousemove", onHdMove);
      window.addEventListener("mouseup", onHdUp, { once:true });
    };

    function onHdMove(e){
      const lcur = getSelection(); if(!lcur) return;
      const start = hdDrag.start;
      const curImg = screenToImage(e.clientX, e.clientY);
      const startImg = screenToImage(hdDrag.startScr.x, hdDrag.startScr.y);

      if (hdDrag.role === "rot"){
        const c = { x:start.x, y:start.y };
        const ang0 = Math.atan2(startImg.y - c.y, startImg.x - c.x);
        const ang1 = Math.atan2(curImg.y   - c.y, curImg.x   - c.x);
        lcur.rotation = ((start.rotation || 0) + (ang1 - ang0) * 180/Math.PI);
      } else {
        const cos = Math.cos((start.rotation || 0) * Math.PI/180);
        const sin = Math.sin((start.rotation || 0) * Math.PI/180);
        const dx = curImg.x - startImg.x;
        const dy = curImg.y - startImg.y;
        const dlx =  cos*dx + sin*dy;
        const dly = -sin*dx + cos*dy;

        // Measure the START box (don’t use the live w/h while dragging)
        const metrics0 = measureLayerBox(start, ctx);
        const w0 = metrics0.width;
        const h0 = metrics0.height;

        function applyScale(signX, signY){
          const dxLocal = (signX||0)*dlx;
          const dyLocal = (signY||0)*dly;
          let sx = start.scaleX + (dxLocal/(w0/2));
          let sy = start.scaleY + (dyLocal/(h0/2));

          const isCorner = signX && signY;
          if (isCorner){
            const uni = Math.max(sx, sy);
            sx = sy = uni;
          }

          if (state.grid.show){
            const g = mmToPx(state.grid.spacingMM);
            const snapW = Math.max(g, Math.round((w0*sx)/g)*g);
            const snapH = Math.max(g, Math.round((h0*sy)/g)*g);
            sx = snapW / w0;
            sy = snapH / h0;
          }
          lcur.scaleX = clamp(sx, 0.05, 20);
          lcur.scaleY = clamp(sy, 0.05, 20);
        }

        switch(hdDrag.role){
          case "tl": applyScale(-1,-1); break;
          case "tr": applyScale( 1,-1); break;
          case "br": applyScale( 1, 1); break;
          case "bl": applyScale(-1, 1); break;
          case "ml": applyScale(-1, 0); break;
          case "mr": applyScale( 1, 0); break;
          case "tm": applyScale( 0,-1); break;
          case "bm": applyScale( 0, 1); break;
        }
      }

      drawOverlay();
      redrawAll();
    }

    function onHdUp(){
      window.removeEventListener("mousemove", onHdMove);
      hdDrag = null;
    }
    if (inlineEditingLayerId) updateInlineEditorPosition();
  }

  // Inspector links
  rotRange.addEventListener("input", ()=>{ const s=getSelection(); if(s){ s.rotation=parseFloat(rotRange.value)||0; drawOverlay(); redrawAll(); }});
  opacityRange.addEventListener("input", ()=>{ const s=getSelection(); if(s){ s.opacity=parseFloat(opacityRange.value)||1; redrawAll(); }});
  lockChk.addEventListener("change", ()=>{ const s=getSelection(); if(s){ s.locked=lockChk.checked; drawOverlay(); redrawAll(); }});

  function normalizeEditorText(str=""){
    return str.replace(/\r/g,"").replace(/\u00a0/g," ").replace(/\t/g," ");
  }
  function escapeHtml(str=""){
    return str
      .replace(/&/g,"&amp;")
      .replace(/</g,"&lt;")
      .replace(/>/g,"&gt;")
      .replace(/\"/g,"&quot;")
      .replace(/'/g,"&#39;");
  }
  function textToHtml(str=""){
    if (!str) return "";
    return str.split(/\n/).map(line => line ? escapeHtml(line) : "&nbsp;").join("<br>");
  }
  function refreshInlineEditorToolbar(){
    if (tinyOverlayEditor){
      tinyOverlayEditor.dispatch("cardmaker-state");
    }
  }
  function loadTinyScript(){
    return new Promise((resolve,reject)=>{
      if (typeof tinymce !== "undefined") return resolve();
      const script = document.querySelector('script[data-role="tinymce"]');
      if (!script){
        reject(new Error("TinyMCE script not found"));
        return;
      }
      script.addEventListener("load", ()=>resolve(), { once:true });
      script.addEventListener("error", ()=>reject(new Error("TinyMCE failed to load")), { once:true });
    });
  }

  function applyToolbarUpdate(overrides){
    commitTextControlValues({ ...overrides, skipInlineContent:true });
  }
  function registerTinyToolbar(editor){
    editor.ui.registry.addMenuButton("cardfont",{
      text:"Font",
      fetch:(callback)=>{
        callback(STANDARD_FONTS.map(([value,label])=>({
          type:"menuitem",
          text:label,
          onAction:()=>applyToolbarUpdate({ fontFamily:value })
        })));
      }
    });
    editor.ui.registry.addMenuButton("cardfontsize",{
      text:"Size",
      fetch:(callback)=>{
        callback(FONT_SIZE_OPTIONS.map(size=>({
          type:"menuitem",
          text:`${size} mm`,
          onAction:()=>applyToolbarUpdate({ fontSize:size })
        })));
      }
    });
    editor.ui.registry.addMenuButton("cardalign",{
      text:"Align",
      fetch:(callback)=>{
        callback([
          { text:"Left", value:"left" },
          { text:"Center", value:"center" },
          { text:"Right", value:"right" }
        ].map(item=>({
          type:"menuitem",
          text:item.text,
          onAction:()=>applyToolbarUpdate({ align:item.value })
        })));
      }
    });
    editor.ui.registry.addToggleButton("cardbold",{
      icon:"bold",
      tooltip:"Bold",
      onAction:()=>{
        const l = getSelection();
        if (!l || l.type!=="text") return;
        const weight = parseInt(l.fontWeight||"400",10) >= 600 ? "400" : "700";
        applyToolbarUpdate({ fontWeight:weight });
      },
      onSetup:(api)=>{
        const handler = ()=>{
          const l = getSelection();
          api.setActive(!!l && l.type==="text" && parseInt(l.fontWeight||"400",10) >= 600);
        };
        editor.on("cardmaker-state", handler);
        handler();
        return ()=>editor.off("cardmaker-state", handler);
      }
    });
    editor.ui.registry.addToggleButton("carditalic",{
      icon:"italic",
      tooltip:"Italic",
      onAction:()=>{
        const l = getSelection();
        if (!l || l.type!=="text") return;
        const next = (l.fontStyle && l.fontStyle!=="normal") ? "normal" : "italic";
        applyToolbarUpdate({ fontStyle:next });
      },
      onSetup:(api)=>{
        const handler = ()=>{
          const l = getSelection();
          api.setActive(!!l && l.type==="text" && l.fontStyle === "italic");
        };
        editor.on("cardmaker-state", handler);
        handler();
        return ()=>editor.off("cardmaker-state", handler);
      }
    });
    editor.ui.registry.addButton("cardcolor",{
      icon:"color-picker",
      tooltip:"Text Color",
      onAction:()=>{
        const l = getSelection();
        const current = l?.fillStyle || "#111111";
        editor.windowManager.open({
          title:"Text Color",
          body:{
            type:"panel",
            items:[{ type:"colorinput", name:"color", label:"Color" }]
          },
          initialData:{ color: current },
          buttons:[
            { type:"cancel", text:"Cancel" },
            { type:"submit", text:"Apply", primary:true }
          ],
          onSubmit(api){
            const data = api.getData();
            if (data?.color){
              applyToolbarUpdate({ fillStyle:data.color });
            }
            api.close();
          }
        });
      }
    });
  }

  function ensureTinyOverlayEditor(){
    if (tinyOverlayEditor) return Promise.resolve(tinyOverlayEditor);
    if (tinyOverlayInitPromise) return tinyOverlayInitPromise;
    tinyOverlayInitPromise = loadTinyScript().then(()=>{
      const baseConfig = {
        target: canvasTextEditor,
        inline: true,
        menubar: "file edit view insert format tools table help",
        toolbar_mode: "sliding",
        toolbar_sticky: true,
        toolbar: "undo redo | cardfont cardfontsize | bold italic underline | cardalign | bullist numlist outdent indent | cardcolor | removeformat",
        skin: "oxide-dark",
        content_css: false,
        setup(editor){
          registerTinyToolbar(editor);
          editor.on("input change keyup", ()=>{
            if (!inlineEditingLayerId) return;
            const text = normalizeEditorText(editor.getContent({ format:"text" }));
            if (textContent && textContent.value !== text){
              textContent.value = text;
            }
            commitTextControlValues({ text, skipInlineContent:true });
          });
          editor.on("keydown", evt=>{
            if (evt.key === "Escape"){
              evt.preventDefault();
              closeInlineCanvasEditor(false);
            }
            if ((evt.key === "Enter" && (evt.metaKey || evt.ctrlKey))){
              evt.preventDefault();
              closeInlineCanvasEditor(true);
            }
          });
          editor.on("blur", ()=>{
            if (inlineEditingLayerId){
              closeInlineCanvasEditor(true);
            }
          });
        }
      };
      if (inlineEditorToolbar){
        baseConfig.fixed_toolbar_container = "#inlineEditorToolbar";
      }
      return tinymce.init(baseConfig);
    }).then(editors=>{
      tinyOverlayEditor = editors[0];
      tinyOverlayInitPromise = null;
      return tinyOverlayEditor;
    }).catch(err=>{
      tinyOverlayInitPromise = null;
      console.error("TinyMCE inline editor failed", err);
      throw err;
    });
    return tinyOverlayInitPromise;
  }
  function commitTextControlValues(opts={}){
    const l = getSelection(); if (!l || l.type!=="text") return;
    if (opts.text !== undefined && textContent) textContent.value = opts.text;
    if (opts.fontFamily !== undefined && fontFamily) fontFamily.value = opts.fontFamily;
    if (opts.fontSize !== undefined && fontSize) fontSize.value = opts.fontSize;
    if (opts.fontWeight !== undefined && fontWeight) fontWeight.value = opts.fontWeight;
    if (opts.align !== undefined && textAlignSel) textAlignSel.value = opts.align;
    if (opts.fillStyle !== undefined && fillStyleInp) fillStyleInp.value = opts.fillStyle;

    l.text = textContent.value;
    l.fontFamily = fontFamily.value;
    l.fontSize = parseFloat(fontSize.value)||10;
    l.fontWeight = fontWeight.value;
    l.align = textAlignSel.value;
    l.fillStyle = fillStyleInp.value;
    l.strokeStyle = strokeStyleInp.value;
    l.lineWidth = parseFloat(lineWidthInp.value)||0;
    l.fontStyle = opts.fontStyle !== undefined ? opts.fontStyle : (l.fontStyle || "normal");

    redrawAll();
    drawOverlay();

    if (inlineEditingLayerId && l.id === inlineEditingLayerId){
      if (!opts.skipInlineContent){
        if (tinyOverlayEditor){
          tinyOverlayEditor.setContent(textToHtml(l.text || ""), { format:'html' });
        } else if (canvasTextEditor){
          canvasTextEditor.textContent = l.text || "";
        }
      }
      updateInlineEditorStyles(l);
    }
    refreshInlineEditorToolbar();
  }
  function updateInlineEditorPosition(){
    if (!canvasTextEditor || !inlineEditingLayerId || !selectionOverlayRect) return;
    if (!inlineEditorIsVisible()) return;
    if (selectionOverlayRect.layerId !== inlineEditingLayerId){
      canvasTextEditor.style.display = "none";
      canvasTextEditor.classList.remove("active");
      return;
    }
    const l = getSelection();
    if (!l || l.id !== inlineEditingLayerId){
      canvasTextEditor.style.display = "none";
      canvasTextEditor.classList.remove("active");
      return;
    }
    const b = selectionOverlayRect;
    const scaleY = b.scaleY || b.scaleX || 1;
    canvasTextEditor.style.display = "block";
    canvasTextEditor.style.left = `${b.centerX}px`;
    canvasTextEditor.style.top = `${b.centerY}px`;
    canvasTextEditor.style.width = `${Math.max(80,b.width)}px`;
    canvasTextEditor.style.minHeight = `${Math.max(40,b.height)}px`;
    canvasTextEditor.style.transform = `translate(-50%,-50%) rotate(${l.rotation||0}deg)`;
    const fontPx = b.fontPxCss || Math.max(12, Math.round(mmToPx(l.fontSize) * scaleY));
    canvasTextEditor.style.fontSize = `${fontPx}px`;
    canvasTextEditor.style.lineHeight = b.lineHeightCss ? `${b.lineHeightCss}px` : "1.4";
  }
  function updateInlineEditorStyles(layer){
    if (!canvasTextEditor || !layer) return;
    const target = tinyOverlayEditor ? tinyOverlayEditor.getBody() : canvasTextEditor;
    const caretColor = layer.fillStyle || "#111";
    target.style.fontFamily = layer.fontFamily;
    target.style.fontWeight = layer.fontWeight;
    target.style.fontStyle = layer.fontStyle || "normal";
    target.style.textAlign = layer.align;
    target.style.color = "transparent";
    target.style.background = "transparent";
    target.style.textShadow = "none";
    target.style.caretColor = caretColor;
  }
  async function openInlineCanvasEditor(){
    const l = getSelection();
    if (!l || l.type!=="text" || !canvasTextEditor || !selectionOverlayRect) return;
    inlineEditingLayerId = l.id;
    canvasTextEditor.classList.add("active");
    updateInlineEditorPosition();
    canvasTextEditor.dataset.layerId = l.id;
    canvasTextEditor.style.display = "block";
    canvasTextEditor.textContent = l.text || "";
    if (inlineEditorToolbar){
      inlineEditorToolbar.classList.add("visible");
      inlineEditorToolbar.classList.remove("hidden");
    }
    try{
      const editor = await ensureTinyOverlayEditor();
      editor.setContent(textToHtml(l.text || ""));
      updateInlineEditorStyles(l);
      refreshInlineEditorToolbar();
      editor.focus();
      updateInlineEditorPosition();
    }catch(err){
      console.error("Inline editor unavailable:", err);
      canvasTextEditor.textContent = l.text || "";
      canvasTextEditor.focus({ preventScroll:true });
    }
  }
  function closeInlineCanvasEditor(save){
    if (!inlineEditingLayerId || !canvasTextEditor) return;
    if (save){
      let newText = "";
      if (tinyOverlayEditor){
        newText = normalizeEditorText(tinyOverlayEditor.getContent({ format:"text" }));
      } else {
        newText = normalizeEditorText(canvasTextEditor.textContent || "");
      }
      if (textContent){
        commitTextControlValues({ text:newText, skipInlineContent:true });
      }
    }
    inlineEditingLayerId = null;
    canvasTextEditor.style.display = "none";
    canvasTextEditor.classList.remove("active");
    canvasTextEditor.textContent = "";
    if (inlineEditorToolbar){
      inlineEditorToolbar.classList.remove("visible");
      inlineEditorToolbar.classList.add("hidden");
    }
    redrawAll();
    drawOverlay();
  }

  // Text controls binding
  function updateTextControls(l){
    textControls.classList.toggle("hidden", !l || l.type!=="text");
    if (!l || l.type!=="text") return;
    textContent.value = l.text;
    fontFamily.value = l.fontFamily;
    fontSize.value = l.fontSize;
    fontWeight.value = l.fontWeight;
    textAlignSel.value = l.align;
    fillStyleInp.value = l.fillStyle;
    strokeStyleInp.value = l.strokeStyle;
    lineWidthInp.value = l.lineWidth;
    if (inlineEditingLayerId && l.id === inlineEditingLayerId){
      if (tinyOverlayEditor){
        tinyOverlayEditor.setContent(textToHtml(l.text || ""));
        updateInlineEditorStyles(l);
        refreshInlineEditorToolbar();
      } else if (canvasTextEditor){
        canvasTextEditor.textContent = l.text || "";
      }
      updateInlineEditorPosition();
    }
  }
  if (textContent){
    textContent.addEventListener("input", ()=>{
      commitTextControlValues();
    });
  }
  [fontFamily,fontSize,fontWeight,textAlignSel,fillStyleInp,strokeStyleInp,lineWidthInp].forEach(el=>{
    if (!el) return;
    el.addEventListener("input", commitTextControlValues);
  });
  if (canvasTextEditor){
    canvasTextEditor.addEventListener("input", ()=>{
      if (tinyOverlayEditor || !inlineEditingLayerId) return;
      const l = getSelection();
      if (!l || l.id !== inlineEditingLayerId) return;
      const val = normalizeEditorText(canvasTextEditor.textContent || "");
      if (textContent) textContent.value = val;
      commitTextControlValues({ text:val, skipInlineContent:true });
    });
    canvasTextEditor.addEventListener("blur", ()=>{
      if (!tinyOverlayEditor){
        closeInlineCanvasEditor(true);
      }
    });
    canvasTextEditor.addEventListener("keydown", (e)=>{
      if (!tinyOverlayEditor){
        if (e.key === "Escape"){
          e.preventDefault();
          closeInlineCanvasEditor(false);
        }
        if ((e.key === "Enter" && (e.metaKey || e.ctrlKey))){
          e.preventDefault();
          closeInlineCanvasEditor(true);
        }
      }
    });
  }

  function updateInspector(){
    const l = getSelection();
    if (!l){
      rotRange.value = 0; opacityRange.value = 1; lockChk.checked = false;
      textControls.classList.add("hidden");
      return;
    }
    rotRange.value = l.rotation||0;
    opacityRange.value = l.opacity??1;
    lockChk.checked = !!l.locked;
    updateTextControls(l);
  }

  // Buttons
  btnAddText.addEventListener("click", ()=>{
    if (state.side==="back" && state.sheetPattern.enabled) return;
    const layer = makeTextLayer();
    getActivePage().layers.push(layer);
    setSelection(layer.id);
    redrawAll();
  });
  imgInput.addEventListener("change", e=>{
    if (state.side==="back" && state.sheetPattern.enabled) { e.target.value=""; return; }
    const file = e.target.files[0];
    if (!file) return;
    const img = new Image();
    img.onload = ()=>{
      const layer = makeImageLayer(img);
      getActivePage().layers.push(layer);
      setSelection(layer.id);
      redrawAll();
    };
    img.src = URL.createObjectURL(file);
    e.target.value="";
  });

  btnDelete.addEventListener("click", ()=>{
    const page = getActivePage();
    const idx = page.layers.findIndex(l=>l.id===state.selection);
    if (idx>=0){ page.layers.splice(idx,1); state.selection=null; redrawAll(); updateInspector(); drawOverlay(); }
  });
  btnDuplicate.addEventListener("click", ()=>{
    const l = getSelection(); if (!l) return;
    const copy = cloneLayer(l);
    copy.id = crypto.randomUUID();
    copy.x += 10; copy.y += 10;
    getActivePage().layers.push(copy);
    setSelection(copy.id);
    redrawAll();
  });
  btnBringFwd.addEventListener("click", ()=>{
    const page = getActivePage(); const idx = page.layers.findIndex(l=>l.id===state.selection);
    if (idx>=0 && idx<page.layers.length-1){ const [it]=page.layers.splice(idx,1); page.layers.splice(idx+1,0,it); redrawAll(); }
  });
  btnSendBack.addEventListener("click", ()=>{
    const page = getActivePage(); const idx = page.layers.findIndex(l=>l.id===state.selection);
    if (idx>0){ const [it]=page.layers.splice(idx,1); page.layers.splice(idx-1,0,it); redrawAll(); }
  });
  btnFlipH.addEventListener("click", ()=>{ const l=getSelection(); if(l){ l.flipX=!l.flipX; redrawAll(); }});
  btnFlipV.addEventListener("click", ()=>{ const l=getSelection(); if(l){ l.flipY=!l.flipY; redrawAll(); }});
  if (btnUnlockLayer && lockedLayerList){
    btnUnlockLayer.addEventListener("click", ()=>{
      const page = getActivePage();
      const val = lockedLayerList.value;
      if (val){
        const layer = page.layers.find(l=>l.id===val);
        if (layer && layer.locked){
          layer.locked = false;
          setSelection(layer.id);
        }
        return;
      }
      let changed = false;
      page.layers.forEach(l=>{ if (l.locked){ l.locked=false; changed=true; }});
      if (changed){
        drawOverlay();
        redrawAll();
      }
    });
  }

  // Background controls
  bgColorFront.addEventListener("input", e=>{ state.bg.front.color = e.target.value; redrawAll(); });
  bgColorBack .addEventListener("input", e=>{ state.bg.back.color  = e.target.value; redrawAll(); });
  bgImgFront.addEventListener("change", e=>loadBgImage(e,"front"));
  bgImgBack .addEventListener("change", e=>loadBgImage(e,"back"));
  function loadBgImage(e, side){
    const f = e.target.files[0]; if(!f) return;
    const img = new Image();
    img.onload = ()=>{ state.bg[side].image = img; redrawAll(); };
    img.src = URL.createObjectURL(f);
    e.target.value="";
  }

  // Back offset (affects both editor and preview)
  [backOffsetXmm,backOffsetYmm].forEach(inp=>inp.addEventListener("input", ()=>{
    state.backOffsetMM.x = parseFloat(backOffsetXmm.value)||0;
    state.backOffsetMM.y = parseFloat(backOffsetYmm.value)||0;
    redrawAll();       // updates canvas immediately
    refreshPrintPreview(); // push to preview immediately
  }));

  // Settings
  unitSel.addEventListener("change", ()=>{ state.units = unitSel.value; });
  showBleedChk.addEventListener("change", ()=>{ state.showBleed = showBleedChk.checked; redrawAll(); });
  showGridChk.addEventListener("change", ()=>{ state.grid.show = showGridChk.checked; redrawAll(); });
  gridSpacingInput.addEventListener("input", ()=>{ state.grid.spacingMM = parseFloat(gridSpacingInput.value)||5; redrawAll(); });
  trimBoxesChk.addEventListener("change", ()=>{ state.trimBoxes = trimBoxesChk.checked; refreshPrintPreview(); });
  backFlipSel.addEventListener("change", ()=>{ state.backFlip = backFlipSel.value; refreshPrintPreview(); });

  function readCardDims(){
    const factor = (state.units==="mm")?1:(25.4/state.dpi);
    state.cardWidthMM  = parseFloat(widthInput.value||"85") * factor;
    state.cardHeightMM = parseFloat(heightInput.value||"55") * factor;
  }
  [widthInput,heightInput,dpiInput,bleedInput,cropMarksChk,gutterInput,pageMarginInput].forEach(el=>el.addEventListener("input", ()=>{
    state.dpi = parseInt(dpiInput.value||"300",10);
    state.bleedMM = parseFloat(bleedInput.value||"3");
    state.cropMarks = !!cropMarksChk.checked;
    state.gutterMM = parseFloat(gutterInput.value||"5");
    state.pageMarginMM = parseFloat(pageMarginInput.value||"10");
    readCardDims();
    resizeCanvasToCard();
  }));

  // Thumbs
  function syncThumbs(){
    if (thumbFront) thumbFront.src = canvasFront.toDataURL("image/png");
    if (thumbBack)  thumbBack.src  = canvasBack.toDataURL("image/png");
    renderCanvasTabs();
  }

  // Pattern back UI behavior
  sheetPatternMode.addEventListener("change", ()=>{
    state.sheetPattern.enabled = sheetPatternMode.checked;
    sheetPatternBox.classList.toggle("hidden", !sheetPatternMode.checked);
    updateBackEditorDisabled();
    drawOverlay();
    refreshPrintPreview();
  });
  sheetPatternImg.addEventListener("change", (e)=>{
    const f = e.target.files[0]; if(!f) return;
    const img = new Image();
    img.onload = ()=>{ state.sheetPattern.image = img; refreshPrintPreview(); };
    img.src = URL.createObjectURL(f);
    e.target.value="";
  });
  [sheetOffsetX, sheetOffsetY].forEach(el=>el.addEventListener("input", ()=>{
    state.sheetPattern.offsetMM.x = parseFloat(sheetOffsetX.value)||0;
    state.sheetPattern.offsetMM.y = parseFloat(sheetOffsetY.value)||0;
    refreshPrintPreview();
  }));

  function updateBackEditorDisabled(){
    if (!backDisabledOverlay){
      backDisabledOverlay = document.createElement("div");
      backDisabledOverlay.style.cssText = `
        position:absolute; inset:0;
        background:rgba(0,0,0,0.35);
        display:flex; align-items:center; justify-content:center;
        color:#fff; font:14px/1.2 system-ui,Arial,sans-serif;
        z-index:5; pointer-events:auto;`;
      backDisabledOverlay.textContent = "Back editor disabled — Pattern Back (full sheet) is ON";
      backDisabledOverlay.style.display = "none";
      canvasHolder.appendChild(backDisabledOverlay);
    }
    const enable = state.sheetPattern.enabled && state.side==="back";
    backDisabledOverlay.style.display = enable ? "flex" : "none";
    canvasBack.style.opacity = enable ? "0.5" : "1";
    overlayEl.style.display = enable ? "none" : "block";
  }

  // Per-card backs
  perCardBacks.addEventListener("change", ()=>{
    state.perCard.enabled = perCardBacks.checked;
    perCardBox.classList.toggle("hidden", !perCardBacks.checked);
    refreshPrintPreview();
  });
  perCardBackFiles.addEventListener("change", (e)=>{
    const files = Array.from(e.target.files||[]);
    if (!files.length) return;
    let remaining = files.length;
    const imgs = [];
    files.forEach((f, idx)=>{
      const img = new Image();
      img.onload = ()=>{ imgs[idx]=img; if(--remaining===0){ state.perCard.images = imgs.filter(Boolean); refreshPrintPreview(); } };
      img.src = URL.createObjectURL(f);
    });
    e.target.value="";
  });

  // 3D Preview
  function init3DLayerStyles(){
    // center all parts inside card3d so transforms use the card center
    [faceFront, faceBack, ...edges].forEach(el=>{
      el.style.position = "absolute";
      el.style.top = "50%";
      el.style.left = "50%";
      el.style.transformOrigin = "50% 50%";
      el.style.backfaceVisibility = "hidden";
      el.style.willChange = "transform";
    });
  }
  init3DLayerStyles();

  let rotX= -10, rotY= 25, startDrag=null, zoomFactor=1;
function update3DTextures(apply = true){
  const fUrl = canvasFront.toDataURL();
  const bUrl = (state.sheetPattern.enabled && state.sheetPattern.image)
    ? (patternCardPatchURL() || canvasBack.toDataURL())
    : canvasBack.toDataURL();

  // Force repaint so browser doesn’t cache identical data: URLs
  faceFront.style.backgroundImage = "none";
  faceBack .style.backgroundImage = "none";
  requestAnimationFrame(() => {
    faceFront.style.backgroundImage = `url(${fUrl})`;
    faceBack .style.backgroundImage = `url(${bUrl})`;
  });

  const baseW = Math.max(300, Math.min(680, canvasFront.width/1.6));
  const w = Math.round(baseW * (parseFloat(zoom3D.value)||1));
  const h = Math.round(w * (canvasFront.height/canvasFront.width));
  const t = Math.max(1, Math.round(mmToPx(parseFloat(thicknessInp.value||"0.4"))/3));

  [faceFront, faceBack].forEach(face => {
    face.style.width = w + "px";
    face.style.height = h + "px";
  });
  faceFront.style.transform = `translate(-50%,-50%) translateZ(${t/2}px)`;
  faceBack .style.transform = `translate(-50%,-50%) rotateY(180deg) translateZ(${t/2}px)`;

  const [edgeTop, edgeRight, edgeBottom, edgeLeft] = edges;
  edgeTop.style.width = w + "px";    edgeTop.style.height = t + "px";
  edgeBottom.style.width = w + "px"; edgeBottom.style.height = t + "px";
  edgeLeft.style.width = t + "px";   edgeLeft.style.height = h + "px";
  edgeRight.style.width = t + "px";  edgeRight.style.height = h + "px";
  edgeTop.style.transform    = `translate(-50%,-50%) translateY(${-h/2}px) rotateX(90deg)`;
  edgeBottom.style.transform = `translate(-50%,-50%) translateY(${ h/2}px) rotateX(270deg)`;
  edgeRight.style.transform  = `translate(-50%,-50%) translateX(${ w/2}px) rotateY(90deg)`;
  edgeLeft.style.transform   = `translate(-50%,-50%) translateX(${-w/2}px) rotateY(270deg)`;

  card3d.style.width  = w + "px";
  card3d.style.height = h + "px";
  if (apply) card3d.style.transform = `rotateX(${rotX}deg) rotateY(${rotY}deg)`;
}




  function apply3DRotation(){
    card3d.style.transform = `rotateX(${rotX}deg) rotateY(${rotY}deg)`;
  }
  $("#btnUpdate3D").addEventListener("click", ()=> update3DTextures(true));
  $("#btnReset3D").addEventListener("click", ()=>{ rotX=-10; rotY=25; zoomFactor=1; zoom3D.value="1"; update3DTextures(true); });
  zoom3D.addEventListener("input", ()=>{ zoomFactor=parseFloat(zoom3D.value)||1; update3DTextures(true); });
  viewport.addEventListener("mousedown", (e)=>{
    startDrag = {x:e.clientX, y:e.clientY, rx:rotX, ry:rotY};
    card3d.classList.add("grabbing");
  });
  window.addEventListener("mousemove", (e)=>{
    if (!startDrag) return;
    const dx = e.clientX - startDrag.x;
    const dy = e.clientY - startDrag.y;
    rotY = startDrag.ry + dx*0.3;
    rotX = clamp(startDrag.rx - dy*0.3, -89, 89);
    apply3DRotation();
  });
  window.addEventListener("mouseup", ()=>{ startDrag=null; card3d.classList.remove("grabbing"); });
  viewport.addEventListener("wheel", (e)=>{
    e.preventDefault();
    const z = clamp(zoomFactor + (e.deltaY<0?0.05:-0.05), 0.5, 2);
    zoomFactor = z; zoom3D.value = z.toString();
    update3DTextures(true);
  }, {passive:false});

  // ----- Imposition / Print -----
  function computeGrid(){
    const A4 = {w:210, h:297};
    const minTop = 0.5; // safe top margin to avoid clip
    const margin = Math.max(minTop, state.pageMarginMM);
    const gutter = state.gutterMM;
    const wWithBleed = state.cardWidthMM + 2*state.bleedMM;
    const hWithBleed = state.cardHeightMM + 2*state.bleedMM;
    const usableW = A4.w - 2*margin;
    const usableH = A4.h - 2*margin;
    function fitCount(total, item, gap){
      if (item<=0) return 0;
      let n = Math.floor((total + gap) / (item + gap));
      return Math.max(0, n);
    }
    const cols = Math.max(1, fitCount(usableW, wWithBleed, gutter));
    const rows = Math.max(1, fitCount(usableH, hWithBleed, gutter));
    const gridW = cols*wWithBleed + (cols-1)*gutter;
    const gridH = rows*hWithBleed + (rows-1)*gutter;
    const startX = Math.max(minTop, margin + (usableW - gridW)/2);
    const startY = Math.max(minTop, margin + (usableH - gridH)/2);
    return {A4, margin, gutter, wWithBleed, hWithBleed, cols, rows, startX, startY};
  }

  function perCardBackDataURLs(){
    if (!state.perCard.enabled) return [];
    return state.perCard.images.map(img=>{
      const cv = document.createElement("canvas");
      cv.width = canvasBack.width; cv.height = canvasBack.height;
      const c = cv.getContext("2d");
      c.fillStyle = state.bg.back.color || "#fff";
      c.fillRect(0,0,cv.width,cv.height);
      const scale = Math.max(cv.width/img.width, cv.height/img.height);
      const w = img.width*scale, h = img.height*scale;
      c.drawImage(img, (cv.width-w)/2, (cv.height-h)/2, w, h);
      return cv.toDataURL("image/png");
    });
  }

function sheetPatternDataURL(){
  if (!(state.sheetPattern.enabled && state.sheetPattern.image)) return null;

  const img = state.sheetPattern.image;
  const A4_W_MM = 210, A4_H_MM = 297;
  const A4_RATIO = A4_W_MM / A4_H_MM;
  const imgRatio = img.width / img.height;
  const isA4ish = Math.abs(imgRatio - A4_RATIO) < 0.01;

  const cv = document.createElement("canvas");
  const c  = cv.getContext("2d");

  // white base
  if (isA4ish){
    cv.width  = img.width;
    cv.height = img.height;
  } else {
    cv.width  = Math.round(mmToPx(A4_W_MM));
    cv.height = Math.round(mmToPx(A4_H_MM));
  }
  c.fillStyle = "#fff";
  c.fillRect(0,0,cv.width,cv.height);

  // apply opacity
  const alpha = Math.max(0, Math.min(1, state.sheetPattern.opacity || 0));
  c.save();
  c.globalAlpha = alpha;

  if (isA4ish){
    const pxPerMMx = img.width  / A4_W_MM;
    const pxPerMMy = img.height / A4_H_MM;
    const offX = (state.sheetPattern.offsetMM.x || 0) * pxPerMMx;
    const offY = (state.sheetPattern.offsetMM.y || 0) * pxPerMMy;
    const x = (cv.width  - img.width)  / 2 + offX;
    const y = (cv.height - img.height) / 2 + offY;
    c.drawImage(img, Math.round(x), Math.round(y));
  } else {
    const contain = Math.min(cv.width / img.width, cv.height / img.height);
    const scale   = Math.min(1, contain);
    const w = Math.round(img.width  * scale);
    const h = Math.round(img.height * scale);
    const offX = mmToPx(state.sheetPattern.offsetMM.x || 0);
    const offY = mmToPx(state.sheetPattern.offsetMM.y || 0);
    const x = Math.round((cv.width  - w) / 2 + offX);
    const y = Math.round((cv.height - h) / 2 + offY);
    c.drawImage(img, x, y, w, h);
  }

  c.restore(); // reset alpha
  return cv.toDataURL("image/png");
}


function patternCardPatchURL(){
  if (!(state.sheetPattern.enabled && state.sheetPattern.image)) return null;

  const A4_W_MM = 210, A4_H_MM = 297;
  const img = state.sheetPattern.image;
  const opacity = Math.max(0, Math.min(1, state.sheetPattern.opacity ?? 1));

  const sheetCv = document.createElement("canvas");
  sheetCv.width = Math.round(mmToPx(A4_W_MM));
  sheetCv.height = Math.round(mmToPx(A4_H_MM));
  const sc = sheetCv.getContext("2d");

  // draw with alpha, NO white background
  sc.save();
  sc.globalAlpha = opacity;

  const A4_RATIO = A4_W_MM / A4_H_MM;
  const isA4ish = Math.abs((img.width/img.height) - A4_RATIO) < 0.01;

  if (isA4ish){
    const pxPerMMx = img.width  / A4_W_MM;
    const pxPerMMy = img.height / A4_H_MM;
    const offX = (state.sheetPattern.offsetMM.x || 0) * pxPerMMx;
    const offY = (state.sheetPattern.offsetMM.y || 0) * pxPerMMy;
    sc.drawImage(img,
      Math.round((sheetCv.width  - img.width)  / 2 + offX),
      Math.round((sheetCv.height - img.height) / 2 + offY)
    );
  } else {
    const contain = Math.min(sheetCv.width / img.width, sheetCv.height / img.height);
    const scale   = Math.min(1, contain);
    const w = Math.round(img.width  * scale);
    const h = Math.round(img.height * scale);
    const offX = mmToPx(state.sheetPattern.offsetMM.x || 0);
    const offY = mmToPx(state.sheetPattern.offsetMM.y || 0);
    sc.drawImage(img,
      Math.round((sheetCv.width  - w) / 2 + offX),
      Math.round((sheetCv.height - h) / 2 + offY),
      w, h
    );
  }
  sc.restore();

  // crop first card (no bleed) to a transparent card-sized canvas
  const G = computeGrid();
  const sx = Math.round(mmToPx(G.startX + state.bleedMM));
  const sy = Math.round(mmToPx(G.startY + state.bleedMM));
  const sw = Math.round(mmToPx(state.cardWidthMM));
  const sh = Math.round(mmToPx(state.cardHeightMM));

  const cardCv = document.createElement("canvas");
  cardCv.width  = Math.round(mmToPx(state.cardWidthMM));
  cardCv.height = Math.round(mmToPx(state.cardHeightMM));
  const cc = cardCv.getContext("2d");

  cc.drawImage(
    sheetCv,
    Math.max(0, Math.min(sheetCv.width  - 1, sx)),
    Math.max(0, Math.min(sheetCv.height - 1, sy)),
    Math.max(1, Math.min(sheetCv.width  - sx, sw)),
    Math.max(1, Math.min(sheetCv.height - sy, sh)),
    0, 0, cardCv.width, cardCv.height
  );

  return cardCv.toDataURL("image/png");
}



  function mapIndex(r,c, cols, rows){
    let rr = r, cc = c;
    if (state.backFlip==="h" || state.backFlip==="hv") cc = cols-1-c;
    if (state.backFlip==="v" || state.backFlip==="hv") rr = rows-1-r;
    return {rr, cc};
  }

  // Back offset visual overlay (SVG) for preview
  function backOffsetOverlaySVG(){
    const x = state.backOffsetMM.x||0, y = state.backOffsetMM.y||0;
    if (!x && !y) return "";
    const scale = 6;                 // px per mm for mini legend
    const ax = 60, ay = 32;          // origin inside legend box (higher)
    const tx = ax + x*scale;
    const ty = ay + y*scale;
    const label = `Back offset  X:${x.toFixed(2)}mm  Y:${y.toFixed(2)}mm`;
    return `
    <svg width="180" height="60" style="position:absolute; right:6mm; top:2mm; z-index:3">
      <rect x="0.5" y="0.5" width="179" height="59" fill="rgba(255,255,255,0.85)" stroke="rgba(0,0,0,0.2)"/>
      <text x="6" y="14" font-size="10" font-family="system-ui,Arial" fill="#111">${label}</text>
      <g transform="translate(6,20)">
        <rect x="0" y="0" width="160" height="32" fill="none" stroke="rgba(0,0,0,0.12)"/>
        <g>
          <circle cx="${ax}" cy="${ay}" r="2.5" fill="#333"/>
          <line x1="${ax}" y1="${ay}" x2="${tx}" y2="${ty}" stroke="#1976d2" stroke-width="2" marker-end="url(#arrow)"/>
        </g>
        <defs>
          <marker id="arrow" markerWidth="6" markerHeight="6" refX="5" refY="3" orient="auto">
            <polygon points="0,0 6,3 0,6" fill="#1976d2"/>
          </marker>
        </defs>
        <text x="${ax-40}" y="${ay-6}" font-size="8" fill="#333">origin</text>
      </g>
    </svg>`;
  }

  function buildImpositionHTML(){
    const G = computeGrid();
    const frontURL = canvasFront.toDataURL("image/png");
    const backURL  = canvasBack.toDataURL("image/png");
    const crop = state.cropMarks;
    const trim = state.trimBoxes;
    const bleed = state.bleedMM;
    const cardWmm = state.cardWidthMM, cardHmm = state.cardHeightMM;

    const perBacks = perCardBackDataURLs();
    const sheetPat = sheetPatternDataURL();

    // back offset (mm) applied to WHOLE card block (crops included)
    const backOffX = state.backOffsetMM.x || 0;
    const backOffY = state.backOffsetMM.y || 0;

    function gridHTML(side){
      const imgURL = (side==="front") ? frontURL : backURL;
      let blocks = "";
      let idx = 0;

      for (let r=0; r<G.rows; r++){
        for (let c=0; c<G.cols; c++){
          const pos = (side==="back") ? mapIndex(r,c, G.cols, G.rows) : {rr:r, cc:c};
          const baseX = +(G.startX + pos.cc*(G.wWithBleed + G.gutter)).toFixed(3);
          const baseY = +(G.startY + pos.rr*(G.hWithBleed + G.gutter)).toFixed(3);

          // apply offset to entire block on BACK page
          const cardX = side==="back" ? baseX + backOffX : baseX;
          const cardY = side==="back" ? baseY + backOffY : baseY;

          const url = (side==="back" && state.perCard.enabled && perBacks[idx])
            ? perBacks[idx]
            : imgURL;

          blocks += `
            <div class="card" style="left:${cardX}mm; top:${cardY}mm; width:${G.wWithBleed}mm; height:${G.hWithBleed}mm; z-index:1;">
              <div class="inner" style="left:${bleed}mm; top:${bleed}mm; width:${cardWmm}mm; height:${cardHmm}mm;">
                <img src="${url}">
                ${trim ? `<div class="trim" style="position:absolute;inset:0;border:0.3mm solid rgba(0,0,0,.6);box-sizing:border-box;"></div>` : ""}
              </div>
              ${crop ? '<div class="crop tl"></div><div class="crop tr"></div><div class="crop bl"></div><div class="crop br"></div>' : ''}
            </div>`;
          idx++;
        }
      }
      return blocks;
    }

    const backOffsetLegend = backOffsetOverlaySVG();

    const html = `<!doctype html><html><head><meta charset="utf-8">
      <title>Print — Business Cards</title>
      <style>
        @page { size: A4; margin: 0; }
        html,body{ margin:0; padding:0; }
        body{ -webkit-print-color-adjust: exact; print-color-adjust: exact;
              font-family: system-ui, Arial, sans-serif; color:#111; }
        .sheet{ position:relative; width:210mm; height:297mm; overflow:hidden; page-break-after: always; }
        .sheet:last-child{ page-break-after: auto; }
        .title{ position:absolute; left:10mm; top:5mm; font-weight:600; font-size:12pt; z-index:2 }
        .card{ position:absolute; background:#fff; }
        .inner{ position:absolute; overflow:hidden }
        .card img{ width:${cardWmm}mm; height:${cardHmm}mm; display:block }
        .crop:before, .crop:after{ content:""; position:absolute; background:#000 }
        .crop.tl:before{ width:0.2mm; height:${bleed}mm; left:${bleed}mm; top:0 }
        .crop.tl:after{ height:0.2mm; width:${bleed}mm; left:0; top:${bleed}mm }
        .crop.br:before{ width:0.2mm; height:${bleed}mm; right:${bleed}mm; bottom:0 }
        .crop.br:after{ height:0.2mm; width:${bleed}mm; right:0; bottom:${bleed}mm }
        .crop.tr:before{ width:0.2mm; height:${bleed}mm; right:${bleed}mm; top:0 }
        .crop.tr:after{ height:0.2mm; width:${bleed}mm; right:0; top:${bleed}mm }
        .crop.bl:before{ width:0.2mm; height:${bleed}mm; left:${bleed}mm; bottom:0 }
        .crop.bl:after{ height:0.2mm; width:${bleed}mm; left:0; bottom:${bleed}mm }
        .bg-full{ position:absolute; left:50%; top:50%; transform:translate(-50%,-50%); width:100%; height:auto; z-index:0; }

      </style></head><body>

      <div class="sheet">
        <div class="title">Front</div>
        ${gridHTML("front")}
      </div>

      <div class="sheet">
        <div class="title">Back</div>
        ${state.sheetPattern.enabled && sheetPat ? `<img class="bg-full" src="${sheetPat}" style="opacity:${state.sheetPattern.opacity}">` : ""}
        ${!state.sheetPattern.enabled ? gridHTML("back") : ""}
        ${!state.sheetPattern.enabled ? backOffsetLegend : "" }
      </div>

      <script>
        // Wait for all images so both pages render before printing
        function whenImagesReady(cb){
          const imgs = Array.from(document.images);
          let left = imgs.length;
          if (!left) return cb();
          imgs.forEach(i=>{
            if (i.complete) { if(--left===0) cb(); }
            else {
              i.addEventListener('load', ()=>{ if(--left===0) cb(); });
              i.addEventListener('error', ()=>{ if(--left===0) cb(); });
            }
          });
        }
        whenImagesReady(()=>{ setTimeout(()=>window.print(), 50); });
      </script>
    </body></html>`;
    return html;
  }

  function openPrintWindow(){
    const html = buildImpositionHTML();
    const w = window.open("", "_blank");
    w.document.open();
    w.document.write(html);
    w.document.close();
  }

  // --- Live preview (iframe) ---
  let previewURL=null;
  function ensurePrintPreview(){
    if (printPreviewWrap) return;

    printPreviewWrap = document.createElement("div");
    printPreviewWrap.style.cssText = `
      margin-top:12px;
      border:1px solid #242b3a;
      border-radius:12px;
      overflow:hidden;
      background:#0f1420;
      display:flex;
      flex-direction:column;
      align-items:center; /* center the A4 preview */
      padding-bottom:10px;
    `;

    const title = document.createElement("div");
    title.textContent = "Print Preview (A4)";
    title.style.cssText = `
      width:100%;
      padding:10px 12px;
      color:#cddbff;
      font-weight:600;
      background:#151821;
      border-bottom:1px solid #242b3a;
      box-sizing:border-box;
    `;

    printPreviewFrame = document.createElement("iframe");
    // ► Key bits: make it as wide as A4; keep aspect ratio of a single A4 page.
    // You can scroll inside to see page 2.
    printPreviewFrame.style.cssText = `
      width:210mm;           /* A4 width */
      max-width:100%;        /* shrink on narrow screens */
      aspect-ratio: 210 / 297;
      height:auto;           /* computed from aspect-ratio */
      background:#fff;
      border:0;
      display:block;
    `;

    printPreviewWrap.appendChild(title);
    printPreviewWrap.appendChild(printPreviewFrame);
    const stage = document.querySelector(".stage");
    stage.appendChild(printPreviewWrap);
  }

  function refreshPrintPreview(){
    ensurePrintPreview();
    const html = buildImpositionHTML().replace("window.print()", ""); // no auto print in preview
    if (previewURL) URL.revokeObjectURL(previewURL);
    const blob = new Blob([html], {type:"text/html"});
    previewURL = URL.createObjectURL(blob);
    // Force reload even if same URL by resetting src first
    printPreviewFrame.src = "about:blank";
    // Use requestAnimationFrame to ensure src change is observed
    requestAnimationFrame(()=>{ printPreviewFrame.src = previewURL; });
  }
  const debouncedPreview = debounce(refreshPrintPreview, 120);

  btnExportPDF.addEventListener("click", openPrintWindow);

// ==== Project save/load (.card ZIP with assets & fonts) ======================

// Lazy-load JSZip if not present
async function ensureJSZip(){
  if (window.JSZip) return window.JSZip;
  await new Promise((res, rej)=>{
    const s = document.createElement("script");
    s.src = "https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js";
    s.onload = res; s.onerror = rej;
    document.head.appendChild(s);
  });
  return window.JSZip;
}

// Helpers
function dataURLToBlob(dataURL){
  const [head, body] = dataURL.split(',');
  const mime = (head.match(/data:(.*?);/)||[])[1] || "application/octet-stream";
  const bin = atob(body);
  const u8  = new Uint8Array(bin.length);
  for (let i=0;i<bin.length;i++) u8[i] = bin.charCodeAt(i);
  return new Blob([u8], {type:mime});
}
async function imageToPNGBlob(img){
  if (!img) return null;
  // If it's already a dataURL (very unlikely for <img>), convert directly
  if (typeof img === "string" && img.startsWith("data:")) return dataURLToBlob(img);

  // Draw image to a canvas to get a clean PNG
  const cv = document.createElement("canvas");
  cv.width = img.width; cv.height = img.height;
  const c = cv.getContext("2d");
  c.drawImage(img, 0, 0);
  const url = cv.toDataURL("image/png");
  return dataURLToBlob(url);
}
function safeExt(name, fallback){ return (name && name.match(/\.[a-z0-9]{2,5}$/i)) ? "" : fallback; }
function cleanName(s, def){ return (s||def).replace(/[^\w.-]+/g,"_").slice(0,64); }

// Collect assets & produce a JSON manifest (clone of state with file paths)
async function buildExportPackage(){
  const manifest = JSON.parse(JSON.stringify({...state, selection:null}));

  // Ensure arrays exist
  manifest.customFonts = manifest.customFonts || [];
  const assets = []; // {path, blob}

  // ---- Backgrounds
  async function addBg(side, key, fileName){
    const img = state.bg?.[side]?.image;
    if (!img) return;
    const blob = await imageToPNGBlob(img);
    const path = `assets/${fileName}`;
    assets.push({path, blob});
    // store path in manifest
    if (!manifest.bg) manifest.bg = {front:{},back:{}};
    manifest.bg[side].image = path;
  }
  await addBg("front", "image", "bg_front.png");
  await addBg("back",  "image", "bg_back.png");

  // ---- Sheet pattern (full sheet image)
  if (state.sheetPattern?.image){
    const blob = await imageToPNGBlob(state.sheetPattern.image);
    const path = `assets/sheet_pattern.png`;
    assets.push({path, blob});
    manifest.sheetPattern = manifest.sheetPattern || {};
    manifest.sheetPattern.image = path;
  }

  // ---- Per-card backs
  if (Array.isArray(state.perCard?.images) && state.perCard.images.length){
    manifest.perCard = manifest.perCard || {enabled:true, images:[]};
    const arr = [];
    for (let i=0;i<state.perCard.images.length;i++){
      const img = state.perCard.images[i];
      if (!img) { arr.push(null); continue; }
      const blob = await imageToPNGBlob(img);
      const path = `assets/percard_${i}.png`;
      assets.push({path, blob});
      arr.push(path);
    }
    manifest.perCard.images = arr;
  }

  // ---- Layer images
  async function addLayerImages(side){
    const layers = state[side].layers || [];
    for (let i=0;i<layers.length;i++){
      const L = layers[i];
      if (L.type === "image" && L.image){
        const blob = await imageToPNGBlob(L.image);
        const path = `assets/${side}_layer_${i}.png`;
        assets.push({path, blob});
        // replace the heavy image object with path
        manifest[side].layers[i].image = path;
      }
    }
  }
  await addLayerImages("front");
  await addLayerImages("back");

  // ---- Custom fonts (expecting you push uploads into state.customFonts[])
  // Each item should look like:
  // { name: "YourFontName", file?: File|Blob, url?: "blob:...", ext?: ".ttf" }
  if (Array.isArray(state.customFonts) && state.customFonts.length){
    manifest.customFonts = [];
    for (let i=0;i<state.customFonts.length;i++){
      const f = state.customFonts[i] || {};
      let blob = null, ext = (f.ext || "").toLowerCase();
      if (f.file instanceof Blob){
        blob = f.file;
        if (!ext){
          if (f.file.name) ext = (f.file.name.match(/\.[a-z0-9]{2,5}$/i)||["",".ttf"])[0];
          else ext = ".ttf";
        }
      } else if (f.url && f.url.startsWith("blob:")){
        // best effort: fetch the blob url
        try{
          blob = await fetch(f.url).then(r=>r.blob());
          if (!ext) ext = ".ttf";
        }catch{}
      } else if (f.dataURL && f.dataURL.startsWith("data:")){
        blob = dataURLToBlob(f.dataURL);
        if (!ext) ext = ".ttf";
      }

      if (blob){
        const nice = cleanName(f.name, `font_${i}`);
        const path = `assets/${nice}${safeExt(nice, ext||".ttf") || ext}`;
        assets.push({path, blob});
        manifest.customFonts.push({ name: f.name, path });
      }
    }
  }

  return {manifest, assets};
}

// ---- EXPORT: make .card (zip) with manifest + assets
btnDownloadProject.addEventListener("click", async ()=>{
  try{
    const JSZip = await ensureJSZip();
    const zip = new JSZip();

    const {manifest, assets} = await buildExportPackage();
    zip.file("project.json", JSON.stringify(manifest, null, 2));

    if (assets.length){
      const folder = zip.folder("assets");
      for (const a of assets) folder.file(a.path.replace(/^assets\//,""), a.blob);
    }

    const blob = await zip.generateAsync({type:"blob", compression:"DEFLATE", compressionOptions:{level:6}});
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `${cleanName(manifest.projectName || "business", "business")}.card`;
    a.click();
    URL.revokeObjectURL(a.href);
  }catch(err){
    console.error(err);
    alert("Export failed. See console for details.");
  }
});

// ---- IMPORT: read .card, rebuild state, reload assets & fonts
async function hydrateStateFromZip(zip){
  const projFile = zip.file("project.json");
  if (!projFile) throw new Error("Not a valid .card file (missing project.json).");
  const manifest = JSON.parse(await projFile.async("string"));

  // helper to load an image from a zip path (returns HTMLImageElement or null)
  async function loadImg(path){
    if (!path) return null;
    const zf = zip.file(path);
    if (!zf) return null;
    const blob = await zf.async("blob");
    return await new Promise(res=>{
      const url = URL.createObjectURL(blob);
      const img = new Image();
      img.onload = ()=>res(img);
      img.onerror = ()=>res(null);
      img.src = url;
    });
  }

  // helper to turn zip path into a blob URL
  async function blobURL(path){
    const zf = zip.file(path);
    if (!zf) return null;
    const blob = await zf.async("blob");
    return URL.createObjectURL(blob);
  }

  // Start from current state to preserve runtime refs, then overwrite
  state = {...state, ...manifest, selection:null};
  state.customFonts = [];
  normalizeFontAliases(state);

  // Backgrounds
  if (manifest.bg?.front?.image) state.bg.front.image = await loadImg(manifest.bg.front.image);
  else state.bg.front.image = null;

  if (manifest.bg?.back?.image) state.bg.back.image  = await loadImg(manifest.bg.back.image);
  else state.bg.back.image = null;

  // Sheet pattern
  if (manifest.sheetPattern?.image){
    state.sheetPattern.image = await loadImg(manifest.sheetPattern.image);
  } else {
    state.sheetPattern.image = null;
  }

  // Per-card backs
  if (Array.isArray(manifest.perCard?.images)){
    const imgs = [];
    for (const p of manifest.perCard.images){
      imgs.push(p ? (await loadImg(p)) : null);
    }
    state.perCard.images = imgs.filter(Boolean);
  } else {
    state.perCard.images = [];
  }

  // Layer images
  async function hydrateLayers(side){
    const layers = state[side].layers || [];
    for (let i=0;i<layers.length;i++){
      const L = layers[i];
      if (L.type === "image" && typeof L.image === "string"){
        L.image = await loadImg(L.image);
      }
    }
  }
  await hydrateLayers("front");
  await hydrateLayers("back");

  // Fonts
  state.customFonts = state.customFonts || [];
  if (Array.isArray(manifest.customFonts)){
    for (const f of manifest.customFonts){
      if (!f?.path) continue;
      const url = await blobURL(f.path);
      if (!url) continue;
      try{
        const face = new FontFace(f.name, `url(${url})`);
        await face.load();
        document.fonts.add(face);
        // Keep a reference so future exports can include it again
        state.customFonts.push({ name:f.name, url, ext: (f.path.match(/\.[a-z0-9]{2,5}$/i)||["",".ttf"])[0] });
      }catch(err){
        console.warn("Font failed to load:", f.name, err);
      }
    }
  }

  // Refresh UI / canvases
  if (document.fonts?.ready){
    try{
      await document.fonts.ready;
    }catch{}
  }
  renderState();
}

btnLoadProject.addEventListener("change", async (e)=>{
  const file = e.target.files[0]; 
  e.target.value = "";
  if (!file) return;

  try{
    const JSZip = await ensureJSZip();
    const zip = await JSZip.loadAsync(file);
    await hydrateStateFromZip(zip);
  }catch(err){
    console.error(err);
    alert("Import failed. This file may be corrupted or not a valid .card.");
  }
});

(btnResetProject) && btnResetProject.addEventListener("click", ()=>{
  const shouldReset = confirm("Reset the project and clear all layers?");
  if (!shouldReset) return;
  state = createEmptyState();
  normalizeFontAliases(state);
  renderState(true);
});

(async function autoLoadDefaultProject(){
  try{
    const resp = await fetch("ACME.card", {cache:"no-cache"});
    if (!resp.ok) return;
    const blob = await resp.blob();
    const JSZip = await ensureJSZip();
    const zip = await JSZip.loadAsync(blob);
    await hydrateStateFromZip(zip);
  }catch(err){
    console.warn("Default project failed to load:", err);
  }
})();

  // ---- Preview tabs ABOVE the canvas (thumbnails) ----
  function renderCanvasTabs(){
    if (!canvasTabs) return;
    canvasTabs.innerHTML = "";
    const btnFront = document.createElement("button");
    btnFront.className = "canvas-tab" + (state.side==="front"?" active":"");
    const imgF = document.createElement("img");
    imgF.style.cssText = "width:64px;height:40px;object-fit:cover;border:1px solid #2f3b55;border-radius:4px;display:block";
    imgF.alt = "Front preview";
    imgF.src = canvasFront.toDataURL("image/png");
    const labelF = document.createElement("div");
    labelF.textContent = "Front";
    labelF.style.cssText = "margin-top:4px;color:#cddbff;font-size:12px;text-align:center";
    btnFront.appendChild(imgF); btnFront.appendChild(labelF);
    btnFront.addEventListener("click", ()=> setSide("front"));

    const btnBack = document.createElement("button");
    btnBack.className = "canvas-tab" + (state.side==="back"?" active":"");
    const imgB = document.createElement("img");
    imgB.style.cssText = "width:64px;height:40px;object-fit:cover;border:1px solid #2f3b55;border-radius:4px;display:block";
    imgB.alt = "Back preview";
    imgB.src = canvasBack.toDataURL("image/png");
    const labelB = document.createElement("div");
    labelB.textContent = "Back";
    labelB.style.cssText = "margin-top:4px;color:#cddbff;font-size:12px;text-align:center";
    btnBack.appendChild(imgB); btnBack.appendChild(labelB);
    btnBack.addEventListener("click", ()=> setSide("back"));

    canvasTabs.appendChild(btnFront);
    canvasTabs.appendChild(btnBack);
  }

  // Init
  function initUIFromState(){
    unitSel.value = state.units;
    dpiInput.value = state.dpi;
    widthInput.value = state.cardWidthMM;
    heightInput.value = state.cardHeightMM;
    bleedInput.value = state.bleedMM;
    cropMarksChk.checked = state.cropMarks;
    gutterInput.value = state.gutterMM;
    pageMarginInput.value = state.pageMarginMM;
    trimBoxesChk.checked = state.trimBoxes;
    showBleedChk.checked = state.showBleed;
    showGridChk.checked = state.grid.show;
    gridSpacingInput.value = state.grid.spacingMM;
    backFlipSel.value = state.backFlip;
  }
  renderState();
  preloadPaperTextures(()=>{
    installPaperSelector();
    // default to white if not set
    applyPaperType(state.paperType || "white");
  });
  // Keyboard helpers
  window.addEventListener("keydown", (e)=>{
    const l = getSelection();
    if(!l) return;
    const active = document.activeElement;
    if (active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA' || active.isContentEditable)) return;

    if (e.key==="Delete"){ e.preventDefault(); btnDelete.click(); return; }
    if (e.key==="[" ){ btnSendBack.click(); return; }
    if (e.key()==="]"){ btnBringFwd.click(); return; }
  });
// --- FONT UPLOAD HANDLER (robust) ---
const fontFamilySelect = document.getElementById("fontFamily");
ensureStandardFontsInSelect(fontFamilySelect);

// inject an upload control next to the Font label if missing
let uploadFontInput = document.getElementById("uploadFont");
if (!uploadFontInput && fontFamilySelect){
  const lbl = fontFamilySelect.closest("label");
  const up = document.createElement("label");
  up.className = "file-btn small";
  up.innerHTML = `Upload Font<input type="file" id="uploadFont" accept=".ttf,.otf,.woff,.woff2">`;
  // place right after the Font label
  lbl.parentElement.insertBefore(up, lbl.nextSibling);
  uploadFontInput = up.querySelector("#uploadFont");
}

if (uploadFontInput){
  uploadFontInput.addEventListener("change", e => {
    const file = e.target.files[0];
    if (!file) return;

    const fontName = file.name.replace(/\.[^/.]+$/, ""); // strip extension
    const url = URL.createObjectURL(file);

    const fontFace = new FontFace(fontName, `url(${url})`);
    fontFace.load().then(loaded => {
      document.fonts.add(loaded);

      // add to dropdown and select it
      const opt = document.createElement("option");
      opt.value = `'${fontName}'`;
      opt.textContent = fontName + " (uploaded)";
      fontFamilySelect.appendChild(opt);
      updateSelectedTextFont(opt.value);
    }).catch(err => {
      alert("Could not load font: " + err);
    });
  });
}

// Arrow-key move support: adds small nudge to the selected item.
// - Uses #units and #dpi controls to convert mm->px when needed.
// - Hold Shift to nudge 10×.
// - Tries several common project hooks: moveSelectedBy(dx,dy),
//   window.selection.moveBy(dx,dy), window.selectedItem.{x,y} update.
// - Fallback: dispatches "move-selected" CustomEvent with {dx,dy}.
(function(){
  window.addEventListener('keydown', function onArrowKey(e){
    const keys = ['ArrowUp','ArrowDown','ArrowLeft','ArrowRight'];
    if (!keys.includes(e.key)) return;

    // don't hijack typing in inputs/textareas/contenteditable
    const active = document.activeElement;
    if (active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA' || active.isContentEditable)) return;

    e.preventDefault();

    // base step: 1px, or 0.5mm in mm-mode
    let stepPx = 1;
    if ((unitSel && unitSel.value === "mm") || state.units === "mm") {
      stepPx = mmToPx(0.5);
    }

    if (e.shiftKey) stepPx *= 10;

    let dx = 0, dy = 0;
    if (e.key === 'ArrowUp') dy = -stepPx;
    if (e.key === 'ArrowDown') dy = stepPx;
    if (e.key === 'ArrowLeft') dx = -stepPx;
    if (e.key === 'ArrowRight') dx = stepPx;

    // Built-in editor hook
    if (moveSelectionBy(dx, dy)) return;

    // Try project hook: moveSelectedBy(dx, dy)
    if (typeof window.moveSelectedBy === 'function') {
      window.moveSelectedBy(dx, dy);
      return;
    }

    // Try selection API: window.selection.moveBy(dx,dy)
    if (window.selection && typeof window.selection.moveBy === 'function') {
      window.selection.moveBy(dx, dy);
      return;
    }

    // Try common global selected item object
    const sel = window.selectedItem || window.selected;
    if (sel && typeof sel.x === 'number' && typeof sel.y === 'number') {
      sel.x += dx;
      sel.y += dy;
      // call common render/update functions if present
      if (typeof window.render === 'function') return window.render();
      if (typeof window.updateCanvas === 'function') return window.updateCanvas();
      return;
    }

    // Fallback: dispatch a custom event others can listen for
    window.dispatchEvent(new CustomEvent('move-selected', { detail: { dx, dy } }));
  });
})();

})();
