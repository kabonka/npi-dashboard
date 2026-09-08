/* ============================================================
 * 範圍截圖 / 長截圖 — 通用腳本（html2canvas 1.4.1）
 * 依賴：html2canvas.min.js（與頁面同目錄，file:// 必須本地引用，不可用 CDN）
 * 依賴：assets/crop-screenshot.html 的 modal 結構（id: cropBk / cropImg /
 *       cropSel / cropImgWrap / crX / crY / crW / crH / crInfo / crStage）
 *
 * 用法（在頁面 JS 裡設定一次）：
 *   CROP_SHOT.opts = {
 *     closeSelectors : ['.msel.open'],        // 拍攝前要關閉的面板（class 移除 'open'）
 *     expandSelectors: ['.wrap'],             // 有 overflow:auto / max-height 的捲動容器 → 拍攝前暫時展開
 *     bannerHTML     : () => '<div id="shot-summary">...</div>',  // 截圖時注入頁面頂端的摘要（可 null）
 *     filePrefix     : 'myreport',            // 下載檔名前綴 → myreport_shot_YYYY-MM-DD.png / myreport_crop_YYYY-MM-DD.png
 *     scale          : 1.5                    // html2canvas 放大倍率（清晰度）
 *   };
 *   <button onclick="takeScreenshot()">📷 長截圖</button>
 *   <button onclick="takeCropScreenshot()">✂️ 範圍截圖</button>
 * ============================================================ */

const CROP_SHOT = {
  opts: {
    closeSelectors : [],      // e.g. ['.msel.open'] 或 [['.ms-panel.show','show'], ['#detailModal.active','active']]
    expandSelectors: [],      // e.g. ['.wrap'] — 內部捲動容器（overflow:auto / max-height）
    bannerHTML     : null,    // () => html string，需以 id="shot-summary" 為根節點
    filePrefix     : 'page',
    scale          : 1.5,
    backgroundColor: '#ffffff'  // 頁面底色（深色主題頁面請設為頁面背景色）
  }
};

let cropCanvas = null;                    // 整頁原始 canvas（未縮放的原始畫素）
const cropSel = {x:0, y:0, w:0, h:0};     // 裁切框（一律用「原始 canvas 畫素」座標）

/* ---- 共用拍攝流程：關面板 → 注入摘要 banner → 展開捲動容器 → 拍整頁 → 還原 UI ---- */
async function captureFullCanvas(){
  const opts = CROP_SHOT.opts;
  if(typeof html2canvas === 'undefined'){
    alert('截圖函式庫未載入：請確認 html2canvas.min.js 與本頁面在同一目錄。');
    return null;
  }
  // 1. 關掉任何開啟中的下拉/彈窗，避免入鏡
  //    條目格式：'selector'（移除 'open' class）或 ['selector','className']
  (opts.closeSelectors || []).forEach(item => {
    const [sel, cls] = Array.isArray(item) ? item : [item, 'open'];
    document.querySelectorAll(sel).forEach(el => el.classList.remove(cls));
  });
  document.getElementById('cropBk').classList.remove('show');
  // 2. 注入摘要 banner（id 必須是 shot-summary，方便之後移除）
  let banner = document.getElementById('shot-summary');
  if(banner) banner.remove();
  if(opts.bannerHTML){
    const tmp = document.createElement('div');
    tmp.innerHTML = opts.bannerHTML();
    const node = tmp.firstElementChild;
    const header = document.querySelector('header');
    if(header && header.nextSibling) document.body.insertBefore(node, header.nextSibling);
    else document.body.insertBefore(node, document.body.firstChild);
  }
  // 3. 暫時解除容器的寬度/高度限制與 overflow 裁切（否則只能拍到可見部分！）
  const saved = [];
  (opts.expandSelectors || []).forEach(sel =>
    document.querySelectorAll(sel).forEach(el => {
      saved.push({el, maxWidth:el.style.maxWidth, maxHeight:el.style.maxHeight, overflow:el.style.overflow});
      el.style.maxWidth = 'none';
      el.style.maxHeight = 'none';
      el.style.overflow = 'visible';
    }));
  window.scrollTo(0, 0);   // 捲回左上角，sticky 表頭才會在正確位置
  // 4. 以完整展開後的捲動範圍拍整頁（含水平溢出）
  const docEl = document.documentElement;
  const fullW = Math.max(docEl.scrollWidth,  document.body.scrollWidth);
  const fullH = Math.max(docEl.scrollHeight, document.body.scrollHeight);
  let canvas = null;
  try{
    canvas = await html2canvas(document.body, {
      width: fullW, height: fullH,
      windowWidth: fullW, windowHeight: fullH,
      scrollX: 0, scrollY: 0, x: 0, y: 0,
      backgroundColor: opts.backgroundColor || '#ffffff',
      scale: opts.scale || 1.5,
      useCORS: true,
      logging: false
    });
  }catch(e){
    console.error(e);
    alert('截圖失敗：' + e.message);
  }finally{
    // 5. 還原 UI
    const rm = document.getElementById('shot-summary');
    if(rm) rm.remove();
    saved.forEach(s => { s.el.style.maxWidth = s.maxWidth; s.el.style.maxHeight = s.maxHeight; s.el.style.overflow = s.overflow; });
  }
  return canvas;
}

/* ---- 拍攝單一元素（給「📸 截此彈窗」這類場景用） ----
   el  : HTMLElement 或 selector 字串
   注入 banner 於元素頂端後拍攝，並依元素 scrollWidth/scrollHeight 抓完整內容。
   攝影期間 elBox 已經在呼叫端先解開 maxHeight/overflow，這裡不再還原。*/
async function captureElementCanvas(el, bannerInEl){
  const opts = CROP_SHOT.opts;
  if(typeof html2canvas === 'undefined'){
    alert('截圖函式庫未載入：請確認 html2canvas.min.js 與本頁面在同一目錄。');
    return null;
  }
  if(typeof el === 'string') el = document.querySelector(el);
  if(!el) return null;
  let banner = document.getElementById('shot-summary');
  if(banner) banner.remove();
  if(opts.bannerHTML){
    const tmp = document.createElement('div');
    tmp.innerHTML = opts.bannerHTML();
    const node = tmp.firstElementChild;
    el.insertBefore(node, el.firstChild);
  }
  const w = Math.max(el.scrollWidth,  el.offsetWidth);
  const h = Math.max(el.scrollHeight, el.offsetHeight);
  let canvas = null;
  try{
    canvas = await html2canvas(el, {
      width: w, height: h,
      windowWidth: w, windowHeight: h,
      scrollX: 0, scrollY: 0, x: 0, y: 0,
      backgroundColor: opts.backgroundColor || '#ffffff',
      scale: opts.scale || 1.5,
      useCORS: true,
      logging: false
    });
  }catch(e){
    console.error(e);
    alert('截圖失敗：' + e.message);
  }finally{
    const rm = document.getElementById('shot-summary');
    if(rm) rm.remove();
  }
  return canvas;
}

/* ---- 按鈕 3：彈窗內「📸 截此彈窗」→ 只拍當前彈窗內容（自動解開 maxHeight 限制） ---- */
async function takeElementShot(selector, filePrefix){
  const el = typeof selector === 'string' ? document.querySelector(selector) : selector;
  if(!el){ alert('找不到目標元素'); return; }
  const btn = document.activeElement && document.activeElement.classList.contains('btn-shot-modal')
              ? document.activeElement : null;
  const old = btn ? btn.textContent : '';
  if(btn){ btn.textContent = '…'; btn.disabled = true; }
  // 暫時解除 max-height / overflow 限制，否則只能拍到可見部分
  const saved = [];
  const unlock = (n) => {
    const cs = getComputedStyle(n);
    if(cs.maxHeight !== 'none' && cs.maxHeight !== '0px'){
      saved.push({el:n, k:'maxHeight', v:n.style.maxHeight});
      n.style.maxHeight = 'none';
    }
    if(cs.overflow !== 'visible' && cs.overflowY !== 'visible'){
      saved.push({el:n, k:'overflow', v:n.style.overflow});
      n.style.overflow = 'visible';
    }
    if(cs.overflowY === 'auto' || cs.overflowY === 'scroll'){
      saved.push({el:n, k:'overflowY', v:n.style.overflowY});
      n.style.overflowY = 'visible';
    }
  };
  unlock(el);
  el.querySelectorAll('*').forEach(unlock);
  let canvas = null;
  try{
    if(btn) btn.textContent = old;
    canvas = await captureElementCanvas(el);
  }finally{
    saved.forEach(s => { s.el.style[s.k] = s.v; });
    if(btn){ btn.textContent = old; btn.disabled = false; }
  }
  if(canvas && filePrefix) downloadCanvas(canvas, `${filePrefix}_shot_${dateTag()}.png`);
  return canvas;
}

/* ---- 按鈕 4：彈窗內「✂️ 範圍截圖」→ 拍彈窗後彈出裁切框讓使用者再選區 ---- */
async function takeElementCrop(selector, filePrefix){
  const el = typeof selector === 'string' ? document.querySelector(selector) : selector;
  if(!el){ alert('找不到目標元素'); return; }
  const btn = document.activeElement && document.activeElement.classList.contains('btn-crop-modal')
              ? document.activeElement : null;
  const old = btn ? btn.textContent : '';
  if(btn){ btn.textContent = '…'; btn.disabled = true; }
  const saved = [];
  const unlock = (n) => {
    const cs = getComputedStyle(n);
    if(cs.maxHeight !== 'none' && cs.maxHeight !== '0px'){
      saved.push({el:n, k:'maxHeight', v:n.style.maxHeight});
      n.style.maxHeight = 'none';
    }
    if(cs.overflow !== 'visible' && cs.overflowY !== 'visible'){
      saved.push({el:n, k:'overflow', v:n.style.overflow});
      n.style.overflow = 'visible';
    }
    if(cs.overflowY === 'auto' || cs.overflowY === 'scroll'){
      saved.push({el:n, k:'overflowY', v:n.style.overflowY});
      n.style.overflowY = 'visible';
    }
  };
  unlock(el);
  el.querySelectorAll('*').forEach(unlock);
  let canvas = null;
  try{
    if(btn) btn.textContent = old;
    canvas = await captureElementCanvas(el);
  }finally{
    saved.forEach(s => { s.el.style[s.k] = s.v; });
    if(btn){ btn.textContent = old; btn.disabled = false; }
  }
  if(!canvas) return;
  cropCanvas = canvas;
  document.getElementById('cropBk').classList.add('show');
  document.getElementById('cropStage').scrollTop  = 0;
  document.getElementById('cropStage').scrollLeft = 0;
  const img = document.getElementById('cropImg');
  const applyCenterCrop = () => {
    const W = canvas.width, H = canvas.height;
    cropSel.x = 0; cropSel.y = 0;
    cropSel.w = W;  cropSel.h = H;
    cropDraw();
  };
  img.onload = applyCenterCrop;
  img.src = canvas.toDataURL('image/png');
  try{ await img.decode(); applyCenterCrop(); } catch(e){ /* fallback: onload */ }
}

/* ---- 按鈕 1：長截圖（整頁直接下載） ---- */
async function takeScreenshot(){
  const btn = document.querySelector('.btn-shot');
  const old = btn ? btn.textContent : '';
  if(btn){ btn.textContent = '截圖中…'; btn.disabled = true; }
  let canvas = null;
  try{
    // 拍攝前先還原按鈕文字，避免「截圖中…」被拍進結果圖
    if(btn) btn.textContent = old;
    canvas = await captureFullCanvas();
  }
  finally{ if(btn){ btn.textContent = old; btn.disabled = false; } }
  if(!canvas) return;
  downloadCanvas(canvas, `${CROP_SHOT.opts.filePrefix}_shot_${dateTag()}.png`);
}

/* ---- 按鈕 2：範圍截圖（先拍整頁 → 開裁切框 modal） ---- */
async function takeCropScreenshot(){
  const btn = document.querySelector('.btn-crop');
  const old = btn ? btn.textContent : '';
  if(btn){ btn.textContent = '截圖中…'; btn.disabled = true; }
  let canvas = null;
  try{
    if(btn) btn.textContent = old;
    canvas = await captureFullCanvas();
  }
  finally{ if(btn){ btn.textContent = old; btn.disabled = false; } }
  if(!canvas) return;
  cropCanvas = canvas;
  document.getElementById('cropBk').classList.add('show');
  document.getElementById('cropStage').scrollTop  = 0;
  document.getElementById('cropStage').scrollLeft = 0;
  const img = document.getElementById('cropImg');
  /* 預設選取 = 畫布完整寬度 + 完整高度 */
  const applyCenterCrop = () => {
    const W = canvas.width, H = canvas.height;
    cropSel.x = 0;
    cropSel.y = 0;
    cropSel.w = W;
    cropSel.h = H;
    cropDraw();
  };
  img.onload = applyCenterCrop;
  img.src = canvas.toDataURL('image/png');
  try{ await img.decode(); applyCenterCrop(); } catch(e){ /* fallback: onload */ }
}

/* ---- 顯示比例：img 顯示寬 / 原始 canvas 寬 ---- */
function cropK(){
  const img = document.getElementById('cropImg');
  return cropCanvas ? img.getBoundingClientRect().width / cropCanvas.width : 1;
}

/* ---- 重畫裁切框（clamp：≥10px、不出界；同步數值欄，正在打字的欄位不覆寫） ---- */
function cropDraw(){
  if(!cropCanvas) return;
  const selEl = document.getElementById('cropSel');
  const k = cropK();
  cropSel.w = Math.min(Math.max(10, Math.round(cropSel.w)), cropCanvas.width);
  cropSel.h = Math.min(Math.max(10, Math.round(cropSel.h)), cropCanvas.height);
  cropSel.x = Math.min(Math.max(0, Math.round(cropSel.x)), cropCanvas.width  - cropSel.w);
  cropSel.y = Math.min(Math.max(0, Math.round(cropSel.y)), cropCanvas.height - cropSel.h);
  selEl.style.left   = (cropSel.x * k) + 'px';
  selEl.style.top    = (cropSel.y * k) + 'px';
  selEl.style.width  = (cropSel.w * k) + 'px';
  selEl.style.height = (cropSel.h * k) + 'px';
  const ae = document.activeElement;
  const setv = (id, v) => { const el = document.getElementById(id); if(el !== ae) el.value = v; };
  setv('crX', cropSel.x); setv('crY', cropSel.y);
  setv('crW', cropSel.w); setv('crH', cropSel.h);
  document.getElementById('crInfo').textContent =
    `原圖 ${cropCanvas.width}×${cropCanvas.height} px · 裁切 ${cropSel.w}×${cropSel.h} px（原圖畫素）`;
}

function cropSelectAll(){
  if(!cropCanvas) return;
  cropSel.x = 0; cropSel.y = 0;
  cropSel.w = cropCanvas.width; cropSel.h = cropCanvas.height;
  cropDraw();
}

function cropClose(){
  document.getElementById('cropBk').classList.remove('show');
  document.getElementById('cropImg').src = '';
  cropCanvas = null;
}

function cropDownload(){
  if(!cropCanvas) return;
  const out = document.createElement('canvas');
  out.width = cropSel.w; out.height = cropSel.h;
  out.getContext('2d').drawImage(cropCanvas,
    cropSel.x, cropSel.y, cropSel.w, cropSel.h,
    0, 0, cropSel.w, cropSel.h);
  downloadCanvas(out, `${CROP_SHOT.opts.filePrefix}_crop_${dateTag()}.png`);
}

function downloadCanvas(canvas, fname){
  const a = document.createElement('a');
  a.href = canvas.toDataURL('image/png');
  a.download = fname;
  a.click();
}

function dateTag(){ return new Date().toISOString().slice(0, 10); }

/* ---- 裁切框拖曳：握把 = 8 方向縮放；框內 = 整框移動；框外 = 重拉新框 ---- */
document.addEventListener('DOMContentLoaded', function initCropDrag(){
  const wrapEl = document.getElementById('cropImgWrap');
  if(!wrapEl) return;
  let drag = null;
  wrapEl.addEventListener('pointerdown', e => {
    if(!cropCanvas) return;
    const handle = e.target.closest('.crp-h');
    let mode;
    if(handle) mode = handle.dataset.d;
    else if(e.target.closest('#cropSel')) mode = 'move';
    else{
      const r = wrapEl.getBoundingClientRect();
      const k = cropK();
      cropSel.x = (e.clientX - r.left) / k;
      cropSel.y = (e.clientY - r.top)  / k;
      cropSel.w = 10; cropSel.h = 10;
      mode = 'se';   // 從空白處起手 = 往右下拉新框
    }
    drag = {mode, sx:e.clientX, sy:e.clientY, ox:cropSel.x, oy:cropSel.y, ow:cropSel.w, oh:cropSel.h};
    wrapEl.setPointerCapture(e.pointerId);
    e.preventDefault();
    cropDraw();
  });
  wrapEl.addEventListener('pointermove', e => {
    if(!drag || !cropCanvas) return;
    const k = cropK();
    const dx = (e.clientX - drag.sx) / k;
    const dy = (e.clientY - drag.sy) / k;
    let x = drag.ox, y = drag.oy, w = drag.ow, h = drag.oh;
    const m = drag.mode;
    if(m === 'move'){ x += dx; y += dy; }
    else{
      if(m.includes('e')) w = drag.ow + dx;
      if(m.includes('s')) h = drag.oh + dy;
      if(m.includes('w')){ x = drag.ox + dx; w = drag.ow - dx; }
      if(m.includes('n')){ y = drag.oy + dy; h = drag.oh - dy; }
      if(w < 0){ x += w; w = -w; }   // 拖過對邊 → 自動翻轉
      if(h < 0){ y += h; h = -h; }
    }
    cropSel.x = x; cropSel.y = y; cropSel.w = w; cropSel.h = h;
    cropDraw();
  });
  const end = () => { drag = null; };
  wrapEl.addEventListener('pointerup', end);
  wrapEl.addEventListener('pointercancel', end);
});

/* ---- 數值輸入框：直接鍵入 X / Y / 寬 / 高（原圖畫素）微調 ---- */
['crX','crY','crW','crH'].forEach(id => {
  const el = document.getElementById(id);
  if(!el) return;
  el.addEventListener('input', () => {
    if(!cropCanvas) return;
    const v = id2 => { const n = parseFloat(document.getElementById(id2).value); return isNaN(n) ? 0 : n; };
    cropSel.x = v('crX'); cropSel.y = v('crY');
    cropSel.w = v('crW') || 10; cropSel.h = v('crH') || 10;
    cropDraw();
  });
});

/* 視窗縮放後顯示比例改變 → 重畫裁切框 */
window.addEventListener('resize', () => { if(cropCanvas) cropDraw(); });
