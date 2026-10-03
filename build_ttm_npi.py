#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build TTM-style dashboard for NPI_vx from Spec總表.xlsx.

Outputs a single self-contained HTML (TTM_dashboard.html) with embedded data
and inlined JS libraries for Excel export + screenshot.

-------------------------------------------------------------------------------
要修改 TTM_dashboard.html，就是改這支程式，然後重新生成：

    python build_ttm_npi.py
    （或直接雙擊「一键生成.bat」/「一键生成2.bat」，兩支的 STEP 6 都會跑這支）

**不要直接編輯 TTM_dashboard.html** —— 那是產物，下次重新生成會被整份覆蓋。

改動時的三個原則
  1. 大段 CSS / JS 不要直接塞進 f-string（裡面的 { } 都要寫成 {{ }}）。
     請另外寫成模組層級常數（r'''...'''），再用
     html.replace("/*__XXX__*/", 常數) 注入；template 裡已預留
     __CROPCSS__ / __CROPMODAL__ / __REPORTCSS__ / __REPORTMODAL__ /
     __LIB__ / __SHOTLIB__ / __EXPORTJS__ / __REPORTJS__ 等 placeholder。
  2. 資料由 build_merged_rows() 依 (Model + Kickoff..MP 時間軸) 分群合併，
     每列會帶 _MKT/_CPU/_GPU/_NPM/_Stage/_BTO ready/_CurrentStatus/_Highlight
     等「底線陣列」。HTML 欄位與底線陣列都必須由 dedup_values() 產生，
     兩者要同步，否則畫面與匯出會不一致。
  3. 改完務必用瀏覽器實際驗證（file:// 直接開即可），
     尤其是週報（openReport / rpSaveExcel / rpShot）與截圖模式。

輸入 / 輸出
  輸入：Spec總表.xlsx（工作表 Schedule）
  輸出：TTM_dashboard.html
  依賴：openpyxl；lib/ 底下的 jszip.js、xlsx.js、html2canvas.min.js、
        crop-screenshot.{js,css,html}（缺檔會在這裡就報錯，不會產生半成品）
-------------------------------------------------------------------------------
"""
import os
import re
import json
from datetime import datetime, date
from collections import defaultdict

import openpyxl
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SPEC_PATH = os.path.join(BASE_DIR, "Spec總表.xlsx")
OUTPUT_HTML = os.path.join(BASE_DIR, "TTM_dashboard.html")
LIB_DIR = os.path.join(BASE_DIR, "lib")

LIB_JSZIP = os.path.join(LIB_DIR, "jszip.js")
LIB_XLSX = os.path.join(LIB_DIR, "xlsx.js")
LIB_H2C = os.path.join(LIB_DIR, "html2canvas.min.js")
LIB_CROPJS = os.path.join(LIB_DIR, "crop-screenshot.js")
LIB_CROPCSS = os.path.join(LIB_DIR, "crop-screenshot.css")
LIB_CROPHTML = os.path.join(LIB_DIR, "crop-screenshot.html")

# Stage 篩選下拉的「選項」一律取自 Spec總表.xlsx 的實際值，Spec總表.xlsx 改了
# 選項就會跟著變（不寫死、不會無中生有）。DEFAULT_STAGES 只決定「開頁面時預設
# 勾選哪幾個」，語意是：
#   資料裡真的有、且列在這份清單裡的值 → 預設勾選
#   資料裡有、但不在這份清單裡的值     → 預設不勾（仍可手動勾）
#   清單裡有、但資料裡沒有的值         → 不會出現在下拉（清單只是偏好）
# 比對採**字串完全比對**：資料裡的 ATS 階段實際值是 'ATS-'（尾端有連字號），
#   所以清單寫 'ATS' 不會勾到 'ATS-'。要連 ATS- 一起預設勾，就把 "ATS-" 也寫進來。
DEFAULT_STAGES = ["Study", "Design", "DVT", "EVT", "MVT", "ATS"]
DEFAULT_MPYEARS = ["2026", "2027"]

HTML_MILESTONE_KEYS = {
    "Kick off": "Kickoff",
    "DVT-start": "DVT-start",
    "EVT-start": "EVT-start",
    "MVT-start": "MVT-start",
    "ATS-start": "ATS-start",
    "MP / 入庫": "MP",
}

SPEC_MILESTONE_COLS = {
    "Kickoff": "K",
    "DVT-start": "L",
    "EVT-start": "M",
    "MVT-start": "N",
    "BTO ready": "O",
    "ATS-start": "P",
    "MP": "Q",
}


def read_text(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


def normalize_date(v):
    """Return ISO date string yyyy-mm-dd or None."""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    s = str(v).strip()
    if not s:
        return None
    for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    return s  # fallback to raw string


def format_date_yyyy_mm_dd(v):
    d = normalize_date(v)
    if d is None:
        return ""
    return d.replace("-", "/")


def extract_year(v):
    d = normalize_date(v)
    if d and len(d) >= 4:
        try:
            return str(int(d[:4]))
        except ValueError:
            pass
    return ""


def escape_html(s):
    return (str(s)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))


def norm_key(s):
    """Normalisation key for de-duplication.

    Whitespace is dropped entirely so purely typographic variants such as
    "X: 3.1 G/O" and "X:3.1 G/O" (a common data-entry slip in the spec sheet)
    collapse into one entry. Meaningful differences (e.g. "NA" vs "N/A")
    are preserved.
    """
    return re.sub(r"\s+", "", s)


def dedup_values(values):
    """Collapse values that are identical after normalisation.

    Order is preserved and the first occurrence's original text is kept.
    """
    out = []
    keys = set()
    for v in values:
        s = str(v).strip() if v is not None else ""
        k = norm_key(s)
        if k and k not in keys:
            keys.add(k)
            out.append(s)
    return out


def merge_field(values, dedup=False):
    """Return HTML cell content for a list of values.

    When dedup=True, repeated values are collapsed (order preserved) so the
    cell shows each unique value only once. Used for Current Status / Highlight.
    """
    stripped = [str(v).strip() if v is not None else "" for v in values]
    if dedup:
        stripped = dedup_values(stripped)
    if not any(stripped):
        return ""
    if len(stripped) == 1:
        return escape_html(stripped[0])
    parts = []
    for i, s in enumerate(stripped, 1):
        display = s if s else "(空白)"
        parts.append(f'<div class="merged-item"><span class="seq-badge">{i}</span>{escape_html(display)}</div>')
    return "".join(parts)


def build_merged_rows(rows):
    """Group rows by (Model + full Kickoff..MP timeline) and merge differing fields."""
    groups = []
    seen = {}
    for r in rows:
        key = (r["Model"], r["Kickoff"], r["DVT-start"], r["EVT-start"],
               r["MVT-start"], r["BTO ready"], r["ATS-start"], r["MP"])
        if key not in seen:
            seen[key] = len(groups)
            groups.append([])
        groups[seen[key]].append(r)

    merged = []
    for grp in groups:
        base = grp[0].copy()
        for field in ["MKT", "CPU", "GPU", "NPM", "Stage", "BTO ready", "CurrentStatus", "Highlight"]:
            vals = [r[field] for r in grp]
            do_dedup = field in ("CPU", "NPM", "Stage", "BTO ready", "CurrentStatus", "Highlight")
            base[field] = merge_field(vals, dedup=do_dedup)
            if do_dedup:
                base[f"_{field}"] = dedup_values(vals)
            else:
                base[f"_{field}"] = [str(v).strip() if v is not None else "" for v in vals]
        base["_Model"] = [base["Model"]]
        base["_MPYear"] = [base["MPYear"]]
        base["_Stage"] = sorted(list(set(str(r["Stage"]).strip() if r["Stage"] is not None else "" for r in grp)))
        base["_MPYear"] = sorted(list(set(str(r["MPYear"]).strip() if r["MPYear"] is not None else "" for r in grp)))
        merged.append(base)
    return merged


EXPORT_JS = r'''
/* ===================== 匯出 Excel（格式同 HTML） ===================== */
const MERGE_FIELDS = ['MKT', 'CPU', 'GPU', 'NPM', 'Stage', 'CurrentStatus', 'Highlight'];
const XLSX_HEAD_BG = 'FF1F8A4C';
const XLSX_HEAD_FG = 'FFFFFFFF';
// 註：原本有 XLSX_MERGE_BG = 'FFE8F5EC'，用來把「多值合併格」填成淺綠。
// 2026-10-03 使用者要求「下載 excel 時不用顏色」，已移除填色與這個常數。
// 資料列的底色一律保持白色，只留框線、欄寬與列高。
const XLSX_BORDER = {
  top: { style: 'thin', color: { rgb: 'FFD8E3DC' } },
  bottom: { style: 'thin', color: { rgb: 'FFD8E3DC' } },
  left: { style: 'thin', color: { rgb: 'FFD8E3DC' } },
  right: { style: 'thin', color: { rgb: 'FFD8E3DC' } }
};

function activeColumns() {
  const cols = [
    ['Model', 'Model name', 7],
    ['MKT', 'MKT name', 18],
    ['CPU', 'CPU', 12],
    ['GPU', 'GPU', 12],
    ['NPM', 'NPM', 4],
    ['Stage', 'Stage', 5]
  ];
  if (midVisible) {
    cols.push(['Kickoff', 'Kickoff', 18]);
    cols.push(['DVT-start', 'DVT-start', 4]);
    cols.push(['EVT-start', 'EVT-start', 4]);
    cols.push(['MVT-start', 'MVT-start', 4]);
    cols.push(['ATS-start', 'ATS-start', 4]);
  } else {
    cols.push(['Kickoff', 'Kickoff', 18]);
  }
  cols.push(['MP', 'MP', 18]);
  cols.push(['CurrentStatus', 'Current Status', 45]);
  cols.push(['Highlight', 'Highlight', 45]);
  return cols;
}

function cellText(r, key) {
  const arr = r['_' + key];
  if (Array.isArray(arr) && arr.length) {
    if (arr.length === 1) return arr[0] || '';
    return arr.map((v, i) => (i + 1) + '. ' + (v || '(空白)')).join('\n');
  }
  return r[key] || '';
}

function cellLineCount(r, key) {
  const arr = r['_' + key];
  if (Array.isArray(arr) && arr.length > 1) return arr.length;
  return 1;
}

function downloadExcel() {
  if (typeof XLSX === 'undefined') { alert('Excel 匯出函式庫未載入。'); return; }
  const cols = activeColumns();
  const aoa = [cols.map(c => c[1])];
  LAST_FILTERED.forEach(r => {
    aoa.push(cols.map(c => {
      const k = c[0];
      if (MERGE_FIELDS.indexOf(k) >= 0) return cellText(r, k);
      return r[k] || '';
    }));
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = cols.map(c => ({ wch: c[2] }));

  function gc(addr) {
    if (!ws[addr]) ws[addr] = { t: 's', v: '' };
    ws[addr].s = ws[addr].s || {};
    return ws[addr];
  }
  function setS(addr, o) {
    gc(addr).s = Object.assign(gc(addr).s, o);
  }
  const center = { horizontal: 'center', vertical: 'center', wrapText: true };
  const left = { horizontal: 'left', vertical: 'top', wrapText: true };

  for (let c = 0; c < cols.length; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c: c });
    gc(addr).s.fill = { patternType: 'solid', fgColor: { rgb: XLSX_HEAD_BG } };
    setS(addr, {
      font: { bold: true, color: { rgb: XLSX_HEAD_FG } },
      alignment: center,
      border: XLSX_BORDER
    });
  }
  for (let i = 0; i < LAST_FILTERED.length; i++) {
    const r = LAST_FILTERED[i];
    const rr = i + 1;
    let maxLines = 1;
    for (let c = 0; c < cols.length; c++) {
      const addr = XLSX.utils.encode_cell({ r: rr, c: c });
      const k = cols[c][0];
      gc(addr).s.border = XLSX_BORDER;
      setS(addr, { alignment: left });
      if (MERGE_FIELDS.indexOf(k) >= 0) {
        // 只算列高，不上色。
        // 原本 n>1 會填 XLSX_MERGE_BG（淺綠 #E8F5EC）標示「多值合併格」，
        // 但 2026-10-03 使用者要求「下載 excel 時不用顏色」，所以移除填色。
        // 多值的資訊仍可從內容看出（cellText 會寫成 "1. xxx / 2. yyy" 多行）。
        maxLines = Math.max(maxLines, cellLineCount(r, k));
      }
    }
    if (!ws['!rows']) ws['!rows'] = [];
    ws['!rows'][rr] = { hpt: Math.max(18, maxLines * 15) };
  }
  if (!ws['!rows']) ws['!rows'] = [];
  ws['!rows'][0] = { hpt: 22 };

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'TTM Dashboard');
  const tag = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  XLSX.writeFile(wb, 'TTM_dashboard_' + tag + '.xlsx');
}

/* ===================== 長截圖 / 範圍截圖 ===================== */
function filterSummary() {
  const parts = [];
  for (const key in FILTER_META) {
    const label = FILTER_META[key];
    const total = document.querySelectorAll('.cb-' + key).length;
    const sel = Array.from(document.querySelectorAll('.cb-' + key + ':checked'))
      .map(cb => cb.value || '(空白)');
    if (!total) continue;
    if (sel.length === total) parts.push(label + '：全部');
    else if (!sel.length) parts.push(label + '：未選');
    else parts.push(label + '：' + sel.join(', '));
  }
  const s = document.getElementById('mp-start').value;
  const e = document.getElementById('mp-end').value;
  if (s || e) parts.push('MP 範圍：' + (s || '不限') + ' ~ ' + (e || '不限'));
  return parts.join('　｜　');
}

if (typeof CROP_SHOT !== 'undefined') {
  CROP_SHOT.opts = {
    closeSelectors: [['.mf-menu.open', 'open']],
    expandSelectors: [],
    filePrefix: 'TTM_dashboard',
    scale: 1.5,
    backgroundColor: '#f6faf7',
    bannerHTML: function () {
      const stats = document.getElementById('stats');
      return '<div id="shot-summary">'
        + '<div class="ss-title">📷 TTM Dashboard　' + new Date().toLocaleString() + '</div>'
        + '<div class="ss-row">' + escapeHtml(stats ? stats.textContent : '') + '</div>'
        + '<div class="ss-row">' + escapeHtml(filterSummary()) + '</div>'
        + '</div>';
    }
  };
}
'''


REPORT_CSS = r'''
/* ============================================================
 * 週報（Email 格式）產生器
 * 單一方案：Dear All, / Update XXX weekly status below. / 兩欄垂直表
 * 多方案  ：標題列 + 方案名稱列 + N 個方案欄（可跨 Plan 合併格）
 * ============================================================ */
.rp-bk { display: none; position: fixed; inset: 0; background: rgba(0,0,0,.6); z-index: 300; overflow: auto; padding: 22px; }
.rp-bk.show { display: block; }
.rp-shell { max-width: 1180px; margin: 0 auto; background: #fff; border-radius: 8px; box-shadow: 0 12px 44px rgba(0,0,0,.38); }
.rp-head { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; padding: 9px 14px; border-bottom: 1px solid #ddd; background: #fff; font-size: 12px; }
.rp-head b { color: var(--green-d); font-size: 13px; margin-right: 4px; }
.rp-lbl { color: #666; font-size: 11px; }
.rp-num { width: 62px; padding: 3px 6px; border: 1px solid #ccc; border-radius: 4px; font-size: 12px; font-family: Consolas, Menlo, monospace; }
.rp-file { width: 230px; padding: 3px 6px; border: 1px solid #ccc; border-radius: 4px; font-size: 12px; }
.rp-info { color: #888; font-size: 11px; margin-left: auto; }
.rp-head button { padding: 4px 12px; font-size: 12px; border: 1px solid var(--line); border-radius: 4px; background: #fff; cursor: pointer; }
.rp-head button:hover { background: var(--green-l); }
.rp-head button.rp-pri { background: var(--green); color: #fff; border-color: var(--green); }
.rp-head button.rp-pri:hover { background: var(--green-d); }
.rp-head button:disabled { opacity: .6; cursor: default; }
.rp-fields { display: flex; flex-wrap: wrap; gap: 4px 12px; align-items: center; padding: 8px 14px; border-bottom: 1px solid #ddd; background: #fbfdfb; font-size: 12px; }
.rp-fields .rp-fl { font-weight: 700; color: var(--green-d); }
.rp-fitem { display: inline-flex; align-items: center; gap: 4px; cursor: pointer; white-space: nowrap; }
.rp-fbtns { display: inline-flex; gap: 6px; margin-left: 6px; }
.rp-fbtns button { padding: 2px 8px; font-size: 11px; border: 1px solid var(--line); border-radius: 4px; background: #fff; color: var(--green-d); cursor: pointer; }
.rp-fbtns button:hover { background: var(--green-l); }
.rp-hint { padding: 6px 14px; font-size: 11px; color: #7a7a7a; border-bottom: 1px solid #eee; background: #fcfcfc; }
.rp-body { padding: 0; }

/* ---- 報表本體（只有這塊會被長截圖／另存新檔） ---- */
#rpPage { background: #fff; padding: 26px 30px; color: #000; font-size: 13px; line-height: 1.55; }
#rpPage .rp-greet { margin: 0 0 12px; }
#rpPage .rp-block { margin: 0 0 30px; }
#rpPage .rp-block:last-child { margin-bottom: 0; }
#rpPage .rp-subject { margin: 0 0 10px; }
#rpPage .rp-empty { color: #999; }
#rpPage table.rp-table { width: 100%; border-collapse: collapse; table-layout: fixed; border: 0; border-radius: 0; box-shadow: none; background: #fff; margin: 0; }
#rpPage table.rp-table tr,
#rpPage table.rp-table tr:nth-child(even),
#rpPage table.rp-table tr:hover { background: #fff; }
#rpPage table.rp-table td { border: 1px solid #444; padding: 6px 9px; vertical-align: top; background: #fff; white-space: pre-wrap; word-break: break-word; font-weight: 400; }
#rpPage table.rp-table td.rp-k { width: 150px; font-weight: 700; position: relative; }
#rpPage table.rp-table td.rp-title { background: #fff2cc; font-weight: 400; }
#rpPage table.rp-table td.rp-plan { background: #cfe2f3; text-align: center; vertical-align: middle; font-weight: 700; position: relative; }
#rpPage table.rp-table td.rp-v { position: relative; }
#rpPage table.rp-table td.rp-v-m { padding-left: 24px; }
#rpPage .rp-ktxt, #rpPage .rp-ptxt { outline: none; min-height: 1.2em; }
#rpPage .rp-edit { outline: none; min-height: 1.35em; }
#rpPage .rp-edit:empty::before { content: attr(data-ph); color: #c9c9c9; }
/* 格內左上角的合併勾選框 */
#rpPage .rp-mcb2 { position: absolute; left: 5px; top: 7px; margin: 0; cursor: pointer; }
#rpPage .rp-hide { position: absolute; right: 2px; top: 2px; width: 18px; height: 18px; padding: 0; line-height: 15px; font-size: 11px; border: 1px solid #d4a7a2; border-radius: 3px; background: #fdf0f0; color: #b42318; cursor: pointer; display: none; }
#rpPage.rp-hasmerge .rp-hide { display: block; }
/* 截圖／匯出時把所有編輯用控制項藏起來，且還原格內縮排 */
#rpPage.rp-capturing .rp-hide,
#rpPage.rp-capturing .rp-mcb2 { display: none !important; }
#rpPage.rp-capturing table.rp-table td.rp-v-m { padding-left: 9px; }
#rpPage.rp-capturing .rp-edit:empty::before { content: ''; }
#rpHidden button { padding: 1px 8px; font-size: 11px; border: 1px solid var(--line); border-radius: 4px; background: #fff; color: var(--green-d); cursor: pointer; }
#rpHidden button:hover { background: var(--green-l); }
.rp-clear { padding: 1px 8px; font-size: 11px; border: 1px solid var(--line); border-radius: 4px; background: #fff; color: var(--green-d); cursor: pointer; margin: 0 4px; }
.rp-clear:hover { background: var(--green-l); }
'''


REPORT_MODAL = r'''
<!-- ============================================================
  週報（Email 格式）產生器 — 貼到 </body> 之前
============================================================ -->
<div class="rp-bk" id="rpBk">
  <div class="rp-shell">
    <div class="rp-head">
      <b>📧 週報格式（Email）</b>
      <span class="rp-lbl">Model 前綴</span><input type="text" id="rpPrefix" class="rp-num" style="width:56px" value="MS" onchange="rpBuild()">
      <span class="rp-lbl">檔名</span><input type="text" id="rpFileName" class="rp-file">
      <span class="rp-info" id="rpInfo"></span>
      <button type="button" class="rp-pri" onclick="rpShot(this)">📷 長截圖</button>
      <button type="button" class="rp-pri" onclick="rpSaveExcel()">💾 另存新檔（Excel）</button>
      <button type="button" onclick="rpSave()">另存 HTML</button>
      <button type="button" onclick="rpClose()">關閉</button>
    </div>
    <div class="rp-fields" id="rpFields"></div>
    <div class="rp-hint">表格內容與欄位名可直接點擊修改；方案名稱可直接改，<b>✕</b> 可隱藏方案。<b>Current Status / Highlight / Others</b> 每格左上角可勾選要合併的<b>相鄰</b>方案（只勾相鄰兩格就會合併成一格，中間跳號會自動補滿；取消勾選即拆回）。<b>另存新檔</b>會存成 Excel（.xlsx），顏色框線與畫面一致。<button type="button" class="rp-clear" onclick="rpClearMerges()">清除全部合併</button><span id="rpHidden"></span></div>
    <div class="rp-body" id="rpBody"><div id="rpPage"></div></div>
  </div>
</div>
'''


REPORT_JS = r'''
/* ===================== 週報（Email 格式）產生器 =====================
 * 資料來源：目前篩選結果 LAST_FILTERED
 *   同一個 Model 只有一種排程 → 單欄（Dear All, / Update XXX weekly status below.）
 *   同一個 Model 有多種排程 → 多欄（標題列 + 方案名稱列 + N 個方案欄）
 * 可勾選欄位、可直接編輯內容與欄位名、可跨 Plan 合併格、可長截圖或另存 HTML。
 * ================================================================= */
const RP_FIELDS = [
  { id: 'Model',         label: 'Model Name' },
  { id: 'MarkingName',   label: 'MKT Name' },
  { id: 'CPU',           label: 'CPU' },
  { id: 'GPU',           label: 'GPU' },
  { id: 'Stage',         label: 'Stage' },
  { id: 'Kickoff',       label: 'Kick off' },
  { id: 'DVT-start',     label: 'DVT-start' },
  { id: 'EVT-start',     label: 'EVT-start' },
  { id: 'MVT-start',     label: 'MVT-start' },
  { id: 'BTO',           label: 'BTO ready' },
  { id: 'ATS-start',     label: 'ATS-start' },
  { id: 'MP',            label: 'MP' },
  { id: 'CurrentStatus', label: 'Current Status' },
  { id: 'Highlight',     label: 'Highlight' },
  { id: 'Others',        label: 'Others' }
];

/* 只有這三個欄位提供「跨 Plan 合併格」 */
const RP_MERGEABLE = ['CurrentStatus', 'Highlight', 'Others'];

const RP_LABEL_EDITS = {};   // 使用者改過的欄位名（key: field id）
const RP_TITLE_EDITS = {};   // 使用者改過的標題（key: model）
const RP_PLAN_NAMES  = {};   // 使用者改過的方案名稱（key: model#idx）
const RP_PLAN_EDITED = {};   // 方案名稱是否被使用者改過
const RP_PLAN_HIDE   = {};   // 被隱藏的方案（key: model#idx）
const RP_MERGE_SEL   = {};   // 跨 Plan 合併格：key = model#fieldId → 被勾選的方案原始索引陣列（必相鄰）
let RP_PLAN_SIG = '';

function rpField(id) {
  return RP_FIELDS.filter(function (f) { return f.id === id; })[0];
}

function stripTags(s) {
  return String(s == null ? '' : s).replace(/<[^>]*>/g, '').trim();
}

function rpUniq(arr) {
  const out = [];
  const keys = [];
  (arr || []).forEach(function (v) {
    const s = String(v == null ? '' : v).trim();
    const k = s.replace(/\s+/g, '');
    if (s && keys.indexOf(k) < 0) { keys.push(k); out.push(s); }
  });
  return out;
}

/* 多值欄位 → 每個值一行（純文字，去掉儀表板的序號徽章） */
function rpLines(arr, fallback) {
  if (Array.isArray(arr) && arr.length) {
    const lines = arr.map(function (v) { return String(v == null ? '' : v).trim(); })
      .filter(function (v) { return v !== ''; });
    if (lines.length) {
      return lines.map(function (v) { return '<div>' + escapeHtml(v) + '</div>'; }).join('');
    }
  }
  return escapeHtml(stripTags(fallback));
}

function rpValue(r, id, prefix) {
  if (id === 'Model') return escapeHtml(prefix + (r.Model || ''));
  if (id === 'MarkingName') {
    const mkt = r['_MKT'] || [];
    const gpu = r['_GPU'] || [];
    const lines = [];
    for (let i = 0; i < mkt.length; i++) {
      const m = String(mkt[i] == null ? '' : mkt[i]).trim();
      const g = gpu[i] ? ' (' + String(gpu[i]).trim() + ')' : '';
      if (m || g) lines.push('<div>' + escapeHtml(m + g) + '</div>');
    }
    if (lines.length) return lines.join('');
    return escapeHtml(stripTags(r.MKT));
  }
  if (id === 'GPU') return rpLines(rpUniq(r['_GPU']), r.GPU);
  if (id === 'BTO') return rpLines(r['_BTO ready'], r['BTO ready']);
  if (id === 'CPU') return rpLines(r['_CPU'], r.CPU);
  if (id === 'Stage') return rpLines(r['_Stage'], r.Stage);
  if (id === 'CurrentStatus') return rpLines(r['_CurrentStatus'], r.CurrentStatus);
  if (id === 'Highlight') return rpLines(r['_Highlight'], r.Highlight);
  if (id === 'Others') return '';
  return escapeHtml(stripTags(r[id]));
}

/* ---------------- 欄位加選 ---------------- */
function rpRenderFields() {
  const box = document.getElementById('rpFields');
  if (!box) return;
  box.innerHTML = '<span class="rp-fl">欄位加選：</span>'
    + RP_FIELDS.map(function (f) {
        return '<label class="rp-fitem"><input type="checkbox" class="rp-cb" value="'
          + f.id + '" checked onchange="rpBuild()"><span>' + f.label + '</span></label>';
      }).join('')
    + '<span class="rp-fbtns"><button type="button" onclick="rpAll(true)">全選</button>'
    + '<button type="button" onclick="rpAll(false)">取消全選</button></span>';
}

function rpAll(checked) {
  document.querySelectorAll('.rp-cb').forEach(function (cb) { cb.checked = checked; });
  rpBuild();
}

function rpCheckedIds() {
  const out = [];
  RP_FIELDS.forEach(function (f) {
    const cb = document.querySelector('.rp-cb[value="' + f.id + '"]');
    if (!cb || cb.checked) out.push(f.id);
  });
  return out;
}

/* ---------------- 分組：同一 Model 的多種排程（plan） ---------------- */
function rpGroupByModel() {
  const order = [];
  const map = {};
  (LAST_FILTERED || []).forEach(function (r) {
    const m = r.Model || '';
    if (!map[m]) { map[m] = []; order.push(m); }
    map[m].push(r);
  });
  return order.map(function (m) { return { model: m, plans: map[m] }; });
}

/* 從 Current Status 的前綴猜方案特徵（例如 "RPL-HX R + DDR5"） */
function rpStatusHint(r) {
  const s = stripTags(r.CurrentStatus || '').replace(/^[\s\u3000]+/, '');
  const head = s.split(/[:：\n]/)[0].trim();
  return (head && head.length <= 48) ? head : '';
}

/* 方案預設名稱：優先 DDRx → 再來 Current Status 前綴 → 最後 #n */
function rpPlanDefaults(model, plans, prefix) {
  const base = prefix + model;
  const ddr = plans.map(function (r) {
    const m = rpStatusHint(r).match(/(?:LP)?DDR\d+X?/i);
    return m ? m[0].toUpperCase() : '';
  });
  if (ddr.every(function (v) { return v; }) && rpUniq(ddr).length === plans.length) {
    return plans.map(function (r, i) { return base + '(' + ddr[i] + ')'; });
  }
  const hints = plans.map(rpStatusHint);
  if (hints.every(function (v) { return v; }) && rpUniq(hints).length === plans.length) {
    return plans.map(function (r, i) { return base + ' (' + hints[i] + ')'; });
  }
  return plans.map(function (r, i) { return base + ' #' + (i + 1); });
}

function rpPlanKey(model, i) { return model + '#' + i; }

function rpEnsurePlanNames(groups, prefix) {
  const keys = [];
  groups.forEach(function (g) {
    if (g.plans.length > 1) {
      g.plans.forEach(function (r, i) { keys.push(rpPlanKey(g.model, i)); });
    }
  });
  const sig = keys.join('|') + '||' + prefix;
  if (sig !== RP_PLAN_SIG) {
    RP_PLAN_SIG = sig;
    keys.forEach(function (k) {
      delete RP_PLAN_NAMES[k];
      delete RP_PLAN_EDITED[k];
      delete RP_PLAN_HIDE[k];
    });
    Object.keys(RP_MERGE_SEL).forEach(function (k) { delete RP_MERGE_SEL[k]; });
  }
  // 依「目前可見」的方案重算預設名稱（使用者改過的不覆蓋）
  groups.forEach(function (g) {
    if (g.plans.length < 2) return;
    const vis = [];
    g.plans.forEach(function (r, i) {
      const k = rpPlanKey(g.model, i);
      if (!RP_PLAN_HIDE[k]) vis.push({ key: k, row: r });
    });
    if (!vis.length) return;
    const defaults = rpPlanDefaults(g.model, vis.map(function (v) { return v.row; }), prefix);
    vis.forEach(function (v, i) {
      if (!RP_PLAN_EDITED[v.key]) RP_PLAN_NAMES[v.key] = defaults[i];
    });
  });
}

function rpOnPlanName(el) {
  const k = el.getAttribute('data-pkey');
  RP_PLAN_NAMES[k] = el.textContent.trim();
  RP_PLAN_EDITED[k] = true;
  rpBuild();
}

function rpOnTitleEdit(el) {
  RP_TITLE_EDITS[el.getAttribute('data-tmodel')] = el.textContent.trim();
  rpBuild();
}

/* ---------------- 欄位名（可編輯） ---------------- */
function rpLabel(id) {
  if (RP_LABEL_EDITS[id] != null) return RP_LABEL_EDITS[id];
  const f = rpField(id);
  return f ? f.label : id;
}

function rpOnLabelEdit(el) {
  RP_LABEL_EDITS[el.getAttribute('data-fid')] = el.textContent.trim();
  rpBuild();
}

/* ---------------- 隱藏／還原方案 ---------------- */
function rpHidePlan(model, idx) {
  RP_PLAN_HIDE[rpPlanKey(model, idx)] = true;
  rpBuild();
}

function rpRestorePlans() {
  Object.keys(RP_PLAN_HIDE).forEach(function (k) { delete RP_PLAN_HIDE[k]; });
  rpBuild();
}

/* ---------------- 跨 Plan 合併格（僅限相鄰方案） ----------------
 * RP_MERGE_SEL['model#fieldId'] = 被勾選的方案「原始索引」陣列，
 * 勾選時自動補滿 min..max，因此一定是一段連續區間。 */
function rpMergeSelKey(model, fid) { return model + '#' + fid; }

function rpMergeRange(model, fid) {
  const a = (RP_MERGE_SEL[rpMergeSelKey(model, fid)] || []).slice()
    .sort(function (x, y) { return x - y; });
  if (a.length < 2) return null;
  return [a[0], a[a.length - 1]];
}

function rpTogglePlanMerge(model, fid, idx, on) {
  const key = rpMergeSelKey(model, fid);
  let a = (RP_MERGE_SEL[key] || []).slice();
  if (on) {
    if (a.indexOf(idx) < 0) a.push(idx);
  } else {
    a = a.filter(function (i) { return i !== idx; });
  }
  if (a.length >= 2) {                       // 以相鄰為限：補滿中間的方案
    const lo = Math.min.apply(null, a);
    const hi = Math.max.apply(null, a);
    const full = [];
    for (let i = lo; i <= hi; i++) full.push(i);
    a = full;
  }
  RP_MERGE_SEL[key] = a;
  rpBuild();
}

function rpClearMerges() {
  Object.keys(RP_MERGE_SEL).forEach(function (k) { delete RP_MERGE_SEL[k]; });
  rpBuild();
}

/* 取消某一欄的合併（合併格左上角的勾取消時） */
function rpClearMergeField(model, fid) {
  delete RP_MERGE_SEL[rpMergeSelKey(model, fid)];
  rpBuild();
}

/* ---------------- 組表格 ---------------- */
/* 格內左上角的合併勾選框。clearAll=true 代表這是「已合併」的格子，
   取消勾選即整欄拆回；否則勾選相鄰方案後自動合併。 */
function rpMergeBox(model, fid, idx, checked, clearAll) {
  const t = clearAll ? '取消勾選即拆回各方案' : '勾選相鄰方案即可合併';
  const act = clearAll
    ? 'rpClearMergeField(\'' + escapeHtml(model) + '\',\'' + fid + '\')'
    : 'rpTogglePlanMerge(\'' + escapeHtml(model) + '\',\'' + fid + '\',' + idx + ',this.checked)';
  return '<input type="checkbox" class="rp-mcb2" title="' + t + '"'
    + ' data-fid="' + fid + '" data-idx="' + idx + '"'
    + (checked ? ' checked' : '')
    + ' onclick="event.stopPropagation()" onchange="' + act + '">';
}

function rpValCell(r, id, prefix) {
  const inner = rpValue(r, id, prefix) || '';
  return '<div class="rp-edit" contenteditable="true" data-ph="（可直接輸入）">' + inner + '</div>';
}

function rpKCell(id) {
  return '<td class="rp-k"><div class="rp-ktxt" contenteditable="true" data-fid="' + id
    + '" onblur="rpOnLabelEdit(this)">' + escapeHtml(rpLabel(id)) + '</div></td>';
}

/* vis: [{row, idx}] —— idx 是方案在原始 plans 裡的索引 */
function rpRowHtml(g, vis, id, prefix) {
  let html = '<tr>' + rpKCell(id);
  const mergeable = (RP_MERGEABLE.indexOf(id) >= 0) && vis.length > 1;
  const range = mergeable ? rpMergeRange(g.model, id) : null;
  const inRange = range ? vis.filter(function (v) {
    return v.idx >= range[0] && v.idx <= range[1];
  }) : [];
  const sel = RP_MERGE_SEL[rpMergeSelKey(g.model, id)] || [];

  if (range && inRange.length >= 2) {
    const first = inRange[0].idx;
    vis.forEach(function (v) {
      if (v.idx < range[0] || v.idx > range[1]) {
        // 區間外的方案：維持獨立一格，左上角仍可勾選（可再併入）
        html += '<td class="rp-v rp-v-m">'
          + rpMergeBox(g.model, id, v.idx, false, false)
          + rpValCell(v.row, id, prefix) + '</td>';
      } else if (v.idx === first) {
        // 合併格：把各方案原本的預設內容以換行方式疊在一起，仍可手動編輯
        const merged = inRange.map(function (m) { return rpValue(m.row, id, prefix) || ''; }).join('');
        html += '<td class="rp-v rp-v-m" colspan="' + inRange.length + '">'
          + rpMergeBox(g.model, id, first, true, true)
          + '<div class="rp-edit" contenteditable="true" data-ph="（可直接輸入）">' + merged + '</div></td>';
      }
    });
    return html + '</tr>';
  }

  vis.forEach(function (v) {
    html += '<td class="rp-v' + (mergeable ? ' rp-v-m' : '') + '">'
      + (mergeable ? rpMergeBox(g.model, id, v.idx, sel.indexOf(v.idx) >= 0, false) : '')
      + rpValCell(v.row, id, prefix) + '</td>';
  });
  return html + '</tr>';
}

/* 單一方案：維持原本的 Email 格式 */
function rpBlockSingle(r, ids, prefix) {
  const modelName = prefix + (r.Model || '');
  let html = '<div class="rp-block">';
  html += '<div class="rp-greet">Dear All,</div>';
  html += '<div class="rp-subject">Update ' + escapeHtml(modelName) + ' weekly status below.</div>';
  html += '<table class="rp-table">';
  ids.forEach(function (id) {
    html += '<tr>' + rpKCell(id)
      + '<td class="rp-v">' + rpValCell(r, id, prefix) + '</td></tr>';
  });
  return html + '</table></div>';
}

/* 多方案：標題列 + 方案名稱列 + N 個方案欄。vis = [{row, idx}] */
function rpBlockMulti(g, ids, prefix, vis) {
  const n = vis.length;
  const modelName = prefix + g.model;
  const lw = 16;
  const pw = ((100 - lw) / n).toFixed(3);
  let html = '<div class="rp-block">';
  html += '<table class="rp-table" style="min-width:' + (170 + 165 * n) + 'px">';
  html += '<colgroup><col style="width:' + lw + '%">';
  for (let i = 0; i < n; i++) html += '<col style="width:' + pw + '%">';
  html += '</colgroup>';

  const title = (RP_TITLE_EDITS[g.model] != null)
    ? RP_TITLE_EDITS[g.model]
    : (modelName + ' Weekly Status 更新如下表，具體可參照附件 Control table, 有問題請聯繫，謝謝！');
  html += '<tr><td class="rp-title" colspan="' + (n + 1) + '" contenteditable="true" data-tmodel="'
    + escapeHtml(g.model) + '" onblur="rpOnTitleEdit(this)">' + escapeHtml(title) + '</td></tr>';

  html += '<tr><td class="rp-k"><div class="rp-ktxt" contenteditable="true" data-fid="Model" onblur="rpOnLabelEdit(this)">'
    + escapeHtml(rpLabel('Model')) + '</div></td>';
  vis.forEach(function (v) {
    const key = rpPlanKey(g.model, v.idx);
    html += '<td class="rp-plan"><div class="rp-ptxt" contenteditable="true" data-pkey="' + escapeHtml(key)
      + '" onblur="rpOnPlanName(this)">' + escapeHtml(RP_PLAN_NAMES[key] || '') + '</div>'
      + '<button type="button" class="rp-hide" onclick="rpHidePlan(\'' + escapeHtml(g.model) + '\',' + v.idx
      + ')" title="隱藏這個方案">✕</button></td>';
  });
  html += '</tr>';

  ids.forEach(function (id) {
    if (id === 'Model') return;
    html += rpRowHtml(g, vis, id, prefix);
  });
  return html + '</table></div>';
}

function rpBuild() {
  const page = document.getElementById('rpPage');
  const info = document.getElementById('rpInfo');
  if (!page) return;
  const pfEl = document.getElementById('rpPrefix');
  const prefix = pfEl ? pfEl.value : '';
  const ids = rpCheckedIds();
  const groups = rpGroupByModel();

  rpEnsurePlanNames(groups, prefix);

  if (!groups.length) {
    page.innerHTML = '<div class="rp-empty">目前沒有符合篩選的資料，請先調整篩選條件。</div>';
    page.classList.remove('rp-hasmerge');
    if (info) info.textContent = '0 個 Model';
    return;
  }

  let html = '';
  let multi = 0;
  let hidden = 0;
  groups.forEach(function (g) {
    const vis = [];
    g.plans.forEach(function (r, i) {
      if (!RP_PLAN_HIDE[rpPlanKey(g.model, i)]) vis.push({ row: r, idx: i });
    });
    hidden += g.plans.length - vis.length;
    if (vis.length > 1) { multi++; html += rpBlockMulti(g, ids, prefix, vis); }
    else if (vis.length === 1) { html += rpBlockSingle(vis[0].row, ids, prefix); }
  });
  page.innerHTML = html;
  page.classList.toggle('rp-hasmerge', multi > 0);

  const hb = document.getElementById('rpHidden');
  if (hb) {
    hb.innerHTML = hidden > 0
      ? '　已隱藏 ' + hidden + ' 個方案 <button type="button" onclick="rpRestorePlans()">還原全部</button>'
      : '';
  }

  if (info) {
    info.textContent = groups.length + ' 個 Model（' + multi + ' 個多方案）／' + ids.length + ' 個欄位';
  }
}

function rpDefaultFileName() {
  const first = (LAST_FILTERED[0] || {}).Model || 'report';
  const pf = document.getElementById('rpPrefix');
  const tag = new Date().toISOString().slice(0, 10).replace(/-/g, '');
  return '週報_' + ((pf ? pf.value : '') + first) + '_' + tag;
}

function openReport() {
  if (!(LAST_FILTERED || []).length) {
    alert('目前沒有符合篩選的資料，請先調整篩選條件後再產生週報。');
    return;
  }
  const fn = document.getElementById('rpFileName');
  if (fn) fn.value = rpDefaultFileName();
  document.getElementById('rpBk').classList.add('show');
  rpBuild();
  document.getElementById('rpBk').scrollTop = 0;
}

function rpClose() {
  document.getElementById('rpBk').classList.remove('show');
}

/* 拍攝前暫時解除高度/捲動限制，避免報表被裁掉 */
async function rpPrepareCapture() {
  const bk = document.getElementById('rpBk');
  const shell = bk.querySelector('.rp-shell');
  const body = document.getElementById('rpBody');
  const page = document.getElementById('rpPage');
  const saved = {
    bk: bk.getAttribute('style') || '',
    shell: shell.getAttribute('style') || '',
    body: body.getAttribute('style') || ''
  };
  bk.style.position = 'static';
  bk.style.overflow = 'visible';
  bk.style.padding = '0';
  bk.style.background = 'transparent';
  shell.style.boxShadow = 'none';
  body.style.overflow = 'visible';
  body.style.maxHeight = 'none';
  page.classList.add('rp-capturing');
  await new Promise(function (res) { setTimeout(res, 40); });
  return { bk: bk, shell: shell, body: body, page: page, saved: saved };
}

function rpRestoreCapture(ctx) {
  ctx.bk.setAttribute('style', ctx.saved.bk);
  ctx.shell.setAttribute('style', ctx.saved.shell);
  ctx.body.setAttribute('style', ctx.saved.body);
  ctx.page.classList.remove('rp-capturing');
}

/* 對報表本體做長截圖（PNG） */
async function rpShot(btn) {
  if (typeof html2canvas === 'undefined') { alert('截圖函式庫未載入。'); return; }
  const oldText = btn ? btn.textContent : '';
  if (btn) { btn.disabled = true; btn.textContent = '截圖中…'; }
  const ctx = await rpPrepareCapture();
  try {
    const page = ctx.page;
    const canvas = await html2canvas(page, {
      scale: 2,
      backgroundColor: '#ffffff',
      useCORS: true,
      logging: false,
      width: page.scrollWidth,
      height: page.scrollHeight,
      windowWidth: document.documentElement.clientWidth
    });
    const fnEl = document.getElementById('rpFileName');
    const name = (fnEl && fnEl.value ? fnEl.value : rpDefaultFileName()).trim();
    const a = document.createElement('a');
    a.download = name + '.png';
    a.href = canvas.toDataURL('image/png');
    document.body.appendChild(a);
    a.click();
    a.remove();
  } catch (e) {
    alert('長截圖失敗：' + e);
  } finally {
    rpRestoreCapture(ctx);
    if (btn) { btn.disabled = false; btn.textContent = oldText; }
  }
}

/* 確定後另存新檔：輸出可獨立開啟的單檔 HTML */
/* ===================== 週報 → Excel（格式、顏色同 HTML） =====================
 * 用 SheetJS 的 .s 樣式物件；這個 build 的 xlsx.js 會實際寫出 xl/styles.xml，
 * 所以填色／框線／粗體都會保留。
 *   單一方案 → Dear All, / Update XXX weekly status below. + 兩欄表
 *   多方案   → 標題列(#fff2cc) + 方案名稱列(#cfe2f3) + N 個方案欄；
 *              跨 Plan 合併格用真正的 !merges，內容就是畫面上看到的那一格。
 * ========================================================================== */

/* DOM → 純文字：<div>/<p>/<li> 與 <br> 都當換行，保留使用者手動編輯的斷行 */
function rpPlain(el) {
  if (!el) return '';
  const lines = [''];
  const cur = function () { return lines[lines.length - 1]; };
  const setLine = function (v) { lines[lines.length - 1] = v; };
  const nl = function () { if (cur() !== '') lines.push(''); };
  (function walk(node) {
    Array.prototype.forEach.call(node.childNodes, function (n) {
      if (n.nodeType === 3) { setLine(cur() + n.nodeValue); return; }
      if (n.nodeType !== 1) return;
      const tag = String(n.tagName).toLowerCase();
      if (tag === 'br') { nl(); return; }
      const block = (tag === 'div' || tag === 'p' || tag === 'li');
      if (block) nl();
      walk(n);
      if (block) nl();
    });
  })(el);
  const out = lines.map(function (s) { return s.replace(/[ \t]+$/, ''); });
  while (out.length && out[0] === '') out.shift();
  while (out.length && out[out.length - 1] === '') out.pop();
  return out.join('\n');
}

/* 單行文字（欄位名／方案名／標題） */
function rpText(el) {
  return el ? String(el.textContent || '').replace(/\s+/g, ' ').trim() : '';
}

/* 估算文字在給定欄寬下佔幾行（中日韓全形字算 2 個單位） */
function rpUnits(s) {
  let u = 0;
  for (let i = 0; i < s.length; i++) u += (s.charCodeAt(i) > 0x2E80 ? 2 : 1);
  return u;
}

function rpEstLines(text, wch) {
  const w = Math.max(4, wch - 1);
  return String(text == null ? '' : text).split('\n')
    .reduce(function (n, p) { return n + Math.max(1, Math.ceil(rpUnits(p) / w)); }, 0);
}

/* 把畫面上的 #rpPage 轉成一張 worksheet */
function rpSheet() {
  const page = document.getElementById('rpPage');
  const BORD = {
    top: { style: 'thin', color: { rgb: 'FF444444' } },
    bottom: { style: 'thin', color: { rgb: 'FF444444' } },
    left: { style: 'thin', color: { rgb: 'FF444444' } },
    right: { style: 'thin', color: { rgb: 'FF444444' } }
  };
  const AL_L = { horizontal: 'left', vertical: 'top', wrapText: true };
  const AL_C = { horizontal: 'center', vertical: 'center', wrapText: true };
  const S_TITLE = { fill: { patternType: 'solid', fgColor: { rgb: 'FFFFF2CC' } }, font: { sz: 10 }, alignment: AL_L, border: BORD };
  const S_PLAN  = { fill: { patternType: 'solid', fgColor: { rgb: 'FFCFE2F3' } }, font: { sz: 10, bold: true }, alignment: AL_C, border: BORD };
  const S_KEY   = { font: { sz: 10, bold: true }, alignment: AL_L, border: BORD };
  const S_VAL   = { font: { sz: 10 }, alignment: AL_L, border: BORD };
  const S_PLAIN = { font: { sz: 10 } };

  const W_KEY = 20, W_VAL = 30, LINE_H = 13.2;
  const grid = [], merges = [], rowH = [];
  let maxCol = 1;

  function set(r, c, v, s) {
    if (!grid[r]) grid[r] = [];
    grid[r][c] = { t: 's', v: v == null ? '' : String(v), s: s };
    if (c + 1 > maxCol) maxCol = c + 1;
  }
  function rowLines(r, n) { rowH[r] = Math.max(rowH[r] || 0, n); }

  let R = 0;
  const blocks = Array.prototype.slice.call(page.querySelectorAll('.rp-block'));

  blocks.forEach(function (blk, bi) {
    if (bi > 0) R += 1;                                   // 區塊之間留一列空白
    const greet = blk.querySelector('.rp-greet');
    const subject = blk.querySelector('.rp-subject');
    if (greet) { set(R, 0, rpText(greet), S_PLAIN); rowLines(R, 1); R += 1; }
    if (subject) { set(R, 0, rpText(subject), S_PLAIN); rowLines(R, 1); R += 1; }

    const trs = Array.prototype.slice.call(blk.querySelectorAll('table.rp-table tr'));
    trs.forEach(function (tr) {
      const tds = Array.prototype.filter.call(tr.children, function (e) {
        return String(e.tagName).toLowerCase() === 'td';
      });
      if (!tds.length) return;
      let c = 0, lines = 1;
      tds.forEach(function (td) {
        const cls = ' ' + (td.className || '') + ' ';
        const span = Math.max(1, parseInt(td.getAttribute('colspan') || '1', 10) || 1);
        let txt = '', style = S_VAL;
        if (cls.indexOf(' rp-title ') >= 0) {
          txt = rpText(td); style = S_TITLE;
        } else if (cls.indexOf(' rp-plan ') >= 0) {
          txt = rpText(td.querySelector('.rp-ptxt')); style = S_PLAN;
        } else if (cls.indexOf(' rp-k ') >= 0) {
          txt = rpText(td.querySelector('.rp-ktxt')); style = S_KEY;
        } else {
          txt = rpPlain(td.querySelector('.rp-edit')); style = S_VAL;
        }
        set(R, c, txt, style);
        if (span > 1) merges.push({ s: { r: R, c: c }, e: { r: R, c: c + span - 1 } });
        lines = Math.max(lines, rpEstLines(txt, span * (c === 0 ? W_KEY : W_VAL)));
        c += span;
      });
      if (c > maxCol) maxCol = c;
      rowLines(R, lines);
      R += 1;
    });
  });

  const lastRow = Math.max(1, R);
  const ws = {};
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: lastRow - 1, c: maxCol - 1 } });
  grid.forEach(function (cells, r) {
    (cells || []).forEach(function (cell, c) {
      if (cell) ws[XLSX.utils.encode_cell({ r: r, c: c })] = cell;
    });
  });
  if (merges.length) ws['!merges'] = merges;
  ws['!cols'] = [];
  for (let c = 0; c < maxCol; c++) ws['!cols'].push({ wch: c === 0 ? W_KEY : W_VAL });
  ws['!rows'] = [];
  for (let r = 0; r < lastRow; r++) {
    ws['!rows'].push({ hpt: Math.max(15, Math.round((rowH[r] || 1) * LINE_H + 3)) });
  }
  return ws;
}

/* 確定並另存新檔 → Excel（.xlsx） */
function rpSaveExcel() {
  if (typeof XLSX === 'undefined') { alert('Excel 匯出函式庫未載入。'); return; }
  const page = document.getElementById('rpPage');
  if (!page || !page.querySelector('.rp-block')) { alert('目前沒有可匯出的週報內容。'); return; }
  const fnEl = document.getElementById('rpFileName');
  const raw = ((fnEl && fnEl.value ? fnEl.value : rpDefaultFileName()) || 'weekly_report').trim();
  const name = raw.replace(/\.(html?|xlsx?)$/i, '').replace(/[\\/:*?"<>|]/g, '_') || 'weekly_report';
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, rpSheet(), name.replace(/[\\/?*\[\]:]/g, '_').slice(0, 31) || 'Weekly');
  XLSX.writeFile(wb, name + '.xlsx');
}

function rpSave() {
  const page = document.getElementById('rpPage');
  const clone = page.cloneNode(true);
  clone.removeAttribute('id');
  clone.classList.remove('rp-hasmerge');
  clone.classList.remove('rp-capturing');
  clone.querySelectorAll('[contenteditable]').forEach(function (e) { e.removeAttribute('contenteditable'); });

  const cssEl = document.getElementById('rpCss');
  const css = cssEl ? cssEl.textContent : '';
  const fnEl = document.getElementById('rpFileName');
  const name = ((fnEl && fnEl.value ? fnEl.value : rpDefaultFileName()) || 'weekly_report')
    .trim().replace(/[\\/:*?"<>|]/g, '_');
  const doc = '<!DOCTYPE html>\n<html lang="zh-TW">\n<head>\n<meta charset="UTF-8">\n'
    + '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
    + '<title>' + escapeHtml(name) + '</title>\n<style>\n' + css + '\n'
    + 'body { margin: 0; background: #fff; font-family: "Microsoft JhengHei", "Segoe UI", sans-serif; }\n'
    + '</style>\n</head>\n<body>\n<div id="rpPage">' + clone.innerHTML + '</div>\n</body>\n</html>\n';

  const blob = new Blob([doc], { type: 'text/html;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.download = name + '.html';
  a.href = url;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(function () { URL.revokeObjectURL(url); }, 5000);
}
'''


def main():
    missing = [p for p in (LIB_JSZIP, LIB_XLSX, LIB_H2C, LIB_CROPJS, LIB_CROPCSS, LIB_CROPHTML)
               if not os.path.exists(p)]
    if missing:
        raise FileNotFoundError("缺少 lib 函式庫檔案：\n  " + "\n  ".join(missing))

    wb = openpyxl.load_workbook(SPEC_PATH, data_only=True)
    ws = wb["Schedule"]
    header = {}
    for c in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=c).value
        if val:
            header[str(val).strip()] = c

    spec_rows = []
    filter_values = defaultdict(set)

    for r in range(2, ws.max_row + 1):
        model_name = ws.cell(row=r, column=header.get("Model Name", 1)).value
        if model_name is None or str(model_name).strip() == "":
            continue

        stage = ws.cell(row=r, column=header.get("Stage", 9)).value
        mkt_name = ws.cell(row=r, column=header.get("MKT Name", 2)).value
        cpu = ws.cell(row=r, column=header.get("CPU", 5)).value
        gpu = ws.cell(row=r, column=header.get("GPU", 6)).value
        npm = ws.cell(row=r, column=header.get("NPM", 7)).value
        current_status = ws.cell(row=r, column=header.get("Current Status", 18)).value
        highlight = ws.cell(row=r, column=header.get("Highlight", 19)).value

        dates = {}
        for key, col_letter in SPEC_MILESTONE_COLS.items():
            col = header.get(key)
            if col is None:
                col = openpyxl.utils.column_index_from_string(col_letter)
            dates[key] = normalize_date(ws.cell(row=r, column=col).value)

        row = {
            "Model": str(model_name).strip() if model_name else "",
            "MKT": str(mkt_name).strip() if mkt_name else "",
            "CPU": str(cpu).strip() if cpu else "",
            "GPU": str(gpu).strip() if gpu else "",
            "NPM": str(npm).strip() if npm else "",
            "Stage": str(stage).strip() if stage else "",
            "CurrentStatus": str(current_status).strip() if current_status else "",
            "Highlight": str(highlight).strip() if highlight else "",
            "MPYear": extract_year(dates.get("MP")),
            **{k: format_date_yyyy_mm_dd(v) for k, v in dates.items()},
            "_dates": dates,
        }
        spec_rows.append(row)
        filter_values["Model"].add(row["Model"])
        filter_values["MKT"].add(row["MKT"])
        filter_values["CPU"].add(row["CPU"])
        filter_values["GPU"].add(row["GPU"])
        filter_values["Stage"].add(row["Stage"])
        filter_values["MPYear"].add(row["MPYear"])

    wb.close()

    filters = {k: sorted(v, key=lambda x: (x == "", x)) for k, v in filter_values.items()}
    merged_rows = build_merged_rows(spec_rows)

    rows_json = json.dumps(merged_rows, ensure_ascii=False)
    filters_json = json.dumps(filters, ensure_ascii=False)

    filter_keys = [
        ("Model", "Model"),
        ("MKT", "MKT"),
        ("CPU", "CPU"),
        ("GPU", "GPU"),
        ("Stage", "Stage"),
        ("MPYear", "MP 年份"),
    ]
    filter_html = "\n".join(
        f'''  <div class="filter-group">
    <label>{label}</label>
    <div class="mf-dropdown" id="dd-{key}">
      <button type="button" class="mf-toggle" id="tgl-{key}" onclick="toggleMenu('{key}');event.stopPropagation();">全部 {label} ▼</button>
      <div class="mf-menu" id="menu-{key}">
        <input type="text" class="mf-search" id="search-{key}" placeholder="模糊搜尋 {label}…" oninput="filterItems('{key}')">
        <div class="mf-actions">
          <button type="button" onclick="selectAll('{key}', true)">全選</button>
          <button type="button" onclick="selectAll('{key}', false)">取消全選</button>
          <button type="button" onclick="selectMatched('{key}')">僅選匹配</button>
        </div>
        <div class="mf-list" id="list-{key}"></div>
        <div class="mf-match" id="match-{key}"></div>
      </div>
    </div>
  </div>'''
        for key, label in filter_keys
    )

    range_html = '''
  <div class="filter-group" style="min-width:230px;flex:0 0 230px;">
    <label>MP 時間範圍 (yyyy/mm)</label>
    <div style="display:flex;align-items:center;gap:6px;padding-top:2px;">
      <input type="text" id="mp-start" class="mp-range" placeholder="2026/01" oninput="applyFilters()">
      <span class="mp-sep">~</span>
      <input type="text" id="mp-end" class="mp-range" placeholder="2026/12" oninput="applyFilters()">
    </div>
  </div>'''

    filter_html = filter_html + range_html

    html = f'''<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="Expires" content="0">
<script>
// 强制绕过缓存：每次访问都请求最新版本
// （2026-10-03 使用者要求，與 npi_dashboard.html / npi_dashboard2.html 一致）
// 注意：這是 Python f-string，所有 JS 的 {{ }} 都要寫成 {{{{ }}}}
(function(){{
  var ts = Date.now();
  var href = location.href.split('#')[0];
  if(href.indexOf('_nocache_') === -1){{
    var sep = href.indexOf('?') === -1 ? '?' : '&';
    location.replace(href + sep + '_nocache_=' + ts);
  }}
}})();
</script>
<title>TTM Dashboard (NPI_vx)</title>
<style>
:root {{ color-scheme: light; --green:#1f8a4c; --green-d:#15693a; --green-l:#e8f5ec; --line:#d8e3dc; --ink:#1f2a24; }}
body {{ font-family: "Microsoft JhengHei", "Segoe UI", sans-serif; margin: 16px; background: #f6faf7; color: var(--ink); }}
h1 {{ margin: 0 0 12px; font-size: 22px; color: var(--green-d); }}
.filter-bar {{ display: flex; flex-wrap: wrap; gap: 12px; align-items: flex-end; margin-bottom: 12px; background: #fff; padding: 12px; border-radius: 8px; box-shadow: 0 1px 3px rgba(31,138,76,.05); border: 1px solid var(--line); }}
.filter-group {{ display: flex; flex-direction: column; min-width: 200px; flex: 1 1 200px; }}
.filter-group label {{ font-size: 12px; color: #555; margin-bottom: 4px; font-weight: 600; }}
button {{ padding: 6px 14px; border: 1px solid var(--line); border-radius: 4px; background: #fff; cursor: pointer; font-size: 13px; }}
button:hover {{ background: var(--green-l); }}
button.primary {{ background: var(--green); color: #fff; border-color: var(--green); }}
button.primary:hover {{ background: var(--green-d); }}
.stats {{ margin-bottom: 12px; font-size: 14px; color: #444; }}
table {{ width: 100%; border-collapse: collapse; table-layout: fixed; background: #fff; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(31,138,76,.04); font-size: 13px; border: 1px solid var(--line); }}
th, td {{ border: 1px solid var(--line); padding: 8px; text-align: left; vertical-align: top; overflow-wrap: anywhere; }}
th {{ background: var(--green); color: #fff; position: sticky; top: 0; z-index: 2; }}
tr:nth-child(even) {{ background: #f9fafb; }}
tr:hover {{ background: #eef4ff; }}
.highlight-cell {{ white-space: pre-wrap; word-break: break-word; color: #b42318; }}
.status-cell {{ text-align: left; white-space: normal; word-break: break-word; font-weight: 600; }}
.merged-item {{ display: flex; align-items: flex-start; gap: 6px; margin-bottom: 5px; line-height: 1.4; }}
.merged-item:last-child {{ margin-bottom: 0; }}
.seq-badge {{ display: inline-flex; align-items: center; justify-content: center; width: 20px; height: 20px; border-radius: 50%; background: #1f8a4c; color: #fff; font-size: 11px; font-weight: 700; flex-shrink: 0; margin-top: 1px; }}
.empty {{ color: #999; font-style: italic; }}
th.sortable {{ cursor: pointer; user-select: none; white-space: nowrap; }}
th.sortable:hover {{ background: var(--green-d); }}
.sort-ind {{ display: inline-block; font-size: 10px; margin-left: 2px; opacity: .85; }}
.mp-range {{ width: 92px; padding: 6px 8px; border: 1px solid var(--line); border-radius: 6px; font-size: 13px; text-align: center; }}
.mp-range:focus {{ outline: 2px solid var(--green); }}
.mp-sep {{ color: #777; }}

/* 多選下拉元件 */
.mf-dropdown {{ position: relative; display: inline-block; width: 100%; }}
.mf-toggle {{ width: 100%; background: #fff; border: 1px solid var(--line); border-radius: 6px; padding: 6px 12px; font-size: 13px; color: var(--ink); cursor: pointer; text-align: left; }}
.mf-toggle:hover {{ background: var(--green-l); }}
.mf-menu {{ display: none; position: absolute; top: 100%; left: 0; margin-top: 4px; min-width: 100%; max-height: 320px; overflow-y: auto; background: #fff; border: 1px solid var(--line); border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,.12); z-index: 20; padding: 6px 0; }}
.mf-menu.open {{ display: block; }}
.mf-actions {{ display: flex; gap: 6px; padding: 6px 10px 8px; border-bottom: 1px solid var(--line); }}
.mf-actions button {{ background: #fff; border: 1px solid var(--line); border-radius: 5px; padding: 3px 8px; font-size: 11px; color: var(--green-d); cursor: pointer; }}
.mf-actions button:hover {{ background: var(--green-l); }}
.mf-item {{ display: flex; align-items: center; gap: 6px; padding: 5px 10px; font-size: 12px; cursor: pointer; white-space: nowrap; }}
.mf-item:hover {{ background: #f6faf7; }}
.mf-search {{ width: calc(100% - 20px); margin: 8px 10px 4px; padding: 5px 8px; border: 1px solid var(--line); border-radius: 6px; font-size: 12px; }}
.mf-search:focus {{ outline: none; border-color: var(--green); }}
.mf-match {{ padding: 4px 10px 6px; font-size: 11px; color: #5a6b61; }}

/* 排序專區（獨立於多選下拉） — 4 顆欄位按鈕 */
.sort-bar {{ display: flex; align-items: center; gap: 10px; flex-wrap: wrap; margin-bottom: 12px; padding: 10px 14px; background: #fff; border: 1px solid var(--line); border-left: 4px solid var(--green); border-radius: 8px; box-shadow: 0 1px 3px rgba(31,138,76,.05); font-size: 13px; }}
.sort-label {{ font-weight: 700; color: var(--green-d); }}
.sort-btn {{ display: inline-flex; align-items: center; gap: 5px; padding: 6px 14px; border: 1px solid var(--line); border-radius: 6px; background: #fff; color: var(--ink); font-size: 13px; cursor: pointer; transition: background .15s, border-color .15s, color .15s; }}
.sort-btn:hover {{ background: var(--green-l); border-color: var(--green); }}
/* 使用中：綠底白字（與原本的 #sort-dir 視覺一致） */
.sort-btn.active {{ background: var(--green); border-color: var(--green); color: #fff; font-weight: 700; }}
.sort-btn.active:hover {{ background: var(--green-d); border-color: var(--green-d); }}
/* 第 2 層（含）以後的排序鍵：用較淺的綠底＋深綠字，與主鍵區分 */
.sort-btn.sort-lv2 {{ background: #dff0e5; border-color: var(--green); color: var(--green-d); }}
.sort-btn.sort-lv2:hover {{ background: #cfe8d8; }}
.sort-btn .sort-ind {{ font-size: 11px; opacity: .75; }}
.sort-btn.active .sort-ind, .sort-btn.sort-lv2 .sort-ind {{ opacity: 1; font-size: 12px; font-weight: 700; }}
.sort-clear {{ padding: 6px 12px; border: 1px solid var(--line); border-radius: 6px; background: #fff; color: #5a6b61; font-size: 12px; cursor: pointer; }}
.sort-clear:hover {{ background: #f2f4f3; color: var(--ink); }}
.sort-hint {{ color: #888; font-size: 12px; }}

/* 長截圖 / 範圍截圖樣式 */
/*__CROPCSS__*/
</style>
<style id="rpCss">/*__REPORTCSS__*/</style>
</head>
<body>
<h1>TTM Dashboard (NPI_vx)</h1>
<div class="filter-bar">
{filter_html}
  <div class="filter-group" style="flex:0 0 auto;min-width:auto;">
    <label>&nbsp;</label>
    <div>
      <button class="primary" onclick="applyFilters()">套用篩選</button>
      <button onclick="clearFilters()">清除</button>
      <button id="btn-mid" onclick="toggleMid()">顯示中間里程碑</button>
      <button onclick="downloadExcel()">💾 下載 Excel</button>
      <button class="btn-report" onclick="openReport()">📧 週報格式</button>
      <button class="btn-shot" onclick="takeScreenshot()">📷 長截圖</button>
      <button class="btn-crop" onclick="takeCropScreenshot()">✂️ 範圍截圖</button>
    </div>
  </div>
</div>
<div class="sort-bar">
  <span class="sort-label">排序：</span>
  <button class="sort-btn" data-key="Model" onclick="toggleSort('Model')">Model<span class="sort-ind" data-key="Model">⇅</span></button>
  <button class="sort-btn" data-key="CPU" onclick="toggleSort('CPU')">CPU<span class="sort-ind" data-key="CPU">⇅</span></button>
  <button class="sort-btn" data-key="Kickoff" onclick="toggleSort('Kickoff')">Kickoff<span class="sort-ind" data-key="Kickoff">⇅</span></button>
  <button class="sort-btn" data-key="MP" onclick="toggleSort('MP')">MP<span class="sort-ind" data-key="MP">⇅</span></button>
  <button id="sort-clear" class="sort-clear" onclick="clearSort()">重設排序</button>
  <span class="sort-hint">依點選順序多層排序：第 1 顆為主鍵、第 2 顆為次鍵…；同顆再點切換升／降序，再點一次取消</span>
</div>
<div class="stats" id="stats"></div>
<table id="data-table">
  <thead>
    <tr>
      <th style="width:3.5%" class="sortable" title="點擊排序" onclick="toggleSort('Model')">Model name<span class="sort-ind" data-key="Model">⇅</span></th>
      <th style="width:9%">MKT name</th>
      <th style="width:6%" class="sortable" title="點擊排序" onclick="toggleSort('CPU')">CPU<span class="sort-ind" data-key="CPU">⇅</span></th>
      <th style="width:6%">GPU</th>
      <th style="width:2%">NPM</th>
      <th style="width:2.5%">Stage</th>
      <th style="width:9%" class="sortable" title="點擊排序" onclick="toggleSort('Kickoff')">Kickoff<span class="sort-ind" data-key="Kickoff">⇅</span></th>
      <th class="mid-milestone" style="width:2%">DVT-start</th>
      <th class="mid-milestone" style="width:2%">EVT-start</th>
      <th class="mid-milestone" style="width:2%">MVT-start</th>
      <th class="mid-milestone" style="width:2%">ATS-start</th>
      <th style="width:9%" class="sortable" title="點擊排序" onclick="toggleSort('MP')">MP<span class="sort-ind" data-key="MP">⇅</span></th>
      <th style="width:22.5%">Current Status</th>
      <th style="width:22.5%">Highlight</th>
    </tr>
  </thead>
  <tbody id="tbody"></tbody>
</table>

<!--__CROPMODAL__-->

<!--__REPORTMODAL__-->
<script>/*__LIB__*/</script>
<script>/*__SHOTLIB__*/</script>
<script>
const RAW_ROWS = {rows_json};
const FILTERS = {filters_json};
const TOTAL_ORIGINAL = {len(spec_rows)};
const FILTER_META = {{ "Model": "Model", "MKT": "MKT", "CPU": "CPU", "GPU": "GPU", "Stage": "Stage", "MPYear": "MP 年份" }};
const DEFAULT_STAGES = {json.dumps(DEFAULT_STAGES, ensure_ascii=False)};
const DEFAULT_MPYEARS = {json.dumps(DEFAULT_MPYEARS, ensure_ascii=False)};
let midVisible = false;
// ── 多層排序（依點選順序）──
// sortSpecs = [{{key:'MP', dir:'asc'}}, {{key:'CPU', dir:'desc'}}, ...]
// 陣列順序 = 優先順序：第 1 個是主要排序鍵，第 2 個是平手時的次要鍵，依此類推。
// 同一欄不會重複出現（要改就調方向）。
let sortSpecs = [];
const SORT_KEYS = ['Model', 'CPU', 'Kickoff', 'MP'];
let LAST_FILTERED = [];

function init() {{
  for (const key of Object.keys(FILTER_META)) {{
    renderFilter(key);
  }}
  rpRenderFields();
  applyFilters();
  applyMidVisibility();
}}

function toggleMid() {{
  midVisible = !midVisible;
  applyMidVisibility();
  document.getElementById('btn-mid').textContent = midVisible ? '隱藏中間里程碑' : '顯示中間里程碑';
}}

function applyMidVisibility() {{
  document.querySelectorAll('.mid-milestone').forEach(e => e.style.display = midVisible ? '' : 'none');
}}

/* 選項一律來自 FILTERS（= Spec總表.xlsx 的實際值）；
   DEFAULT_STAGES / DEFAULT_MPYEARS 只決定預設勾選，不會新增不存在的選項。
   比對 Stage 用**字串完全比對**：清單寫 'ATS' 只會勾到資料裡的 'ATS'，
   資料裡是 'ATS-' 就**不會**被預設勾到（要連 ATS- 一起預設勾就把它寫進 DEFAULT_STAGES）。 */
function isDefaultChecked(key, value) {{
  if (key === 'Stage') return DEFAULT_STAGES.includes(value);
  if (key === 'MPYear') return DEFAULT_MPYEARS.includes(value);
  return true;
}}

function renderFilter(key) {{
  const opts = FILTERS[key] || [];
  const list = document.getElementById('list-' + key);
  if (!list) return;
  list.innerHTML = opts.map((v) => {{
    const text = v || '(空白)';
    const def = isDefaultChecked(key, v) ? 'checked' : '';
    return '<label class="mf-item" data-text="' + escapeHtml(text) + '">' +
      '<input type="checkbox" class="cb-' + key + '" value="' + escapeHtml(v) + '" ' + def + ' onchange="updateToggleLabels(); applyFilters();">' +
      '<span>' + escapeHtml(text) + '</span></label>';
  }}).join('');
  filterItems(key);
  updateToggleLabels();
}}

function escapeHtml(s) {{
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}}

function toggleMenu(key) {{
  const menu = document.getElementById('menu-' + key);
  const wasOpen = menu.classList.contains('open');
  document.querySelectorAll('.mf-menu').forEach(m => m.classList.remove('open'));
  if (!wasOpen) menu.classList.add('open');
}}

document.addEventListener('click', e => {{
  if (!e.target.closest('.mf-dropdown')) {{
    document.querySelectorAll('.mf-menu').forEach(m => m.classList.remove('open'));
  }}
}});

function filterItems(key) {{
  const q = (document.getElementById('search-' + key).value || '').toLowerCase();
  let matched = 0, total = 0;
  document.querySelectorAll('.cb-' + key).forEach(cb => {{
    total++;
    const label = cb.closest('.mf-item');
    const text = label.textContent.toLowerCase();
    const ok = !q || text.includes(q);
    label.style.display = ok ? '' : 'none';
    if (ok) matched++;
  }});
  document.getElementById('match-' + key).textContent = '匹配 ' + matched + ' / ' + total + ' 個';
}}

function selectAll(key, checked) {{
  document.querySelectorAll('.cb-' + key).forEach(cb => cb.checked = checked);
  updateToggleLabels();
  applyFilters();
}}

function selectMatched(key) {{
  const q = (document.getElementById('search-' + key).value || '').toLowerCase();
  document.querySelectorAll('.cb-' + key).forEach(cb => {{
    const label = cb.closest('.mf-item');
    const text = label.textContent.toLowerCase();
    cb.checked = !q || text.includes(q);
  }});
  updateToggleLabels();
  applyFilters();
}}

function getSelected(key) {{
  return Array.from(document.querySelectorAll('.cb-' + key + ':checked')).map(cb => cb.value);
}}

function updateToggleLabels() {{
  for (const [key, label] of Object.entries(FILTER_META)) {{
    const total = document.querySelectorAll('.cb-' + key).length;
    const sel = document.querySelectorAll('.cb-' + key + ':checked').length;
    const tgl = document.getElementById('tgl-' + key);
    tgl.textContent = sel === total ? '全部 ' + label + ' ▼' : '已選 ' + sel + ' / ' + total + ' ' + label + ' ▼';
  }}
}}

function normYm(s) {{
  if (!s) return null;
  const m = ("" + s).trim().match(/^(\\d{{4}})[/-](\\d{{1,2}})$/);
  if (!m) return null;
  const y = parseInt(m[1], 10);
  const mo = parseInt(m[2], 10);
  if (mo < 1 || mo > 12) return null;
  return y + "/" + String(mo).padStart(2, "0");
}}

function mpMonth(s) {{
  if (!s) return null;
  const m = ("" + s).trim().match(/^(\\d{{4}})[/-](\\d{{1,2}})/);
  if (!m) return null;
  const mo = parseInt(m[2], 10);
  if (mo < 1 || mo > 12) return null;
  return m[1] + "/" + String(mo).padStart(2, "0");
}}

function sortVal(r, key) {{
  if (key === 'Model') return (r['Model'] || '').trim();
  if (key === 'CPU') return cpuSortKey(r);
  if (key === 'Kickoff') return (r['Kickoff'] || '').trim();
  if (key === 'MP') return (r['MP'] || '').trim();
  return '';
}}

// 從 r['_CPU']（純文字陣列，去重後）取前 3 碼。
// ⚠️ 絕對不可以改回讀 r['CPU'] 字串！
//    r['CPU'] 是「畫面用 HTML」，合併列會是一串
//    <div class="merged-item"><span class="seq-badge">1</span>NVL-HX 55W</div>…
//    序號徽章的數字是純文字節點、stripTags 剝不掉，去標籤後會變成
//    "1NVL-HX 55W2NVL-HX 55W…" → 前 3 碼變成 "1NV" → 該列被排到最前面（2026-10-03 使用者回報的 bug）。
//    r['_CPU'] 是 build_merged_rows() 產的 clean 陣列，每個元素都已 trim 過。
function cpuSortKey(r) {{
  const arr = r['_CPU'];
  let s = '';
  if (Array.isArray(arr) && arr.length) {{
    for (const v of arr) {{
      // 逐一剝標籤再 trim；空字串就往下一筆找
      const t = ('' + (v == null ? '' : v)).replace(/<[^>]+>/g, '').trim();
      if (t) {{ s = t; break; }}
    }}
  }} else {{
    // 後備：沒有 _CPU 時才退回字串欄位（此時一定不是合併列）
    s = ('' + (r['CPU'] == null ? '' : r['CPU'])).replace(/<[^>]+>/g, '').trim();
  }}
  return s.slice(0, 3).toUpperCase();
}}

function toggleSort(key) {{
  if (!SORT_KEYS.includes(key)) return;
  // 三態循環（依點選順序多層排序）：
  //   尚未加入      → 加到「最後一層」，升序
  //   已在清單、升序 → 原地切成降序（順位不變）
  //   已在清單、降序 → 從清單移除（後面的層次自動遞補）
  const idx = sortSpecs.findIndex(s => s.key === key);
  if (idx < 0) {{
    sortSpecs.push({{ key: key, dir: 'asc' }});
  }} else if (sortSpecs[idx].dir === 'asc') {{
    sortSpecs[idx].dir = 'desc';
  }} else {{
    sortSpecs.splice(idx, 1);
  }}
  syncSortUI();
  applyFilters();
}}

function clearSort() {{
  // 重設：清空所有排序層
  sortSpecs = [];
  syncSortUI();
  applyFilters();
}}

function syncSortUI() {{
  // 每顆按鈕顯示自己在排序清單中的「層級」與方向，例如 1▲ / 2▼
  const levelMap = {{}};
  sortSpecs.forEach((s, i) => {{ levelMap[s.key] = {{ n: i + 1, dir: s.dir }}; }});
  document.querySelectorAll('.sort-btn').forEach(btn => {{
    const k = btn.getAttribute('data-key');
    const info = levelMap[k];
    btn.classList.toggle('active', !!info);
    btn.classList.toggle('sort-lv2', !!info && info.n > 1);  // 第 2 層以後用不同色
    const ind = btn.querySelector('.sort-ind');
    if (ind) ind.textContent = info ? (info.n + (info.dir === 'asc' ? '▲' : '▼')) : '⇅';
    btn.title = info
      ? ('第 ' + info.n + ' 層排序：' + k + (info.dir === 'asc' ? ' 升序' : ' 降序'))
      : ('加入排序：' + k);
  }});
  // 表頭指示器同步（顯示層級＋方向）
  document.querySelectorAll('#data-table thead .sort-ind').forEach(el => {{
    const k = el.getAttribute('data-key');
    const info = levelMap[k];
    el.textContent = info ? (info.n + (info.dir === 'asc' ? '▲' : '▼')) : '⇅';
  }});
}}

function updateSortIndicators() {{
  syncSortUI();
}}

function applyFilters() {{
  const fModel = getSelected('Model');
  const fMKT = getSelected('MKT');
  const fCPU = getSelected('CPU');
  const fGPU = getSelected('GPU');
  const fStage = getSelected('Stage');
  const fMPYear = getSelected('MPYear');
  const mpStart = normYm(document.getElementById('mp-start').value);
  const mpEnd = normYm(document.getElementById('mp-end').value);

  const filtered = RAW_ROWS.filter(r => {{
    if (fModel.length && !(r._Model||[]).some(v => fModel.includes(v))) return false;
    if (fMKT.length && !(r._MKT||[]).some(v => fMKT.includes(v))) return false;
    if (fCPU.length && !(r._CPU||[]).some(v => fCPU.includes(v))) return false;
    if (fGPU.length && !(r._GPU||[]).some(v => fGPU.includes(v))) return false;
    if (fStage.length && !(r._Stage||[]).some(v => fStage.includes(v))) return false;
    const rangeSet = !!(mpStart || mpEnd);
    if (!rangeSet && fMPYear.length && !(r._MPYear||[]).some(v => fMPYear.includes(v))) return false;
    if (rangeSet) {{
      const ym = mpMonth(r.MP);
      if (!ym) return false;
      if (mpStart && ym < mpStart) return false;
      if (mpEnd && ym > mpEnd) return false;
    }}
    return true;
  }});

  // 多層排序：依 sortSpecs 的順序逐層比較，第一個分出勝負的層級決定順序
  if (sortSpecs.length) {{
    filtered.sort((a, b) => {{
      for (const spec of sortSpecs) {{
        const va = sortVal(a, spec.key);
        const vb = sortVal(b, spec.key);
        const ea = va === '';
        const eb = vb === '';
        if (ea && eb) continue;          // 兩邊都空 → 看下一層
        if (ea) return 1;                // 空值一律排最後
        if (eb) return -1;
        const cmp = va.localeCompare(vb, 'zh-Hant', {{ numeric: true, sensitivity: 'base' }});
        if (cmp !== 0) return spec.dir === 'asc' ? cmp : -cmp;
      }}
      return 0;
    }});
  }}

  const tbody = document.getElementById('tbody');
  tbody.innerHTML = '';
  filtered.forEach(r => {{
    const tr = document.createElement('tr');
    const beforeMid = ['Model','MKT','CPU','GPU','NPM','Stage','Kickoff'];
    beforeMid.forEach(k => {{
      const td = document.createElement('td');
      if (['MKT','CPU','GPU','NPM','Stage'].includes(k)) {{
        td.innerHTML = r[k] || '';
      }} else {{
        td.textContent = r[k] || '';
      }}
      tr.appendChild(td);
    }});
    const midCols = ['DVT-start','EVT-start','MVT-start','ATS-start'];
    midCols.forEach(k => {{
      const td = document.createElement('td');
      td.textContent = r[k] || '';
      td.className = 'mid-milestone';
      tr.appendChild(td);
    }});
    const tdMP = document.createElement('td');
    tdMP.textContent = r['MP'] || '';
    tr.appendChild(tdMP);

    const tdStatus = document.createElement('td');
    tdStatus.className = 'status-cell';
    tdStatus.innerHTML = r['CurrentStatus'] || '';
    tr.appendChild(tdStatus);

    const tdHighlight = document.createElement('td');
    tdHighlight.className = 'highlight-cell';
    tdHighlight.innerHTML = r['Highlight'] || '';
    tr.appendChild(tdHighlight);

    tbody.appendChild(tr);
  }});
  applyMidVisibility();
  LAST_FILTERED = filtered;

  document.getElementById('stats').textContent = `顯示 ${{filtered.length}} 筆（已合併）/ 原始 ${{TOTAL_ORIGINAL}} 筆`;
}}

function clearFilters() {{
  for (const key of Object.keys(FILTER_META)) {{
    document.getElementById('search-' + key).value = '';
    document.querySelectorAll('.cb-' + key).forEach(cb => {{
      cb.checked = isDefaultChecked(key, cb.value);
    }});
    filterItems(key);
  }}
  document.getElementById('mp-start').value = '';
  document.getElementById('mp-end').value = '';
  applyFilters();
}}

/*__EXPORTJS__*/
/*__REPORTJS__*/
init();
</script>
</body>
</html>'''

    lib = read_text(LIB_JSZIP) + "\n" + read_text(LIB_XLSX)
    shotlib = read_text(LIB_H2C) + "\n" + read_text(LIB_CROPJS)

    html = html.replace("/*__CROPCSS__*/", read_text(LIB_CROPCSS))
    html = html.replace("<!--__CROPMODAL__-->", read_text(LIB_CROPHTML))
    html = html.replace("/*__REPORTCSS__*/", REPORT_CSS)
    html = html.replace("<!--__REPORTMODAL__-->", REPORT_MODAL)
    html = html.replace("/*__LIB__*/", lib)
    html = html.replace("/*__SHOTLIB__*/", shotlib)
    html = html.replace("/*__EXPORTJS__*/", EXPORT_JS)
    html = html.replace("/*__REPORTJS__*/", REPORT_JS)

    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Generated: {OUTPUT_HTML}")
    print(f"Rows: {len(spec_rows)} | Merged: {len(merged_rows)} | size(KB): {round(len(html) / 1024)}")


if __name__ == "__main__":
    main()
