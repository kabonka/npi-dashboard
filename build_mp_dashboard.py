# -*- coding: utf-8 -*-
"""Build a filterable HTML dashboard for MP变动记录.xlsx grouped by Model."""
import openpyxl, collections, json, os, re

# Runtime-relative paths so the whole NPI_vx folder can be copied to another PC and run directly.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE_DIR, "MP变动记录.xlsx")
SPEC = os.path.join(BASE_DIR, "Spec總表.xlsx")
OUT = os.path.join(BASE_DIR, "MP变动记录_dashboard.html")

wb = openpyxl.load_workbook(SRC, data_only=True)
ws = wb["MP变动记录"]
rows = list(ws.iter_rows(values_only=True))
hdr = list(rows[0])
data = []
for r in rows[1:]:
    if not any(c is not None for c in r):
        continue
    d = dict(zip(hdr, r))
    data.append(d)

def norm(v):
    if v is None:
        return ""
    return str(v).strip()

def norm_date(v):
    s = norm(v)
    if not s:
        return ""
    # "2026/09/14" -> "2026-09-14" for sorting
    return s.replace("/", "-")

def parse_modtime(v):
    s = norm(v)
    if not s:
        return ""
    return s  # already sortable as text "YYYY/MM/DD HH:MM:SS"

def yr_of(v):
    s = norm_date(v)
    return s[:4] if len(s) >= 4 else ""

# Build Stage2 lookup from Spec總表 (Model Name + MKT Name -> Stage)
spec_wb = openpyxl.load_workbook(SPEC, data_only=True)
spec_ws = spec_wb["Schedule"]
spec_hdr = list(next(spec_ws.iter_rows(min_row=1, max_row=1, values_only=True)))
stage2_map = {}
for r in spec_ws.iter_rows(min_row=2, values_only=True):
    if not any(c is not None for c in r):
        continue
    sd = dict(zip(spec_hdr, r))
    k = (norm(sd.get('Model Name')), norm(sd.get('MKT Name')))
    if k[0] and k not in stage2_map:
        stage2_map[k] = norm(sd.get('Stage'))

def make_rec(d, mt):
    hl = norm(d['Highlight'])
    hl_up = hl.upper()
    has_hl = bool(hl) and hl_up not in ('NA', 'N/A')
    delta = d['新MP-原MP(天)']
    try:
        delta = int(delta)
    except (TypeError, ValueError):
        delta = 0
    return {
        "model": norm(d['Model']),
        "mkt": norm(d['MKT Name']),
        "cpu": norm(d['CPU']),
        "gpu": norm(d['GPU']),
        "npm": norm(d['NPM']),
        "stage": norm(d['Stage']),
        "stage2": stage2_map.get((norm(d['Model']), norm(d['MKT Name'])), "(未匹配)"),
        "oldMP": norm_date(d['原MP']),
        "newMP": norm_date(d['新MP']),
        "delta": delta,
        "highlight": hl if has_hl else "",
        "hasHL": has_hl,
        "modTime": mt,
        "year": yr_of(d['新MP']) or yr_of(d['原MP']) or yr_of(d['修改时间']),
    }

# Raw records: every source row, NO dedup, NO Pending removal (for the "原始/不去重" view)
rawRecs = [make_rec(d, parse_modtime(d['修改时间'])) for d in data]

# Dedupe: within the same Model, if 原MP/新MP/Highlight are identical, keep ONLY the first occurrence.
# (The duplicate rows in the source are re-logs that share 原MP/新MP/Highlight but have DIFFERENT
#  修改时间, so 修改时间 is NOT part of the identity -> they collapse to the first logged row.
#  MKT and Stage are also excluded from the identity.)
seen = {}
for d in data:
    key = (norm(d['Model']), norm_date(d['原MP']), norm_date(d['新MP']),
           norm(d['Highlight']))
    if key not in seen:
        seen[key] = dict(d, _mt=parse_modtime(d['修改时间']))
deduped = [make_rec(d, d['_mt']) for d in seen.values()]

# remove Pending stage from the change detail (focus on actual MP changes)
deduped = [r for r in deduped if r['stage'] != 'Pending']

# stable pre-sort: model, mkt asc
deduped = sorted(deduped, key=lambda r: (r['model'], r['mkt']))

# per model aggregates (built from RAW so every model incl. all-Pending ones is listed)
models = sorted(set(r['model'] for r in rawRecs))
model_meta = {}
for m in models:
    recs = [r for r in rawRecs if r['model'] == m]
    mkt_names = sorted(set(r['mkt'] for r in recs))
    deltas = [r['delta'] for r in recs]
    model_meta[m] = {
        "count": len(recs),
        "mktCount": len(mkt_names),
        "mktNames": mkt_names,
        "stages": sorted(set(r['stage2'] for r in recs)),
        "avgDelta": round(sum(deltas)/len(deltas), 1) if deltas else 0,
        "delayed": sum(1 for x in deltas if x > 0),
        "pulled": sum(1 for x in deltas if x < 0),
        "withHL": sum(1 for r in recs if r['hasHL']),
        "lastMod": max((r['modTime'] for r in recs), default=""),
    }

# Stage2 下拉的「選項」= 資料裡實際出現的 Stage2 值（一路來自 Spec總表.xlsx 的 Stage 欄），
# Spec總表.xlsx 改了選項就跟著變，不寫死。未匹配到 Spec總表 的列會落在 "(未匹配)"。
stages_all = sorted(set(r['stage2'] for r in rawRecs))

# Theme keywords for highlight analysis
THEMES = [
    ("Schedule/時程", ["schedule", "時程", "排程", "ats", "start", "eta", "時間"]),
    ("備料/交料", ["備料", "交料", "料況", "料號", "缺料", "備貨", "pull"]),
    ("驗證/測試", ["驗證", "測試", "test", "validation", "試產"]),
    ("SW/Software", ["software", "sw", "bsp", "bios", "fw", "firmware", "程式"]),
    ("PCB", ["pcb", "主板", "板"]),
    ("SKU/機種", ["sku", "機種", "新機種", "variant", "機構", "黑 edition", "black edition"]),
    ("客戶/Customer", ["customer", "客戶", "leading", "oem", "brand"]),
    ("散熱/Thermal", ["散熱", "thermal", "fan", "風扇"]),
    ("良率/品質", ["良率", "yield", "品質", "quality", "defect", "良率"]),
    ("電源/Power", ["電源", "power", "battery", "充電", "adapter"]),
    ("設計/Design", ["設計", "design", "spec", "规格", "規格", "變更"]),
]
def theme_counts(recs):
    cnt = collections.Counter()
    for r in recs:
        if not r['hasHL']:
            continue
        h = r['highlight'].lower()
        for label, kws in THEMES:
            if any(kw.lower() in h for kw in kws):
                cnt[label] += 1
    return cnt

payload = {
    "records": deduped,
    "raw": rawRecs,
    "models": models,
    "modelMeta": model_meta,
    "stage2s": stages_all,
    "years": sorted(set(r['year'] for r in rawRecs if r['year'])),
    "themes": [t[0] for t in THEMES],
    "totalRaw": len(data),
    "totalDedup": len(deduped),
}

html = """<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="Expires" content="0">
<script>
// 强制绕过缓存：每次访问都请求最新版本
// （2026-10-03 使用者要求，與 TTM_dashboard.html 等一致）
// 注意：這段是 Python 純字串（非 f-string），所以 JS 的大括號直接寫單層即可
(function(){
  var ts = Date.now();
  var href = location.href.split('#')[0];
  if(href.indexOf('_nocache_') === -1){
    var sep = href.indexOf('?') === -1 ? '?' : '&';
    location.replace(href + sep + '_nocache_=' + ts);
  }
})();
</script>
<title>MP 變動記錄分析 Dashboard</title>
<style>
:root{
  --bg:#f4f6f8; --card:#ffffff; --ink:#1f2933; --sub:#647084; --line:#e3e8ee;
  --brand:#2563eb; --brand2:#0ea5e9; --red:#e0413e; --green:#1f9d55; --amber:#d98a00; --gray:#8a94a6;
  --chip:#eef2f7;
}
*{box-sizing:border-box}
body{margin:0;font-family:"Segoe UI","Microsoft JhengHei","PingFang TC",sans-serif;background:var(--bg);color:var(--ink);font-size:14px}
header{background:linear-gradient(120deg,#1e3a8a,#2563eb);color:#fff;padding:18px 24px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px}
header h1{margin:0;font-size:20px;font-weight:700}
header .meta{font-size:12px;opacity:.9}
.wrap{display:flex;gap:16px;padding:16px;align-items:flex-start;max-width:1500px;margin:0 auto}
.side{width:280px;flex:0 0 280px;position:sticky;top:12px}
.main{flex:1;min-width:0}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px;margin-bottom:14px;box-shadow:0 1px 2px rgba(16,24,40,.04)}
.card h3{margin:0 0 10px;font-size:14px;color:var(--sub);text-transform:uppercase;letter-spacing:.4px}
/* view tabs */
.tabs{display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap}
.tab{padding:9px 16px;border:1px solid var(--line);background:#fff;border-radius:8px;cursor:pointer;font-size:13px;font-weight:600;color:var(--sub)}
.tab:hover{background:#f1f5f9}
.tab.active{background:var(--brand);color:#fff;border-color:var(--brand)}
.dl-btn{margin-left:auto;padding:9px 16px;border:1px solid var(--brand);background:#fff;color:var(--brand);border-radius:8px;cursor:pointer;font-size:13px;font-weight:600}
.dl-btn:hover{background:#eff6ff}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:10px;margin-bottom:14px}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:12px;text-align:center}
.kpi .v{font-size:24px;font-weight:800;line-height:1.1}
.kpi .l{font-size:11px;color:var(--sub);margin-top:4px}
.kpi.red .v{color:var(--red)} .kpi.green .v{color:var(--green)} .kpi.blue .v{color:var(--brand)}
/* model dropdown */
.ms{position:relative}
.ms-input{width:100%;padding:9px 10px;border:1px solid var(--line);border-radius:8px;background:#fff;cursor:pointer;display:flex;justify-content:space-between;align-items:center;font-size:13px}
.ms-input .cnt{color:var(--brand);font-weight:600}
.ms-panel{position:absolute;z-index:40;top:calc(100% + 4px);left:0;right:0;background:#fff;border:1px solid var(--line);border-radius:10px;box-shadow:0 8px 24px rgba(16,24,40,.16);max-height:340px;overflow:auto;display:none}
.ms-panel.open{display:block}
.ms-search{padding:8px;border-bottom:1px solid var(--line);position:sticky;top:0;background:#fff}
.ms-search input{width:100%;padding:7px 9px;border:1px solid var(--line);border-radius:7px;font-size:13px}
.ms-actions{display:flex;gap:6px;padding:8px;border-bottom:1px solid var(--line)}
.ms-actions button{flex:1;font-size:12px;padding:6px;border:1px solid var(--line);background:var(--chip);border-radius:7px;cursor:pointer}
.ms-actions button:hover{background:#e2e8f0}
.ms-opt{display:flex;align-items:center;gap:8px;padding:7px 10px;cursor:pointer;font-size:13px}
.ms-opt:hover{background:#f1f5f9}
.ms-opt .mc{margin-left:auto;font-size:11px;color:var(--gray)}
.ms-opt input{accent-color:var(--brand)}
select.filter,input.search{width:100%;padding:8px 9px;border:1px solid var(--line);border-radius:8px;font-size:13px;background:#fff}
.toggle{display:flex;align-items:center;gap:8px;font-size:13px;margin-top:4px;cursor:pointer}
/* model group */
.group{background:var(--card);border:1px solid var(--line);border-radius:10px;margin-bottom:14px;overflow:hidden}
.group-head{display:flex;align-items:center;gap:12px;padding:12px 14px;cursor:pointer;background:#f8fafc;border-bottom:1px solid var(--line)}
.group-head .gname{font-size:16px;font-weight:800}
.group-head .gtags{display:flex;gap:6px;flex-wrap:wrap;font-size:11px}
.gtag{background:var(--chip);color:var(--sub);padding:2px 8px;border-radius:20px}
.group-head .gcaret{margin-left:auto;color:var(--sub);transition:transform .15s}
.group.collapsed .gcaret{transform:rotate(-90deg)}
.group.collapsed .group-body{display:none}
.group-body{padding:6px 10px 10px}
table{width:100%;border-collapse:collapse;font-size:13px}
th,td{padding:7px 8px;text-align:left;border-bottom:1px solid #eef1f5;vertical-align:top}
th{background:#f8fafc;color:var(--sub);font-weight:600;font-size:12px;position:sticky;top:0}
tr:hover td{background:#fafcff}
.mname{font-weight:600;color:#0f172a}
.delta{font-weight:700;text-align:center;min-width:52px}
.delta.pos{color:var(--red)} .delta.neg{color:var(--green)} .delta.zero{color:var(--gray)}
.hl{color:#334155;max-width:520px;line-height:1.45}
.hl.empty{color:var(--gray);font-style:italic}
.stage{display:inline-block;padding:1px 8px;border-radius:6px;font-size:11px;font-weight:700;color:#fff;white-space:nowrap}
.muted{color:var(--gray);font-size:12px}
/* highlight analysis */
.hl-analysis .themes{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px}
.theme-chip{background:var(--chip);border:1px solid var(--line);border-radius:20px;padding:5px 12px;font-size:12px;cursor:pointer}
.theme-chip.active{background:var(--brand);color:#fff;border-color:var(--brand)}
.theme-chip .n{font-weight:700;margin-left:4px}
.hl-list{display:flex;flex-direction:column;gap:8px}
.hl-item{border:1px solid var(--line);border-radius:8px;padding:9px 11px;background:#fff}
.hl-item .h-top{display:flex;gap:10px;font-size:12px;color:var(--sub);margin-bottom:4px;flex-wrap:wrap}
.hl-item .h-txt{color:#1f2933;line-height:1.5}
mark{background:#fff3a3;padding:0 2px;border-radius:3px}
.foot{color:var(--gray);font-size:12px;padding:0 16px 24px;text-align:center}
.empty-state{padding:40px;text-align:center;color:var(--gray)}
@media(max-width:880px){.wrap{flex-direction:column}.side{width:100%;position:static}}
</style>
</head>
<body>
<header>
  <div>
    <h1>MP 變動記錄分析 Dashboard</h1>
    <div class="meta">依 Model 整理變更項目與 Highlight 說明 · 來源：MP變動記錄.xlsx</div>
  </div>
  <div class="meta" id="srcMeta"></div>
</header>
<div class="wrap">
  <aside class="side">
    <div class="card">
      <h3>篩選 Model</h3>
      <div class="ms">
        <div class="ms-input" id="msInput"><span id="msLabel">全部 Model</span><span class="cnt" id="msCnt"></span></div>
        <div class="ms-panel" id="msPanel">
          <div class="ms-search"><input id="msSearch" placeholder="搜尋 Model..."></div>
          <div class="ms-actions">
            <button id="msAll">全選</button>
            <button id="msNone">清除</button>
          </div>
          <div id="msList"></div>
        </div>
      </div>
    </div>
    <div class="card">
      <h3>其他篩選</h3>
      <label class="muted" style="font-size:12px">Stage2（可多選，依 Spec總表）</label>
      <div class="ms" id="msStageWrap">
        <div class="ms-input" id="msStageInput"><span id="msStageLabel">全部 Stage2</span><span class="cnt" id="msStageCnt"></span></div>
        <div class="ms-panel" id="msStagePanel">
          <div class="ms-search"><input id="msStageSearch" placeholder="搜尋 Stage2..."></div>
          <div class="ms-actions"><button id="msStageAll">全選</button><button id="msStageNone">清除</button></div>
          <div id="msStageList"></div>
        </div>
      </div>
      <label class="muted" style="font-size:12px;display:block;margin-top:10px">MP 年份（可多選，依新MP）</label>
      <div class="ms" id="msYearWrap">
        <div class="ms-input" id="msYearInput"><span id="msYearLabel">全部年份</span><span class="cnt" id="msYearCnt"></span></div>
        <div class="ms-panel" id="msYearPanel">
          <div class="ms-actions"><button id="msYearAll">全選</button><button id="msYearNone">清除</button></div>
          <div id="msYearList"></div>
        </div>
      </div>
      <label class="toggle"><input type="checkbox" id="onlyHL"> 只顯示有 Highlight</label>
      <label class="muted" style="font-size:12px;display:block;margin-top:10px">Highlight 關鍵字搜尋</label>
      <input class="search" id="kwSearch" placeholder="例如 Schedule / 備料 / SKU...">
    </div>
    <div class="card">
      <h3>圖例</h3>
      <div class="muted" style="font-size:12px;line-height:1.7">
        <span style="color:var(--red);font-weight:700">▲ 紅</span> = 延後 (Δ&gt;0，新MP晚於原MP)<br>
        <span style="color:var(--green);font-weight:700">▼ 綠</span> = 提前 (Δ&lt;0，新MP早於原MP)<br>
        <span style="color:var(--gray);font-weight:700">0</span> = 不變<br>
        Δ = 新MP − 原MP (天)
      </div>
    </div>
  </aside>
  <main class="main">
    <div class="tabs" id="tabs">
      <button class="tab active" data-v="dedup">變更明細（去重）</button>
      <button class="tab" data-v="raw">變更明細（原始 / 不去重）</button>
      <button class="dl-btn" id="dlExcel">下載 Excel（去重）</button>
    </div>
    <div class="kpis" id="kpis"></div>
    <div class="card">
      <h3>變更明細（依 Model 分組） <span id="detailSub" class="muted"></span></h3>
      <div id="groups"></div>
    </div>
    <div class="card hl-analysis">
      <h3>Highlight 說明分析</h3>
      <div class="themes" id="themes"></div>
      <input class="search" id="hlSearch" placeholder="在 Highlight 中搜尋關鍵字..." style="margin-bottom:10px">
      <div class="hl-list" id="hlList"></div>
    </div>
  </main>
</div>
<div class="foot" id="foot"></div>
<script>
const DATA = __DATA__;
const MODELS = DATA.models, MM = DATA.modelMeta, STAGE2S = DATA.stage2s, YEARS = DATA.years, THEMES = DATA.themes;
const DEDU = DATA.records, RAW = DATA.raw;
const stageColor = {MVT:'#7c3aed',DVT:'#2563eb',EVT:'#0891b2',ATS:'#0d9488','ATS-':'#0e7490',MP:'#dc2626','BTO-':'#ca8a04',Design:'#4f46e5',Pending:'#64748b','(未匹配)':'#94a3b8'};
// STAGE2S 的「選項」一律來自 Spec總表.xlsx 的實際 Stage 值（不寫死）。
// 這裡只決定「開頁面時預設勾選哪一個」——本表聚焦 MP 階段的變動，所以預設只勾 MP。
let state = { view:'dedup', models:new Set(MODELS), stage2s:new Set(STAGE2S.filter(s=>s==='MP')), years:new Set(YEARS), onlyHL:false, kw:'', theme:null, hlkw:'' };

function esc(s){return (s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}
function fmtDate(s){return s? s.replace(/-/g,'/'):'';}

function activeRecs(){ return state.view==='raw'? RAW : DEDU; }
function filtered(){
  return activeRecs().filter(r=>
    state.models.has(r.model) &&
    state.stage2s.has(r.stage2) &&
    (state.years.size===0 || state.years.has(r.year)) &&
    (!state.onlyHL || r.hasHL) &&
    (!state.kw || r.mkt.toLowerCase().includes(state.kw.toLowerCase()) || r.model.toLowerCase().includes(state.kw.toLowerCase())) &&
    (!state.theme || (r.hasHL && themeMatch(r.highlight, state.theme))) &&
    (!state.hlkw || (r.hasHL && r.highlight.toLowerCase().includes(state.hlkw.toLowerCase())))
  );
}
const THEME_KW = __THEMEKW__;
function themeMatch(h,label){const kws=THEME_KW[label]||[];const t=h.toLowerCase();return kws.some(k=>t.includes(k.toLowerCase()));}

function renderKPIs(){
  const f=filtered(); const all=f.length;
  const delayed=f.filter(r=>r.delta>0).length;
  const pulled=f.filter(r=>r.delta<0).length;
  const withHL=f.filter(r=>r.hasHL).length;
  const md=f.length? (f.reduce((a,r)=>a+r.delta,0)/f.length):0;
  const mset=new Set(f.map(r=>r.model)).size;
  const kpis=[
    {v:mset,l:'Model 數',c:''},
    {v:all,l:'變更記錄',c:'blue'},
    {v:withHL,l:'含 Highlight',c:''},
    {v:md.toFixed(1),l:'平均 Δ(天)',c:''},
    {v:delayed,l:'延後(Δ>0)',c:'red'},
    {v:pulled,l:'提前(Δ<0)',c:'green'},
  ];
  document.getElementById('kpis').innerHTML=kpis.map(k=>
    `<div class="kpi ${k.c}"><div class="v">${k.v}</div><div class="l">${k.l}</div></div>`).join('');
}

function renderGroups(){
  const f=filtered();
  const byModel={};
  f.forEach(r=>{(byModel[r.model]=byModel[r.model]||[]).push(r);});
  const order=Object.keys(byModel).sort((a,b)=>byModel[b].length-byModel[a].length);
  const box=document.getElementById('groups');
  if(!order.length){box.innerHTML='<div class="empty-state">無符合條件的資料</div>';return;}
  box.innerHTML=order.map(m=>{
    // within a Model: sort by new MP time old -> new (empty newMP to the end)
    const list=byModel[m].slice().sort((a,b)=>{
      const ae=!a.newMP, be=!b.newMP;
      if(ae||be){ if(ae&&be) return 0; return ae?1:-1; }
      if(a.newMP===b.newMP) return 0;
      return a.newMP<b.newMP?-1:1;
    });
    const meta={
      count:list.length,
      mktCount:new Set(list.map(r=>r.mkt)).size,
      avgDelta:list.length?(list.reduce((a,r)=>a+r.delta,0)/list.length):0,
      delayed:list.filter(r=>r.delta>0).length,
      pulled:list.filter(r=>r.delta<0).length,
      lastMod:list.length?list.reduce((a,r)=>r.modTime>a?r.modTime:a,''):''
    };
    const rowsHtml=list.map(r=>{
        const dc=r.delta>0?`pos`:r.delta<0?`neg`:`zero`;
        const sign=r.delta>0?'+':'';
        const hl=r.hasHL?`<div class="hl">${esc(r.highlight)}</div>`:`<div class="hl empty">— 無說明 —</div>`;
        const sc=stageColor[r.stage]||'#64748b';
        return `<tr>
          <td class="mname">${esc(r.mkt)}</td>
          <td>${esc(r.cpu)}</td>
          <td>${esc(r.gpu)}</td>
          <td>${esc(r.stage2)}</td>
          <td><span class="stage" style="background:${sc}">${esc(r.stage)}</span></td>
          <td>${fmtDate(r.oldMP)}</td>
          <td>${fmtDate(r.newMP)}</td>
          <td class="delta ${dc}">${sign}${r.delta}</td>
          <td>${fmtDate(r.modTime)}</td>
          <td>${hl}</td>
        </tr>`;
    }).join('');
    return `<div class="group">
      <div class="group-head" onclick="this.parentNode.classList.toggle('collapsed')">
        <span class="gname">${esc(m)}</span>
        <span class="gtags">
          <span class="gtag">${list.length} 筆</span>
          <span class="gtag">${meta.mktCount} MKT</span>
          <span class="gtag">均Δ ${meta.avgDelta}</span>
          <span class="gtag" style="color:var(--red)">延${meta.delayed}</span>
          <span class="gtag" style="color:var(--green)">提${meta.pulled}</span>
          <span class="gtag">最後更新 ${fmtDate(meta.lastMod)||'—'}</span>
        </span>
        <span class="gcaret">▼</span>
      </div>
      <div class="group-body">
        <table>
          <thead><tr><th>MKT Name</th><th>CPU</th><th>GPU</th><th>Stage2</th><th>Stage</th><th>原MP</th><th>新MP</th><th>Δ天</th><th>修改時間</th><th>Highlight 說明</th></tr></thead>
          <tbody>${rowsHtml}</tbody>
        </table>
        <div class="muted" style="padding:4px 8px 2px">▲ 依 新MP 時間由舊到新排列</div>
      </div>
    </div>`;
  }).join('');
}

function renderThemes(){
  const f=filtered().filter(r=>r.hasHL);
  const cnt={}; THEMES.forEach(t=>cnt[t]=0);
  f.forEach(r=>{THEMES.forEach(t=>{if(themeMatch(r.highlight,t))cnt[t]++;});});
  const box=document.getElementById('themes');
  box.innerHTML=THEMES.map(t=>{
    const n=cnt[t]||0; if(!n) return '';
    return `<span class="theme-chip ${state.theme===t?'active':''}" data-t="${esc(t)}" onclick="toggleTheme('${esc(t)}')">${esc(t)}<span class="n">${n}</span></span>`;
  }).join('') || '<span class="muted">無 Highlight 資料</span>';
}
function toggleTheme(t){state.theme = state.theme===t?null:t; renderAll();}

function renderHLList(){
  const f=filtered().filter(r=>r.hasHL);
  const kw=state.hlkw.toLowerCase();
  const box=document.getElementById('hlList');
  if(!f.length){box.innerHTML='<div class="empty-state">無 Highlight 資料</div>';return;}
  const html=f.slice().sort((a,b)=>b.modTime>a.modTime?1:-1).map(r=>{
    let txt=esc(r.highlight);
    if(kw){const re=new RegExp('('+kw.replace(/[\\\\^$.*+?()[\\\\]{}|]/g,'\\\\$&')+')','gi');txt=txt.replace(re,'<mark>$1</mark>');}
    const sc=stageColor[r.stage]||'#64748b';
    return `<div class="hl-item">
      <div class="h-top"><b>${esc(r.model)}</b> · ${esc(r.mkt)} · <span class="stage" style="background:${sc}">${esc(r.stage)}</span> · ${fmtDate(r.modTime)}</div>
      <div class="h-txt">${txt}</div>
    </div>`;
  }).join('');
  box.innerHTML=html;
}

function renderAll(){
  renderKPIs(); renderGroups(); renderThemes(); renderHLList();
  document.getElementById('detailSub').textContent = state.view==='raw'? '· 原始（不去重）' : '· 去重';
  document.getElementById('foot').textContent=`原始 ${DATA.totalRaw} 筆 → 去重後 ${DATA.totalDedup} 筆；目前顯示 ${filtered().length} 筆（${state.view==='raw'?'原始 / 不去重':'去重'}；資料擷取時間 ${new Date().toLocaleString('zh-TW')}）`;
}

/* model multiselect */
const msList=document.getElementById('msList');
function buildMs(){
  msList.innerHTML=MODELS.map(m=>`<label class="ms-opt"><input type="checkbox" value="${esc(m)}" ${state.models.has(m)?'checked':''}><span>${esc(m)}</span><span class="mc">${MM[m].count}</span></label>`).join('');
}
function syncMs(){
  const n=state.models.size, total=MODELS.length;
  document.getElementById('msLabel').textContent = n===total?'全部 Model':(n===0?'未選擇':`已選 ${n} 個`);
  document.getElementById('msCnt').textContent = n===total?'':`${n}/${total}`;
  msList.querySelectorAll('input').forEach(cb=>{cb.checked=state.models.has(cb.value);});
}
document.getElementById('msInput').onclick=()=>document.getElementById('msPanel').classList.toggle('open');
document.getElementById('msSearch').oninput=e=>{
  const q=e.target.value.toLowerCase();
  msList.querySelectorAll('.ms-opt').forEach(o=>{o.style.display=o.textContent.toLowerCase().includes(q)?'':'none';});
};
msList.onchange=e=>{
  if(e.target.tagName==='INPUT'){const v=e.target.value; if(e.target.checked)state.models.add(v);else state.models.delete(v); syncMs(); renderAll();}
};
document.getElementById('msAll').onclick=()=>{MODELS.forEach(m=>state.models.add(m));syncMs();renderAll();};
document.getElementById('msNone').onclick=()=>{state.models.clear();syncMs();renderAll();};
document.addEventListener('click',e=>{if(!e.target.closest('.ms'))document.getElementById('msPanel').classList.remove('open');});

/* other filters */
document.getElementById('onlyHL').onchange=e=>{state.onlyHL=e.target.checked;renderAll();};
document.getElementById('kwSearch').oninput=e=>{state.kw=e.target.value.trim();renderAll();};
document.getElementById('hlSearch').oninput=e=>{state.hlkw=e.target.value.trim();renderHLList();};

/* generic multi-select for Stage / Year (modeled on npi_dashboard2) */
function buildMsFilter(wrapId, inputId, labelId, cntId, panelId, listId, allId, noneId, searchId, items, stateSet, allText){
  const list=document.getElementById(listId);
  list.innerHTML=items.map(it=>`<label class="ms-opt"><input type="checkbox" value="${esc(it)}" ${stateSet.has(it)?'checked':''}><span>${esc(it)}</span></label>`).join('');
  function sync(){
    const n=stateSet.size, total=items.length;
    document.getElementById(labelId).textContent = n===total?allText:(n===0?'未選擇':`已選 ${n}`);
    document.getElementById(cntId).textContent = n===total?'':`${n}/${total}`;
    list.querySelectorAll('input').forEach(cb=>{cb.checked=stateSet.has(cb.value);});
  }
  document.getElementById(inputId).onclick=()=>document.getElementById(panelId).classList.toggle('open');
  if(searchId){document.getElementById(searchId).oninput=e=>{const q=e.target.value.toLowerCase();list.querySelectorAll('.ms-opt').forEach(o=>{o.style.display=o.textContent.toLowerCase().includes(q)?'':'none';});};}
  list.onchange=e=>{if(e.target.tagName==='INPUT'){const v=e.target.value;if(e.target.checked)stateSet.add(v);else stateSet.delete(v);sync();renderAll();}};
  document.getElementById(allId).onclick=()=>{items.forEach(it=>stateSet.add(it));sync();renderAll();};
  document.getElementById(noneId).onclick=()=>{stateSet.clear();sync();renderAll();};
  document.addEventListener('click',e=>{if(!e.target.closest('#'+wrapId))document.getElementById(panelId).classList.remove('open');});
  sync();
}
buildMsFilter('msStageWrap','msStageInput','msStageLabel','msStageCnt','msStagePanel','msStageList','msStageAll','msStageNone','msStageSearch', STAGE2S, state.stage2s, '全部 Stage2');
buildMsFilter('msYearWrap','msYearInput','msYearLabel','msYearCnt','msYearPanel','msYearList','msYearAll','msYearNone', null, YEARS, state.years, '全部年份');

/* download Excel (client-side XML Spreadsheet 2003) */
function xesc(s){return (s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));}
function downloadExcel(){
  const f=filtered();
  if(!f.length){alert('目前無符合條件的資料可匯出');return;}
  const byModel={};
  f.forEach(r=>{(byModel[r.model]=byModel[r.model]||[]).push(r);});
  const order=Object.keys(byModel).sort((a,b)=>byModel[b].length-byModel[a].length);
  const rows=[];
  order.forEach(m=>{
    const list=byModel[m].slice().sort((a,b)=>{
      const ae=!a.newMP, be=!b.newMP;
      if(ae||be){ if(ae&&be) return 0; return ae?1:-1; }
      if(a.newMP===b.newMP) return 0;
      return a.newMP<b.newMP?-1:1;
    });
    rows.push(list);
  });
  const flat=rows.flat();
  const headers=[['Model','MKT Name','CPU','GPU','Stage2','Stage','原MP','新MP','Δ天','修改時間','Highlight']];
  const dataRows=flat.map(r=>[
    r.model, r.mkt, r.cpu, r.gpu, r.stage2, r.stage,
    fmtDate(r.oldMP), fmtDate(r.newMP), r.delta, fmtDate(r.modTime), r.highlight
  ]);
  let xml='<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?>';
  xml+='<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">';
  xml+='<Worksheet ss:Name="MP變動記錄"><Table>';
  headers.concat(dataRows).forEach(row=>{
    xml+='<Row>';
    row.forEach((cell,idx)=>{
      const isNum=idx===8 && typeof cell==='number';
      const t=isNum?'Number':'String';
      xml+=`<Cell><Data ss:Type="${t}">${xesc(String(cell))}</Data></Cell>`;
    });
    xml+='</Row>';
  });
  xml+='</Table></Worksheet></Workbook>';
  const blob=new Blob([xml],{type:'application/vnd.ms-excel'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a');
  a.href=url;
  const tag=state.view==='raw'?'原始':'去重';
  const d=new Date().toISOString().slice(0,10).replace(/-/g,'');
  a.download=`MP變動記錄_${tag}_${d}.xlsx`;
  document.body.appendChild(a); a.click();
  setTimeout(()=>{URL.revokeObjectURL(url);a.remove();},200);
}
document.getElementById('dlExcel').onclick=downloadExcel;

/* view tabs */
document.querySelectorAll('#tabs .tab').forEach(b=>{
  b.onclick=()=>{
    state.view=b.dataset.v;
    document.querySelectorAll('#tabs .tab').forEach(x=>x.classList.toggle('active', x===b));
    document.getElementById('dlExcel').textContent='下載 Excel（'+(state.view==='raw'?'原始':'去重')+'）';
    renderAll();
  };
});

document.getElementById('srcMeta').textContent=`共 ${DATA.totalDedup} 筆(去重) / ${RAW.length} 筆(原始) / ${MODELS.length} Models · Stage2 來自 Spec總表.xlsx`;
buildMs(); syncMs(); renderAll();
</script>
</body>
</html>"""

# theme keywords map for JS
theme_kw_map = {label: kws for label, kws in THEMES}
html = html.replace("__DATA__", json.dumps(payload, ensure_ascii=False))
html = html.replace("__THEMEKW__", json.dumps(theme_kw_map, ensure_ascii=False))

with open(OUT, "w", encoding="utf-8") as f:
    f.write(html)

print("WROTE:", OUT)
print("raw:", len(data), "dedup:", len(deduped), "models:", len(models))
print("themes sample:", theme_counts(deduped).most_common(5))
