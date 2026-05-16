import openpyxl, json, sys
from datetime import datetime, timedelta

sys.stdout.reconfigure(encoding='utf-8')

BASE = r"C:\Users\andreas.hirche@sap.com\OneDrive - SAP SE\HEP dashboard\Requirements"

wb = openpyxl.load_workbook(BASE + r"\MyWishlist.xlsx", data_only=True)
ws = wb.worksheets[0]
rows = list(ws.iter_rows(values_only=True))

def excel_date(v):
    if v is None: return None
    if isinstance(v, (int, float)):
        try: return (datetime(1899,12,30) + timedelta(days=int(v))).strftime('%Y-%m-%d')
        except: return None
    return str(v)

def score(v):
    if v is None: return None
    try: return int(v)
    except: return None

data = []
for r in rows[1:]:
    if r[0] is None: continue
    sc = [score(r[8]), score(r[9]), score(r[10]), score(r[11])]
    real = [v for v in sc if v is not None and v < 999]
    avg = round(sum(real)/len(real)) if real else None
    data.append({
        "id": r[0], "seq": r[1],
        "wish": str(r[2]).strip() if r[2] else "",
        "problem": str(r[3]).strip() if r[3] else "",
        "owner": str(r[4]).strip() if r[4] else "",
        "frm": str(r[5]).strip() if r[5] else "",
        "comment": str(r[6]).strip() if r[6] else "",
        "area": str(r[7]).strip() if r[7] else "",
        "andreas": sc[0], "marco": sc[1], "ingo": sc[2], "bettina": sc[3], "avg": avg,
        "jira_st": str(r[14]).strip() if r[14] else "",
        "planned": excel_date(r[15]),
        "jira": str(r[16]).strip() if r[16] else "",
        "wiki_st": str(r[17]).strip() if r[17] else "",
        "wiki": str(r[18]).strip() if r[18] else "",
    })

areas = sorted(set(d['area'] for d in data if d['area']))
DATA_J = json.dumps(data, ensure_ascii=False)
AREAS_J = json.dumps(areas)
COUNT = len(data)

HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>DC Requirements Wishlist</title>
<link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 32 32'><rect width='32' height='32' rx='7' fill='%230070f2'/><text x='16' y='22' font-family='system-ui' font-size='11' font-weight='900' fill='white' text-anchor='middle'>REQ</text></svg>">
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','72',sans-serif;background:#f0f2f5;color:#1d2d3e;min-height:100vh;padding:32px 28px}
a{color:#0070f2;text-decoration:none}
a:hover{text-decoration:underline}
.page-header{max-width:1400px;margin:0 auto 28px;display:flex;align-items:center;gap:16px;flex-wrap:wrap}
.logo{width:48px;height:48px;background:linear-gradient(135deg,#0070f2,#004fa3);border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:0.75rem;font-weight:900;color:#fff;letter-spacing:0.04em;flex-shrink:0;box-shadow:0 4px 14px #0070f233}
.page-header h1{font-size:1.35rem;font-weight:700;color:#1d2d3e}
.page-header p{font-size:0.8rem;color:#6b7c93;margin-top:2px}
.header-right{margin-left:auto;display:flex;align-items:center;gap:10px}
.count-badge{background:#fff;border:1px solid #d1d9e0;border-radius:20px;padding:4px 14px;font-size:0.78rem;color:#6b7c93;font-weight:600}
.controls{max-width:1400px;margin:0 auto 16px;display:flex;gap:10px;flex-wrap:wrap;align-items:center}
.search-wrap{position:relative;flex:1;min-width:220px}
.search-wrap input{width:100%;padding:8px 12px 8px 34px;border:1px solid #d1d9e0;border-radius:8px;font-size:0.88rem;background:#fff;color:#1d2d3e;outline:none;transition:border-color .2s,box-shadow .2s}
.search-wrap input:focus{border-color:#0070f2;box-shadow:0 0 0 3px #0070f220}
.search-icon{position:absolute;left:10px;top:50%;transform:translateY(-50%);color:#94a3b8;font-size:0.95rem;pointer-events:none}
.sort-select{padding:8px 12px;border:1px solid #d1d9e0;border-radius:8px;font-size:0.85rem;background:#fff;color:#1d2d3e;outline:none;cursor:pointer}
.chips{max-width:1400px;margin:0 auto 18px;display:flex;gap:6px;flex-wrap:wrap}
.chip{padding:5px 12px;border-radius:20px;font-size:0.73rem;font-weight:600;cursor:pointer;border:1.5px solid #d1d9e0;background:#fff;color:#4a5568;transition:all .15s;user-select:none;white-space:nowrap}
.chip:hover{border-color:#0070f2;color:#0070f2}
.chip.active{background:#0070f2;border-color:#0070f2;color:#fff}
.table-wrap{max-width:1400px;margin:0 auto;background:#fff;border-radius:14px;box-shadow:0 2px 10px rgba(0,0,0,0.07);overflow:hidden;border:1px solid #e2e8f0}
table{width:100%;border-collapse:collapse}
thead tr{background:linear-gradient(180deg,#f8fafc,#f1f5f9)}
th{padding:9px 12px;text-align:left;font-size:0.68rem;text-transform:uppercase;letter-spacing:0.07em;color:#6b7c93;border-bottom:1px solid #e2e8f0;font-weight:700;white-space:nowrap;cursor:pointer;user-select:none}
th:hover{color:#0070f2}
td{padding:7px 12px;font-size:0.84rem;border-bottom:1px solid #f1f5f9;vertical-align:middle}
tr:last-child td{border-bottom:none}
tr.data-row:hover td{background:#f5f8ff}
tr.data-row{cursor:pointer}
.id-cell{font-size:0.74rem;font-weight:700;color:#94a3b8;white-space:nowrap}
.seq-badge{display:inline-block;background:#f1f5f9;color:#475569;font-size:0.63rem;font-weight:700;padding:1px 5px;border-radius:4px;margin-left:4px}
.wish-cell{min-width:260px;max-width:420px}
.wish-title{font-weight:600;color:#1d2d3e;line-height:1.35}
.area-tag{display:inline-block;font-size:0.65rem;font-weight:700;padding:2px 8px;border-radius:10px;white-space:nowrap}
.score-cell{text-align:center;white-space:nowrap;min-width:44px}
.score-dot{display:inline-block;width:34px;height:22px;border-radius:5px;font-size:0.71rem;font-weight:700;line-height:22px;text-align:center}
.score-must{background:#fef3c7;color:#92400e;border:1px solid #fcd34d}
.score-high{background:#dcfce7;color:#166534;border:1px solid #86efac}
.score-mid{background:#dbeafe;color:#1e40af;border:1px solid #93c5fd}
.score-low{background:#f1f5f9;color:#64748b;border:1px solid #cbd5e1}
.score-none{color:#cbd5e1}
.avg-cell{text-align:center;min-width:54px}
.avg-bar-wrap{display:flex;flex-direction:column;align-items:center;gap:2px}
.avg-num{font-size:0.78rem;font-weight:700}
.avg-bar{height:4px;border-radius:2px;background:#e2e8f0;width:36px;overflow:hidden}
.avg-fill{height:100%;border-radius:2px;background:linear-gradient(90deg,#0070f2,#00b4d8)}
.status-wrap{display:flex;gap:4px;flex-wrap:wrap;min-width:72px}
.badge-jira{background:#e0f0ff;color:#0052cc;font-size:0.62rem;font-weight:700;padding:2px 7px;border-radius:4px}
.badge-wiki{background:#e8f5e9;color:#1b5e20;font-size:0.62rem;font-weight:700;padding:2px 7px;border-radius:4px}
.badge-st{background:#fce8ec;color:#b91c1c;font-size:0.61rem;padding:2px 6px;border-radius:4px;max-width:110px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.link-icon{display:inline-block;padding:3px 8px;border-radius:6px;border:1px solid #e2e8f0;font-size:0.67rem;font-weight:700;color:#0070f2;margin:1px 2px;transition:background .12s}
.link-icon:hover{background:#f0f7ff;text-decoration:none}
.planned-cell{font-size:0.74rem;color:#64748b;white-space:nowrap}
tr.detail-row td{padding:0;border-bottom:1px solid #e2e8f0;background:#f8fafc}
.detail-inner{padding:14px 20px 16px;display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:10px 24px}
.detail-inner.hidden{display:none}
.det-field label{font-size:0.63rem;font-weight:700;text-transform:uppercase;letter-spacing:0.08em;color:#94a3b8;display:block;margin-bottom:3px}
.det-field p{font-size:0.82rem;color:#334155;line-height:1.5}
.empty-state{text-align:center;padding:60px 20px;color:#94a3b8;font-size:0.9rem}
</style>
</head>
<body>
<header class="page-header">
  <div class="logo">REQ</div>
  <div>
    <h1>DC Requirements Wishlist</h1>
    <p>SAP Discovery Center &ndash; PM Wishlist</p>
  </div>
  <div class="header-right">
    <span class="count-badge" id="count-badge">""" + str(COUNT) + """ items</span>
  </div>
</header>
<div class="controls">
  <div class="search-wrap">
    <span class="search-icon">&#128269;</span>
    <input type="text" id="search" placeholder="Search wishes&hellip;" oninput="applyFilters()">
  </div>
  <select class="sort-select" id="sort-sel" onchange="applyFilters()">
    <option value="seq">Sort: Sequence</option>
    <option value="avg_desc">Sort: Avg score &darr;</option>
    <option value="id">Sort: Wish #</option>
    <option value="area">Sort: Area</option>
  </select>
</div>
<div class="chips" id="chips">
  <div class="chip active" data-area="" onclick="setArea(this)">All</div>
</div>
<div class="table-wrap">
  <table id="wish-table">
    <thead><tr>
      <th onclick="setSortCol('id')">Wish #</th>
      <th onclick="setSortCol('wish')">Wish</th>
      <th onclick="setSortCol('area')">Area</th>
      <th onclick="setSortCol('andreas')" title="Andreas">And.</th>
      <th onclick="setSortCol('marco')" title="Marco">Mar.</th>
      <th onclick="setSortCol('ingo')" title="Ingo">Ing.</th>
      <th onclick="setSortCol('bettina')" title="Bettina">Bet.</th>
      <th onclick="setSortCol('avg')">Avg</th>
      <th onclick="setSortCol('planned')">Planned</th>
      <th>Status</th>
      <th>Links</th>
    </tr></thead>
    <tbody id="tbody"></tbody>
  </table>
  <div class="empty-state" id="empty-state" style="display:none">No wishes match your filter.</div>
</div>

<script>
var DATA=""" + DATA_J + """;
var AREAS=""" + AREAS_J + """;
var activeArea='',sortCol='seq',sortDir=1,expandedId=null;
var COLORS={'DC general':'#dbeafe|#1d4ed8','DC Home':'#fce7f3|#9d174d','Mission Catalog':'#fef3c7|#92400e',
  'Project Board':'#dcfce7|#166534','Mission Overview':'#f3e8ff|#6b21a8','Mission Instances':'#fff7ed|#c2410c',
  'Mission Catalog / Service Catalog pricing':'#ecfdf5|#065f46','Mission Catalog / Mission Overview':'#fef9c3|#713f12',
  'Admin Center':'#f0fdf4|#14532d','Missions':'#ede9fe|#4c1d95','Reporting':'#fdf2f8|#701a75'};
function areaStyle(a){var c=COLORS[a];if(!c)return 'background:#f1f5f9;color:#475569';var p=c.split('|');return 'background:'+p[0]+';color:'+p[1];}
(function(){
  var ch=document.getElementById('chips');
  AREAS.forEach(function(a){
    var d=document.createElement('div');
    d.className='chip';d.dataset.area=a;d.textContent=a;
    d.onclick=function(){setArea(this);};
    ch.appendChild(d);
  });
})();
function setArea(el){document.querySelectorAll('.chip').forEach(function(c){c.classList.remove('active');});el.classList.add('active');activeArea=el.dataset.area;expandedId=null;applyFilters();}
function setSortCol(col){if(sortCol===col){sortDir*=-1;}else{sortCol=col;sortDir=1;}applyFilters();}
function scoreDot(v){
  if(v===null||v===undefined)return '<span class="score-none">&ndash;</span>';
  if(v>=999)return '<span class="score-dot score-must" title="Must have">M!</span>';
  if(v>=900)return '<span class="score-dot score-high">'+v+'</span>';
  if(v>=700)return '<span class="score-dot score-mid">'+v+'</span>';
  return '<span class="score-dot score-low">'+v+'</span>';
}
function avgCell(a){
  if(a===null||a===undefined)return '<span class="score-none">&ndash;</span>';
  var pct=Math.min(100,Math.round(a/9));
  var cls=a>=900?'color:#059669':a>=700?'color:#2563eb':'color:#6b7c93';
  return '<div class="avg-bar-wrap"><span class="avg-num" style="'+cls+'">'+a+'</span><div class="avg-bar"><div class="avg-fill" style="width:'+pct+'%"></div></div></div>';
}
function esc(s){if(!s)return '';return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function applyFilters(){
  var q=document.getElementById('search').value.toLowerCase().trim();
  var sopt=document.getElementById('sort-sel').value;
  var col=sopt==='avg_desc'?'avg':sopt,dir=sopt==='avg_desc'?-1:1;
  if(sortCol!=='seq'&&sortCol!=='avg'){col=sortCol;dir=sortDir;}
  var f=DATA.filter(function(d){
    if(activeArea&&d.area!==activeArea)return false;
    if(q){var h=[d.wish,d.problem,d.area,d.owner,String(d.id)].join(' ').toLowerCase();if(!h.includes(q))return false;}
    return true;
  });
  f.sort(function(a,b){
    var av=a[col],bv=b[col];
    if(av===null||av===undefined)av=dir>0?99999:-1;
    if(bv===null||bv===undefined)bv=dir>0?99999:-1;
    if(typeof av==='string')return dir*av.localeCompare(bv);
    return dir*(av-bv);
  });
  var tbody=document.getElementById('tbody');
  tbody.innerHTML='';
  document.getElementById('count-badge').textContent=f.length+' item'+(f.length!==1?'s':'');
  document.getElementById('empty-state').style.display=f.length?'none':'';
  document.getElementById('wish-table').style.display=f.length?'':'none';
  f.forEach(function(d){
    var tr=document.createElement('tr');tr.className='data-row';tr.dataset.id=d.id;
    var links='';
    if(d.jira)links+='<a class="link-icon" href="'+d.jira+'" target="_blank" rel="noopener" onclick="event.stopPropagation()">Jira &#8599;</a>';
    if(d.wiki)links+='<a class="link-icon" href="'+d.wiki+'" target="_blank" rel="noopener" onclick="event.stopPropagation()">Wiki &#8599;</a>';
    var st='<div class="status-wrap">';
    if(d.jira_st)st+='<span class="badge-st" title="'+esc(d.jira_st)+'">'+esc(d.jira_st.substring(0,22))+'</span>';
    else{if(d.jira)st+='<span class="badge-jira">Jira</span>';if(d.wiki)st+='<span class="badge-wiki">Wiki</span>';}
    st+='</div>';
    var pl=d.planned?'<span class="planned-cell">'+d.planned+'</span>':'<span class="score-none">&ndash;</span>';
    tr.innerHTML='<td class="id-cell">#'+d.id+'<span class="seq-badge">'+(d.seq||'&ndash;')+'</span></td>'
      +'<td class="wish-cell"><div class="wish-title">'+esc(d.wish)+'</div></td>'
      +'<td><span class="area-tag" style="'+areaStyle(d.area)+'">'+esc(d.area)+'</span></td>'
      +'<td class="score-cell">'+scoreDot(d.andreas)+'</td>'
      +'<td class="score-cell">'+scoreDot(d.marco)+'</td>'
      +'<td class="score-cell">'+scoreDot(d.ingo)+'</td>'
      +'<td class="score-cell">'+scoreDot(d.bettina)+'</td>'
      +'<td class="avg-cell">'+avgCell(d.avg)+'</td>'
      +'<td>'+pl+'</td>'
      +'<td>'+st+'</td>'
      +'<td style="white-space:nowrap">'+links+'</td>';
    tr.onclick=function(){toggleDetail(d,tr);};
    tbody.appendChild(tr);
    var dtr=document.createElement('tr');dtr.className='detail-row';
    var fields=[{l:'Challenge / Problem',v:d.problem},{l:'Owner',v:d.owner},{l:'Input from',v:d.frm},{l:'Comment',v:d.comment}].filter(function(x){return x.v;});
    var inner='<div class="detail-inner'+(expandedId===d.id?'':' hidden')+'" id="det-'+d.id+'">'
      +fields.map(function(x){return '<div class="det-field"><label>'+x.l+'</label><p>'+esc(x.v)+'</p></div>';}).join('')+'</div>';
    dtr.innerHTML='<td colspan="11">'+inner+'</td>';
    tbody.appendChild(dtr);
  });
}
function toggleDetail(d,tr){
  var inner=document.getElementById('det-'+d.id);if(!inner)return;
  if(inner.classList.contains('hidden')){
    document.querySelectorAll('.detail-inner:not(.hidden)').forEach(function(el){el.classList.add('hidden');});
    inner.classList.remove('hidden');expandedId=d.id;
  }else{inner.classList.add('hidden');expandedId=null;}
}
applyFilters();
</script>
</body>
</html>"""

out = BASE + r"\wishlist.html"
with open(out, 'w', encoding='utf-8') as f:
    f.write(HTML)
print(f"Written {len(HTML):,} bytes -> {out}")
print(f"  Wishlist: {COUNT} items")
