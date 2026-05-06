#!/usr/bin/env python3
"""
Build static public HTML from Excel and publish to GitHub Pages.

Run:   python build_public.py          # generate docs/index.html
       python build_public.py --push   # generate + git commit + git push
"""
import os, sys, io, re
from datetime import datetime, date
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

_DIR  = os.path.dirname(os.path.abspath(__file__))
DOCS  = os.path.join(_DIR, 'docs')
OUT   = os.path.join(DOCS, 'index.html')
sys.path.insert(0, _DIR)

from serve import (read_file_shared, EXCEL, CC_MAP, TOPIC_COLORS,
                   STATUS_COLS, STATUS_COLS_POST, safe, esc)

import html as _html
def _esc(s): return _html.escape(str(s), quote=True)

# ── topic color ──────────────────────────────────────────────────────────────
def get_color(topic):
    t = topic.lower()
    for k, v in TOPIC_COLORS.items():
        if k in t:
            return v
    return '#94a3b8'

def topic_label(topic):
    """Shorten topic for display."""
    t = topic.strip()
    # strip parenthetical suffixes like "(Option 2 for this date)"
    t = re.sub(r'\s*\(.*?\)\s*$', '', t).strip()
    return t

# ── flag SVG ─────────────────────────────────────────────────────────────────
FLAG_SVGS = {
    'de': '<rect width="5" height="1" fill="#000"/><rect width="5" height="1" y="1" fill="#D00"/><rect width="5" height="1" y="2" fill="#FFCE00"/>',
    'at': '<rect width="3" height="2" fill="#ED2939"/><rect width="3" height="0.667" y="0.667" fill="#fff"/>',
    'ch': '<rect width="32" height="32" fill="#FF0000"/><rect width="6" height="20" x="13" y="6" fill="#fff"/><rect width="20" height="6" x="6" y="13" fill="#fff"/>',
    'fr': '<rect width="3" height="2" fill="#002395"/><rect width="2" height="2" x="1" fill="#fff"/><rect width="1" height="2" x="2" fill="#ED2939"/>',
    'gb': '<rect width="60" height="30" fill="#012169"/><path d="M0,0 L60,30 M60,0 L0,30" stroke="#fff" stroke-width="6"/><path d="M0,0 L60,30 M60,0 L0,30" stroke="#C8102E" stroke-width="4"/><path d="M30,0 V30 M0,15 H60" stroke="#fff" stroke-width="10"/><path d="M30,0 V30 M0,15 H60" stroke="#C8102E" stroke-width="6"/>',
    'es': '<rect width="3" height="2" fill="#AA151B"/><rect width="3" height="1" y="0.5" fill="#F1BF00"/>',
    'it': '<rect width="3" height="2" fill="#009246"/><rect width="2" height="2" x="1" fill="#fff"/><rect width="1" height="2" x="2" fill="#CE2B37"/>',
    'nl': '<rect width="3" height="2" fill="#AE1C28"/><rect width="3" height="0.667" y="0.667" fill="#fff"/><rect width="3" height="0.667" y="1.333" fill="#21468B"/>',
    'be': '<rect width="3" height="2" fill="#000"/><rect width="1" height="2" x="1" fill="#FFD90C"/><rect width="1" height="2" x="2" fill="#F31830"/>',
    'se': '<rect width="5" height="4" fill="#006AA7"/><rect width="5" height="0.6" y="1.7" fill="#FECC02"/><rect width="0.6" height="4" x="1.5" fill="#FECC02"/>',
    'no': '<rect width="22" height="16" fill="#EF2B2D"/><rect width="22" height="4" y="6" fill="#fff"/><rect width="4" height="16" x="6" fill="#fff"/><rect width="22" height="2" y="7" fill="#002868"/><rect width="2" height="16" x="7" fill="#002868"/>',
    'fi': '<rect width="18" height="11" fill="#fff"/><rect width="18" height="3" y="4" fill="#003580"/><rect width="3" height="11" x="4" fill="#003580"/>',
    'dk': '<rect width="28" height="20" fill="#C60C30"/><rect width="28" height="4" y="8" fill="#fff"/><rect width="4" height="20" x="8" fill="#fff"/>',
    'pl': '<rect width="3" height="2" fill="#fff"/><rect width="3" height="1" y="1" fill="#DC143C"/>',
    'cz': '<rect width="3" height="2" fill="#fff"/><rect width="3" height="1" y="1" fill="#D7141A"/><path d="M0,0 L1.5,1 L0,2 Z" fill="#11457E"/>',
    'hu': '<rect width="3" height="2" fill="#CE2939"/><rect width="3" height="0.667" y="0.667" fill="#fff"/><rect width="3" height="0.667" y="1.333" fill="#477050"/>',
    'il': '<rect width="220" height="160" fill="#fff"/><rect width="220" height="30" y="20" fill="#0038b8"/><rect width="220" height="30" y="110" fill="#0038b8"/>',
    'ae': '<rect width="3" height="2" fill="#000"/><rect width="3" height="1.333" fill="#fff"/><rect width="3" height="0.667" fill="#00732F"/><rect width="1" height="2" fill="#FF0000"/>',
    'sa': '<rect width="3" height="2" fill="#006C35"/>',
    'za': '<rect width="3" height="2" fill="#007A4D"/>',
    'in': '<rect width="3" height="2" fill="#FF9933"/><rect width="3" height="0.667" y="0.667" fill="#fff"/><rect width="3" height="0.667" y="1.333" fill="#138808"/>',
    'sg': '<rect width="3" height="2" fill="#EF3340"/><rect width="3" height="1" y="1" fill="#fff"/>',
    'jp': '<rect width="3" height="2" fill="#fff"/><circle cx="1.5" cy="1" r="0.6" fill="#BC002D"/>',
    'cn': '<rect width="3" height="2" fill="#DE2910"/>',
    'au': '<rect width="3" height="2" fill="#00008B"/>',
    'us': '<rect width="3" height="2" fill="#B22234"/><rect width="3" height="0.154" y="0.154" fill="#fff"/><rect width="3" height="0.154" y="0.462" fill="#fff"/><rect width="3" height="0.154" y="0.769" fill="#fff"/><rect width="3" height="0.154" y="1.077" fill="#fff"/><rect width="3" height="0.154" y="1.385" fill="#fff"/><rect width="3" height="0.154" y="1.692" fill="#fff"/>',
    'ca': '<rect width="3" height="2" fill="#fff"/><rect width="0.75" height="2" fill="#FF0000"/><rect width="0.75" height="2" x="2.25" fill="#FF0000"/>',
    'mx': '<rect width="3" height="2" fill="#006847"/><rect width="1" height="2" x="1" fill="#fff"/><rect width="1" height="2" x="2" fill="#CE1126"/>',
    'br': '<rect width="7" height="5" fill="#009C3B"/><polygon points="3.5,0.5 6.5,2.5 3.5,4.5 0.5,2.5" fill="#FFDF00"/>',
    'ar': '<rect width="3" height="2" fill="#74ACDF"/><rect width="3" height="0.667" y="0.667" fill="#fff"/>',
    'co': '<rect width="3" height="2" fill="#CE1126"/><rect width="3" height="1.5" fill="#003893"/><rect width="3" height="1" fill="#FCD116"/>',
    'cl': '<rect width="3" height="2" fill="#D52B1E"/><rect width="3" height="1" fill="#fff"/><rect width="1" height="1" fill="#003087"/>',
    'pe': '<rect width="3" height="2" fill="#D91023"/><rect width="1" height="2" x="1" fill="#fff"/>',
    'kr': '<rect width="3" height="2" fill="#fff"/>',
    'tw': '<rect width="3" height="2" fill="#FE0000"/><rect width="1" height="1" fill="#000095"/>',
    'id': '<rect width="3" height="2" fill="#CE1126"/><rect width="3" height="1" y="1" fill="#fff"/>',
    'my': '<rect width="3" height="2" fill="#CC0001"/>',
    'ph': '<rect width="3" height="2" fill="#0038A8"/><rect width="3" height="1" y="1" fill="#CE1126"/>',
    'th': '<rect width="3" height="2" fill="#A51931"/><rect width="3" height="1.2" y="0.4" fill="#fff"/><rect width="3" height="0.4" y="0.8" fill="#2D2A4A"/>',
    'tr': '<rect width="3" height="2" fill="#E30A17"/>',
    'ua': '<rect width="3" height="2" fill="#005BBB"/><rect width="3" height="1" y="1" fill="#FFD500"/>',
    'sk': '<rect width="3" height="2" fill="#fff"/><rect width="3" height="0.667" y="0.667" fill="#0B4FCD"/><rect width="3" height="0.667" y="1.333" fill="#CE1021"/>',
    'hr': '<rect width="3" height="2" fill="#FF0000"/><rect width="3" height="0.667" y="0.667" fill="#fff"/><rect width="3" height="0.667" y="1.333" fill="#0093DD"/>',
    'ro': '<rect width="3" height="2" fill="#002B7F"/><rect width="1" height="2" x="1" fill="#FCD116"/><rect width="1" height="2" x="2" fill="#CE1126"/>',
    'bg': '<rect width="3" height="2" fill="#fff"/><rect width="3" height="0.667" y="0.667" fill="#00966E"/><rect width="3" height="0.667" y="1.333" fill="#D62612"/>',
    'gr': '<rect width="3" height="2" fill="#0D5EAF"/><rect width="3" height="0.222" y="0.222" fill="#fff"/>',
    'pt': '<rect width="3" height="2" fill="#006600"/><rect width="2" height="2" x="1" fill="#FF0000"/>',
    'ie': '<rect width="3" height="2" fill="#169B62"/><rect width="1" height="2" x="1" fill="#fff"/><rect width="1" height="2" x="2" fill="#FF883E"/>',
    'ke': '<rect width="3" height="2" fill="#006600"/>',
    'eg': '<rect width="3" height="2" fill="#CE1126"/><rect width="3" height="0.667" y="0.667" fill="#fff"/><rect width="3" height="0.667" y="1.333" fill="#000"/>',
    'vn': '<rect width="3" height="2" fill="#DA251D"/>',
}

def flag_svg(cc):
    if not cc or cc not in FLAG_SVGS:
        return ''
    is_wide = cc in ('gb', 'us', 'no', 'fi', 'dk', 'il')
    vb = '0 0 60 30' if cc == 'gb' else ('0 0 22 16' if cc == 'no' else ('0 0 18 11' if cc == 'fi' else ('0 0 28 20' if cc == 'dk' else ('0 0 220 160' if cc == 'il' else ('0 0 5 3' if cc in ('de','sk','pl','nl','be','se','ua','cl','pe','hr','ro','bg') else ('0 0 32 32' if cc == 'ch' else '0 0 3 2'))))))
    w, h = (22,13) if is_wide else (20,12)
    if cc == 'ch': w,h = 12,12; vb = '0 0 32 32'
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="{vb}" width="{w}" height="{h}" '
            f'style="border-radius:2px;vertical-align:middle;flex-shrink:0">'
            f'{FLAG_SVGS[cc]}</svg>')

# ── registration button ───────────────────────────────────────────────────────
def reg_btn(reg):
    if not reg: return '<span class="stay">stay tuned ›</span>'
    if reg.startswith('http'):
        return f'<a class="reg-btn" href="{_esc(reg)}" target="_blank" rel="noopener">Register ↗</a>'
    return f'<span class="stay">{_esc(reg)}</span>'

# ── data reading ──────────────────────────────────────────────────────────────
def get_cc(location):
    loc = location.lower().strip()
    for k, v in CC_MAP.items():
        if k in loc: return v
    return ''

def build_cards():
    TODAY = date.today()
    wb = openpyxl.load_workbook(io.BytesIO(read_file_shared(EXCEL)), data_only=True)

    def parse(ws, source):
        out = []
        for row_cells in list(ws.iter_rows())[1:]:
            def get(idx):
                try:
                    c = row_cells[idx]
                    if c is None: return None
                    if c.hyperlink and c.hyperlink.target:
                        return c.hyperlink.target
                    return c.value
                except IndexError:
                    return None
            topic = safe(get(0))
            if not topic: continue
            start = get(2)
            if isinstance(start, datetime): start = start.date()
            elif isinstance(start, str):
                try: start = datetime.strptime(start, '%Y-%m-%d').date()
                except: start = None
            phase = 'postponed' if source == 'postponed' else (
                'upcoming' if (start is None or start >= TODAY) else 'past')
            out.append({
                'topic':    topic_label(topic),
                'date':     safe(get(1)),
                'start':    start.isoformat() if start else '',
                'location': safe(get(3)),
                'region':   safe(get(4)),
                'language': safe(get(5)),
                'reg':      safe(get(19)),
                'phase':    phase,
            })
        return out

    rows = (
        sorted(parse(wb['2026 Shedule'], '2026'),      key=lambda w: w['start'] or '2099') +
        sorted(parse(wb['Postponed'],    'postponed'), key=lambda w: w['start'] or '2099')
    )
    return rows

# ── HTML generation ───────────────────────────────────────────────────────────
def build_html(rows):
    # Collect unique topics and regions for filter buttons
    topics   = []
    seen_t   = set()
    regions  = []
    seen_r   = set()
    for w in rows:
        t = w['topic']
        if t and t not in seen_t:
            topics.append(t); seen_t.add(t)
        r = w['region']
        if r and r not in seen_r:
            regions.append(r); seen_r.add(r)

    topic_filters = ''.join(
        f'<button class="filter-btn" data-col="topic" data-val="{_esc(t)}" '
        f'style="border-left:3px solid {get_color(t)}">{_esc(t)}</button>'
        for t in topics
    )
    region_filters = ''.join(
        f'<button class="filter-btn" data-col="region" data-val="{_esc(r)}">{_esc(r)}</button>'
        for r in regions
    )

    cards_html = ''
    for w in rows:
        cc    = get_cc(w['location'])
        color = get_color(w['topic'])
        flag  = flag_svg(cc)
        pcls  = {'upcoming':'','past':'row-past','postponed':'row-postponed'}[w['phase']]
        cards_html += (
            f'<div class="card {pcls}" '
            f'data-phase="{w["phase"]}" '
            f'data-topic="{_esc(w["topic"])}" '
            f'data-region="{_esc(w["region"])}" '
            f'data-lang="{_esc(w["language"].lower())}">'
            f'<div><span class="chip" style="background:{color}22;color:{color};border:1px solid {color}55">'
            f'{_esc(w["topic"])}</span></div>'
            f'<div class="loc-row">'
            f'<span class="loc-text">📍 {_esc(w["location"])}</span>'
            f'{flag}</div>'
            f'<div class="date">📅 {_esc(w["date"] or w["start"])}</div>'
            + (f'<div class="lang">🗣 {_esc(w["language"])}</div>' if w['language'] else '') +
            (f'<div class="phase-badge {w["phase"]}">{w["phase"]}</div>' if w['phase'] != "upcoming" else '') +
            f'<div class="reg-wrap">{reg_btn(w["reg"])}</div>'
            f'</div>\n'
        )

    generated = datetime.now().strftime('%Y-%m-%d %H:%M')
    return HTML_TEMPLATE.replace('{{TOPIC_FILTERS}}', topic_filters) \
                        .replace('{{REGION_FILTERS}}', region_filters) \
                        .replace('{{CARDS}}', cards_html) \
                        .replace('{{GENERATED}}', generated)


HTML_TEMPLATE = """\
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>SAP HEP Partner Workshops 2026</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{background:#0a1628;color:#e2e8f0;font-family:'Segoe UI',system-ui,sans-serif;min-height:100vh}
header{background:#0d1f35;border-bottom:1px solid #1e40af44;padding:18px 24px}
header h1{font-size:1.25rem;font-weight:700;color:#fff}
header .sub{font-size:0.82rem;color:#64748b;margin-top:3px}
.filters{padding:12px 24px;display:flex;flex-wrap:wrap;gap:8px;background:#0a1628;border-bottom:1px solid #1e3a5f22;align-items:center}
.filter-label{font-size:0.72rem;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:.04em;margin-right:2px}
.filter-btn{background:#0d1f35;border:1px solid #1e40af44;color:#94a3b8;padding:5px 13px;border-radius:20px;cursor:pointer;font-size:0.75rem;font-weight:600;transition:all .15s;white-space:nowrap}
.filter-btn:hover{border-color:#3b82f6;color:#e2e8f0}
.filter-btn.active{background:#1e3a5f;border-color:#3b82f6;color:#60a5fa}
.filter-sep{width:1px;height:20px;background:#1e3a5f;margin:0 4px}
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:14px;padding:20px 24px}
.card{background:#0d1f35;border:1px solid #1e3a5f44;border-radius:10px;padding:16px;display:flex;flex-direction:column;gap:8px;transition:border-color .15s,transform .1s}
.card:hover{border-color:#3b82f688;transform:translateY(-1px)}
.card.hidden{display:none}
.card.row-past{opacity:.4}
.card.row-postponed{border-color:#92400e66;background:#0d1f2a}
.chip{display:inline-block;padding:3px 10px;border-radius:10px;font-size:0.7rem;font-weight:700;letter-spacing:.02em}
.loc-row{display:flex;align-items:center;gap:8px;font-size:0.85rem;color:#94a3b8}
.loc-text{flex:1}
.date{font-size:0.82rem;color:#94a3b8}
.lang{font-size:0.78rem;color:#64748b}
.phase-badge{display:inline-block;font-size:0.68rem;font-weight:700;padding:2px 8px;border-radius:8px;width:fit-content}
.phase-badge.postponed{background:#451a0366;color:#fcd34d;border:1px solid #92400e}
.phase-badge.past{background:#1e293b;color:#475569}
.reg-wrap{margin-top:4px}
.reg-btn{display:inline-block;background:#1d4ed8;color:#fff;padding:6px 16px;border-radius:7px;text-decoration:none;font-size:0.78rem;font-weight:600;transition:background .15s}
.reg-btn:hover{background:#2563eb}
.stay{color:#475569;font-size:0.75rem}
.no-results{color:#334155;text-align:center;padding:48px;font-size:1rem;grid-column:1/-1}
footer{text-align:center;padding:20px 24px;color:#334155;font-size:0.72rem;border-top:1px solid #1e3a5f22;margin-top:8px}
@media(max-width:640px){.grid{padding:12px 14px;gap:10px}.filters{padding:10px 14px}}
</style>
</head>
<body>
<header>
  <h1>SAP HEP Partner Workshops 2026</h1>
  <div class="sub">Hands-on Excellence Program · upcoming partner workshops</div>
</header>
<div class="filters" id="filters">
  <span class="filter-label">Show</span>
  <button class="filter-btn active" data-action="upcoming">Upcoming</button>
  <button class="filter-btn" data-action="all">All (incl. past)</button>
  <div class="filter-sep"></div>
  <span class="filter-label">Topic</span>
  {{TOPIC_FILTERS}}
  <div class="filter-sep"></div>
  <span class="filter-label">Region</span>
  {{REGION_FILTERS}}
</div>
<div class="grid" id="grid">{{CARDS}}</div>
<footer>Data updated: {{GENERATED}}</footer>
<script>
const grid  = document.getElementById('grid');
const cards = Array.from(grid.querySelectorAll('.card'));
let showAll = false, activeTopic = '', activeRegion = '';

document.getElementById('filters').addEventListener('click', e => {
  const btn = e.target.closest('.filter-btn');
  if (!btn) return;
  const action = btn.dataset.action, col = btn.dataset.col, val = btn.dataset.val;
  if (action === 'upcoming' || action === 'all') {
    showAll = (action === 'all');
    document.querySelectorAll('[data-action]').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
  } else if (col === 'topic') {
    activeTopic = (activeTopic === val) ? '' : val;
    document.querySelectorAll('[data-col="topic"]').forEach(b => b.classList.toggle('active', b.dataset.val === activeTopic));
  } else if (col === 'region') {
    activeRegion = (activeRegion === val) ? '' : val;
    document.querySelectorAll('[data-col="region"]').forEach(b => b.classList.toggle('active', b.dataset.val === activeRegion));
  }
  applyFilters();
});

function applyFilters() {
  let vis = 0;
  cards.forEach(c => {
    const phase = c.dataset.phase;
    let show = showAll || phase === 'upcoming';
    if (activeTopic  && !c.dataset.topic.toLowerCase().includes(activeTopic.toLowerCase()))  show = false;
    if (activeRegion && c.dataset.region !== activeRegion) show = false;
    c.classList.toggle('hidden', !show);
    if (show) vis++;
  });
  let msg = document.getElementById('no-results');
  if (!vis) {
    if (!msg) { msg = document.createElement('div'); msg.id='no-results'; msg.className='no-results'; msg.textContent='No workshops match the current filters.'; grid.appendChild(msg); }
  } else if (msg) msg.remove();
}
applyFilters();
</script>
</body>
</html>
"""

# ── main ──────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    import subprocess
    os.makedirs(DOCS, exist_ok=True)

    print('Reading Excel ...')
    rows = build_cards()
    print(f'  {len(rows)} workshops found')

    html = build_html(rows)
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'Written: {OUT}')

    # create .nojekyll if missing
    nojekyll = os.path.join(DOCS, '.nojekyll')
    if not os.path.exists(nojekyll):
        open(nojekyll, 'w').close()

    if '--push' in sys.argv:
        cmds = [
            ['git', '-C', _DIR, 'add', 'docs/'],
            ['git', '-C', _DIR, 'commit', '-m', f'update: {datetime.now().strftime("%Y-%m-%d %H:%M")}'],
            ['git', '-C', _DIR, 'push'],
        ]
        for cmd in cmds:
            r = subprocess.run(cmd, capture_output=True, text=True)
            print(' '.join(cmd[2:]), ':', r.stdout.strip() or r.stderr.strip() or 'ok')
        print('Published.')
    else:
        print()
        print('To publish:  python build_public.py --push')
