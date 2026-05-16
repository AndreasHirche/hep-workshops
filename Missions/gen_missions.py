import re, json, sys
from html.parser import HTMLParser

sys.stdout.reconfigure(encoding='utf-8')

BASE = r"C:\Users\andreas.hirche@sap.com\OneDrive - SAP SE\HEP dashboard"

class _TP(HTMLParser):
    def __init__(self):
        super().__init__()
        self.in_hdr = self.in_row = self.in_cell = False
        self.headers = []; self.rows = []; self._row = []; self._cell = ''
    def handle_starttag(self, tag, attrs):
        d = dict(attrs)
        if tag == 'tr':
            if 'rowHeader' in d.get('class', ''):
                self.in_hdr = True
            else:
                self.in_row = True; self._row = []
        if tag in ('th', 'td'):
            self.in_cell = True; self._cell = ''
    def handle_endtag(self, tag):
        if tag in ('th', 'td'):
            self.in_cell = False
            t = ' '.join(self._cell.split())
            if self.in_hdr: self.headers.append(t)
            elif self.in_row: self._row.append(t)
        if tag == 'tr':
            if self.in_hdr: self.in_hdr = False
            elif self.in_row:
                if self._row and any(c for c in self._row): self.rows.append(self._row)
                self.in_row = False
    def handle_data(self, data):
        if self.in_cell: self._cell += data

with open(BASE + r"\Requirements\Jira_Validation.xls", 'rb') as f:
    tp = _TP(); tp.feed(f.read().decode('utf-8', errors='replace'))

val_map = {}
for r in tp.rows:
    if not r or r[0] != 'DISCOVALIDATION': continue
    key     = r[1]  if len(r) > 1  else ''
    status  = r[4]  if len(r) > 4  else ''
    resol   = r[6]  if len(r) > 6  else ''
    version = r[15] if len(r) > 15 else ''
    desc    = r[28] if len(r) > 28 else ''
    m = re.search(r'UseCase/view/(UC\d+)', desc)
    ucr_id = m.group(1) if m else ''
    if ucr_id:
        val_map[ucr_id] = {"key": key, "status": status, "resol": resol, "version": version}

print(f"Parsed {len(val_map)} validation tickets with UCR IDs")

html_path = BASE + r"\Missions\index.html"
with open(html_path, encoding='utf-8') as f:
    html = f.read()

replacement = f"// <<VAL_MAP_BEGIN>>\nvar VAL_MAP={json.dumps(val_map, ensure_ascii=False)};\n// <<VAL_MAP_END>>"
html = re.sub(r'// <<VAL_MAP_BEGIN>>.*?// <<VAL_MAP_END>>', replacement, html, flags=re.DOTALL)

with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html)
print(f"Written {len(html):,} bytes -> {html_path}")
print(f"VAL_MAP keys: {sorted(val_map.keys())}")
