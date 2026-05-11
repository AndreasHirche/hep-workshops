#!/usr/bin/env python3
"""
Build static public HTML from hep-dashboard-DEV.html + Excel data.

Run:   python build_public.py          # generate docs/index.html
       python build_public.py --push   # generate + git commit + git push
"""
import os, sys, re, shutil, subprocess
from datetime import datetime

sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

_DIR = os.path.dirname(os.path.abspath(__file__))
DOCS = os.path.join(_DIR, 'docs')
OUT  = os.path.join(DOCS, 'index.html')

sys.path.insert(0, _DIR)
from serve import inject

TEMPLATE = os.path.join(_DIR, 'hep-dashboard-DEV.html')

def build():
    print('Reading template ...')
    with open(TEMPLATE, 'r', encoding='utf-8') as f:
        html = f.read()

    print('Injecting Excel data ...')
    html = inject(html)

    # ── participant mode by default ───────────────────────────────────────────
    # add part-view class to body
    html = re.sub(r'<body([^>]*)>', lambda m: f'<body{m.group(1)} class="part-view"'
                  if 'class=' not in m.group(1)
                  else re.sub(r'class="([^"]*)"', r'class="part-view \1"', m.group(0)), html, count=1)

    # Globe: start with participant-view texture so no flash before updateGlobeTheme() runs
    html = html.replace(
        "  .globeImageUrl('https://unpkg.com/three-globe/example/img/earth-night.jpg')\n",
        "  .globeImageUrl('https://unpkg.com/three-globe/example/img/earth-blue-marble.jpg')\n")
    html = html.replace(
        "  .backgroundColor('#0f172a')\n  .showAtmosphere(true)\n  .atmosphereColor('#2563eb')\n",
        "  .backgroundColor('#2e4a62')\n  .showAtmosphere(true)\n  .atmosphereColor('#4fc3f7')\n")

    # update page title
    html = re.sub(r'<title>.*?</title>', '<title>SAP HEP Partner Workshops 2026</title>', html)

    os.makedirs(DOCS, exist_ok=True)
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'Written:  {OUT}')

    # Copy agenda pages so they're served via GitHub Pages
    AGENDAS = [
        ('Agenda_AI_PaloAlto_2026', 'agenda.html'),
        ('Agenda_IS_Bellevue_2026', 'agenda.html'),
        ('Agenda_AI_Helsinki_2026', 'Agenda_AI_Helsinki_2026.html'),
    ]
    for folder, fname in AGENDAS:
        src = os.path.join(_DIR, folder, fname)
        dst_dir = os.path.join(DOCS, folder)
        dst = os.path.join(dst_dir, fname)
        if os.path.exists(src):
            os.makedirs(dst_dir, exist_ok=True)
            shutil.copy2(src, dst)
            print(f'Copied:   {folder}/{fname}')
        else:
            print(f'Missing:  {folder}/{fname} (skipped)')

    open(os.path.join(DOCS, '.nojekyll'), 'w').close()


if __name__ == '__main__':
    build()

    if '--push' in sys.argv:
        cmds = [
            ['git', '-C', _DIR, 'add', 'docs/'],
            ['git', '-C', _DIR, 'commit', '-m', f'update: {datetime.now().strftime("%Y-%m-%d %H:%M")}'],
            ['git', '-C', _DIR, 'push'],
        ]
        for cmd in cmds:
            r = subprocess.run(cmd, capture_output=True, text=True)
            out = r.stdout.strip() or r.stderr.strip() or 'ok'
            print(' '.join(cmd[2:4]), ':', out[:80])
        print('Published to GitHub Pages.')
    else:
        print()
        print('To publish:  python build_public.py --push')
