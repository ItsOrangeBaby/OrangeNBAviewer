#!/usr/bin/env python3
"""NBA eight-team stats viewer. Standard-library only; no internet or pip needed.

Place this file beside NBA_4_Games_8_Teams_Last_6.xlsx, then run:
    python nba_stats_viewer.py

Optionally:
    python nba_stats_viewer.py --file /path/to/workbook.xlsx --port 8765
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import threading
import webbrowser
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from xml.etree import ElementTree as ET

DEFAULT_FILENAME = 'NBA_4_Games_8_Teams_Last_6.xlsx'
TEAMS = [
    ('BKN', 'Brooklyn Nets'),
    ('BOS', 'Boston Celtics'),
    ('DET', 'Detroit Pistons'),
    ('NYK', 'New York Knicks'),
    ('SAC', 'Sacramento Kings'),
    ('SAS', 'San Antonio Spurs'),
    ('POR', 'Portland Trail Blazers'),
    ('LAL', 'Los Angeles Lakers'),
]
MATCHUPS = [('BKN', 'BOS'), ('DET', 'NYK'), ('SAC', 'SAS'), ('POR', 'LAL')]
MAIN_NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
DOC_REL_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
PKG_REL_NS = '{http://schemas.openxmlformats.org/package/2006/relationships}'


def column_index(cell_ref: str) -> int:
    letters = re.match(r'[A-Z]+', cell_ref)
    if not letters:
        raise ValueError('Invalid Excel cell reference: ' + cell_ref)
    index = 0
    for char in letters.group(0):
        index = index * 26 + ord(char) - ord('A') + 1
    return index - 1


def excel_rows(archive: zipfile.ZipFile, sheet_path: str, shared: list[str]) -> list[list]:
    root = ET.fromstring(archive.read(sheet_path))
    results = []
    for row in root.findall(f'.//{MAIN_NS}sheetData/{MAIN_NS}row'):
        cells = {}
        for cell in row.findall(f'{MAIN_NS}c'):
            key = column_index(cell.attrib.get('r', 'A1'))
            typ = cell.attrib.get('t', '')
            raw = cell.findtext(f'{MAIN_NS}v')
            if typ == 'inlineStr':
                value = ''.join(node.text or '' for node in cell.findall(f'.//{MAIN_NS}t'))
            elif raw is None:
                value = None
            elif typ == 's':
                value = shared[int(raw)]
            elif typ in ('str', 'e'):
                value = raw
            elif typ == 'b':
                value = raw == '1'
            else:
                try:
                    numeric = float(raw)
                    value = int(numeric) if numeric.is_integer() else numeric
                except ValueError:
                    value = raw
            cells[key] = value
        if cells:
            results.append([cells.get(i) for i in range(max(cells) + 1)])
    return results


def load_workbook(path: Path) -> dict:
    if not path.is_file():
        raise FileNotFoundError(f'Workbook not found: {path}\nPut the XLSX beside this script, or pass --file PATH.')

    with zipfile.ZipFile(path) as archive:
        shared = []
        if 'xl/sharedStrings.xml' in archive.namelist():
            ss_root = ET.fromstring(archive.read('xl/sharedStrings.xml'))
            shared = [
                ''.join(t.text or '' for t in si.findall(f'.//{MAIN_NS}t'))
                for si in ss_root.findall(f'{MAIN_NS}si')
            ]
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        relroot = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        relmap = {e.attrib['Id']: e.attrib['Target'] for e in relroot.findall(f'{PKG_REL_NS}Relationship')}
        sheets = {}
        for elem in workbook.findall(f'.//{MAIN_NS}sheets/{MAIN_NS}sheet'):
            target = relmap[elem.attrib[DOC_REL_NS + 'id']]
            if target.startswith('/'):
                filename = target.lstrip('/')
            elif target.startswith('xl/'):
                filename = target
            else:
                filename = 'xl/' + target
            sheets[elem.attrib['name']] = excel_rows(archive, filename, shared)

    for sheet_name in ('Team Stats', 'Player Stats'):
        if sheet_name not in sheets:
            raise ValueError(f'Missing worksheet: {sheet_name}')

    expected_width = {'Team Stats': 25, 'Player Stats': 56}
    data = {}
    for sheet_name, width in expected_width.items():
        rows = sheets[sheet_name]
        if len(rows) < 3:
            raise ValueError(f'{sheet_name}: expected two header rows and data')
        headers = (rows[1] + [None] * width)[:width]
        if len([h for h in headers if h is not None]) != width:
            raise ValueError(f'{sheet_name}: not all {width} column headers found')
        first_stats_col = 1 if sheet_name == 'Team Stats' else 2
        boxes = []
        for start in range(first_stats_col, width, 6):
            box_number = len(boxes) + 1
            boxes.append({
                'number': box_number,
                'name': f'{"Team Box" if sheet_name == "Team Stats" else "Player Stats"} #{box_number}',
                'fields': headers[start:start + 6]
            })
        entries = []
        for row in rows[2:]:
            row = (row + [None] * width)[:width]
            if row[0] is None:
                continue
            entry = {'name': row[0]}
            if sheet_name == 'Player Stats':
                entry['team'] = row[1]
            entry['groups'] = [row[start:start + 6] for start in range(first_stats_col, width, 6)]
            entries.append(entry)
        data[sheet_name] = {'boxes': boxes, 'entries': entries}

    team_names = [team for _, team in TEAMS]
    actual_teams = [t['name'] for t in data['Team Stats']['entries']]
    if set(actual_teams) != set(team_names):
        raise ValueError(f'Unexpected team list in workbook: {actual_teams}')
    known_codes = {code for code, _ in TEAMS}
    unknown = {p['team'] for p in data['Player Stats']['entries']} - known_codes
    if unknown:
        raise ValueError(f'Unknown player team codes: {sorted(unknown)}')
    if len(data['Team Stats']['boxes']) != 4 or len(data['Player Stats']['boxes']) != 9:
        raise ValueError('Expected four team groups and nine player groups')

    return {
        'filename': path.name,
        'teamData': data['Team Stats'],
        'playerData': data['Player Stats'],
        'teams': [{'code': code, 'name': name} for code, name in TEAMS],
        'matchups': [{'home': first, 'away': second} for first, second in MATCHUPS],
    }


HTML = r'''<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>NBA Stats Explorer</title>
<style>
:root {--bg:#0c1220;--surface:#141e31;--surface2:#1b2940;--border:#304059;--ink:#eff4fc;--muted:#a3b1c5;--accent:#63d2b2;--accent2:#83b8ff;--radius:15px}
*{box-sizing:border-box}body{margin:0;background:radial-gradient(circle at 15% 0%,#172941 0%,var(--bg) 43%);color:var(--ink);font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;line-height:1.45}
main{max-width:1240px;margin:auto;padding:30px 24px 70px}
header{display:flex;justify-content:space-between;align-items:flex-start;gap:20px;flex-wrap:wrap;margin-bottom:24px}
.eyebrow{text-transform:uppercase;color:var(--accent);letter-spacing:.16em;font-size:.7rem;font-weight:750;margin-bottom:8px}
h1{font-size:clamp(1.8rem,3.3vw,2.8rem);line-height:1.1;letter-spacing:-.035em;margin:0 0 8px}h2{font-size:1.25rem;letter-spacing:-.018em;margin:0}p{margin:0;color:var(--muted)}
.source{font-size:.76rem;max-width:280px;text-align:right;color:var(--muted);overflow-wrap:anywhere;padding-top:4px}
.topnav{display:flex;gap:8px;padding:6px;border-radius:13px;background:#101a2a;border:1px solid var(--border);width:max-content;max-width:100%;margin-bottom:20px}
button,.field{font:inherit}button{cursor:pointer}button:focus-visible,select:focus-visible,input:focus-visible{outline:2px solid var(--accent);outline-offset:3px}
.tab{background:transparent;color:var(--muted);border:0;border-radius:9px;padding:10px 18px;font-weight:700}.tab.active{background:#293b54;color:white;box-shadow:0 1px 5px #0003}
.controls{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:18px;display:flex;flex-wrap:wrap;gap:14px;align-items:end;margin-bottom:20px}
.control{min-width:190px;flex:1;display:grid;gap:8px}.control.wide{flex:2;min-width:250px}label{font-size:.76rem;font-weight:700;color:#b9c8dc;letter-spacing:.025em}
.field{width:100%;border:1px solid #40536c;color:var(--ink);background:#0f1828;border-radius:10px;padding:11px 12px;min-height:45px}
option{background:#101a2b;color:#fff}.section-header{display:flex;align-items:baseline;justify-content:space-between;gap:14px;flex-wrap:wrap;margin:25px 0 13px}
.meta{font-size:.81rem;color:var(--muted)}.box-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}.stat-box{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);overflow:hidden;min-width:0}
.box-header{background:#1b2a42;padding:14px 17px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;gap:8px;align-items:baseline}
.box-header h3{margin:0;font-size:.96rem;letter-spacing:.01em}.box-header small{color:var(--muted);font-size:.76rem}
.metric-row{display:grid;grid-template-columns:minmax(0,1fr) minmax(70px,.65fr) minmax(70px,.65fr);gap:8px;padding:11px 16px;border-bottom:1px solid #ffffff0b;align-items:center;font-size:.85rem}
.metric-row:last-child{border-bottom:0}.metric-label{color:#b4c3d8;overflow-wrap:anywhere}.metric-val{text-align:right;font-variant-numeric:tabular-nums;font-weight:730;font-size:1rem}.metric-val.left{color:#d8f6ec}.metric-val.right{color:#c4deff}
.column-names{display:grid;grid-template-columns:minmax(0,1fr) minmax(70px,.65fr) minmax(70px,.65fr);gap:8px;padding:12px 16px 10px;border-bottom:1px solid var(--border);font-size:.68rem;letter-spacing:.025em;text-transform:uppercase;color:#9eb3cb}
.column-names span:not(:first-child){text-align:right}.single .metric-row{grid-template-columns:1fr 1fr}.single .metric-val{color:var(--ink)}.player .metric-row{grid-template-columns:1fr 1fr}
.matchup-row{display:flex;flex-wrap:wrap;gap:9px}.matchup-button{background:transparent;color:#b6c8df;border:1px solid var(--border);border-radius:10px;padding:11px 13px;font-size:.82rem;font-weight:700}
.matchup-button.active{background:#1e3847;color:#b3ffe8;border-color:#3a907c}
.footer-note{font-size:.77rem;color:#a3b1c5;margin-top:19px}.hidden{display:none!important}.empty{padding:42px 15px;text-align:center;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);color:var(--muted)}
.stats-count{border:1px solid #3c5c71;border-radius:100px;padding:6px 11px;background:#1a2b3b;font-size:.73rem;color:#ccdfef;white-space:nowrap}
@media(max-width:850px){.box-grid{grid-template-columns:1fr}header{margin-bottom:19px}.source{text-align:left}}
@media(max-width:510px){main{padding:20px 12px 55px}.metric-row,.column-names{padding-left:11px;padding-right:11px}.control{min-width:100%}.controls{padding:13px}.metric-val{font-size:.9rem}}
@media print{body{background:white;color:black}.topnav,.controls,.matchup-row,.source,.footer-note{display:none!important}.stat-box{break-inside:avoid;border-color:#aaa;background:white}.box-header{background:#eee;color:#111}.metric-label,.metric-val,.box-header small,.meta,p{color:#222!important}main{max-width:none;padding:0}}
</style></head><body><main>
<header><div><div class="eyebrow">NBA / Last 6 games</div><h1>Stats Explorer</h1><p>Four matchups · 8 teams · six stats per box</p></div><div class="source" id="source"></div></header>
<nav class="topnav" aria-label="Choose stats page"><button class="tab active" id="teamTab">Team Stats</button><button class="tab" id="playerTab">Player Stats</button></nav>
<section id="teamPage">
 <div class="section-header"><h2>Tonight's matchups</h2><span class="stats-count">4 games · 8 teams</span></div>
 <div class="matchup-row" id="matchupButtons"></div>
 <div class="controls" style="margin-top:16px">
  <div class="control"><label for="teamMode">Display</label><select id="teamMode" class="field"><option value="compare">Compare both teams</option><option value="single">Single team</option></select></div>
  <div class="control" id="teamSelectWrap"><label for="teamSelect">Team</label><select id="teamSelect" class="field"></select></div>
  <div class="control"><label for="teamGroup">Stats box</label><select id="teamGroup" class="field"></select></div>
 </div>
 <div class="section-header"><h2 id="teamHeading">Brooklyn vs Boston</h2><span class="meta" id="teamSubheading"></span></div>
 <div class="box-grid" id="teamBoxes"></div>
</section>
<section id="playerPage" class="hidden">
 <div class="controls">
  <div class="control"><label for="playerTeam">Team</label><select class="field" id="playerTeam"></select></div>
  <div class="control wide"><label for="playerSearch">Find a player</label><input class="field" id="playerSearch" placeholder="Type a name (e.g. Brunson)" autocomplete="off"></div>
  <div class="control wide"><label for="playerSelect">Choose player</label><select class="field" id="playerSelect"></select></div>
  <div class="control"><label for="playerGroup">Stats box</label><select class="field" id="playerGroup"></select></div>
 </div>
 <div class="section-header"><h2 id="playerHeading">Player Stats</h2><span class="meta" id="playerSubheading"></span></div>
 <div class="box-grid player" id="playerBoxes"></div>
</section>
<p class="footer-note">Values come directly from the workbook. An em dash (—) means the source cell was blank; it is not zero. The viewer does not download or modify statistics.</p>
</main>
<script>
'use strict';
let DATA=null;let selectedMatchup=0;const byId=id=>document.getElementById(id);
const names={};const number=v => (v===null || v===undefined || v==='')?'—':typeof v==='number'?String(v):String(v);
function el(tag,cls,txt){const e=document.createElement(tag);if(cls)e.className=cls;if(txt!==undefined)e.textContent=txt;return e}
function setupOptions(select,values){select.replaceChildren();for(const [value,label] of values){const o=el('option','',label);o.value=String(value);select.append(o)}}
function groupOptions(boxes){return [['all','All boxes'],...boxes.map((b,i)=>[String(i),b.name+' · 6 stats'])]}
function boxCard(box,leftValues,rightValues,leftName,rightName,single){
 const card=el('article','stat-box'+(single?' single':''));const hd=el('div','box-header');hd.append(el('h3','',box.name),el('small','','6 stats'));card.append(hd);
 if(!single){const titles=el('div','column-names');titles.append(el('span','', 'Statistic'),el('span','',leftName),el('span','',rightName));card.append(titles)}
 box.fields.forEach((field,i)=>{const row=el('div','metric-row');row.append(el('span','metric-label',field),el('span','metric-val left',number(leftValues[i])));if(!single)row.append(el('span','metric-val right',number(rightValues[i])));card.append(row)});
 return card;
}
function filteredBoxes(boxes,selected){return boxes.map((box,i)=>({box,i})).filter(x=>selected==='all'||String(x.i)===selected)}
function init(){
 DATA=window.__NBA_DATA__;for(const t of DATA.teams)names[t.code]=t.name;
 byId('source').textContent='Source: '+DATA.filename+' · '+DATA.playerData.entries.length+' players';
 setupOptions(byId('teamSelect'),DATA.teams.map(t=>[t.code,t.name]));setupOptions(byId('playerTeam'),[['all','All 8 teams'],...DATA.teams.map(t=>[t.code,t.name])] );
 setupOptions(byId('teamGroup'),groupOptions(DATA.teamData.boxes));setupOptions(byId('playerGroup'),groupOptions(DATA.playerData.boxes));
 const buttons=byId('matchupButtons');DATA.matchups.forEach((m,i)=>{const b=el('button','matchup-button',m.home+'  vs  '+m.away);b.onclick=()=>{selectedMatchup=i;byId('teamSelect').value=m.home;renderMatchupButtons();renderTeam()};buttons.append(b)});
 byId('teamSelect').value=DATA.matchups[0].home;
 for(const id of ['teamMode','teamSelect','teamGroup'])byId(id).addEventListener('change',renderTeam);
 for(const id of ['playerTeam','playerSearch'])byId(id).addEventListener(id==='playerSearch'?'input':'change',()=>{fillPlayers();renderPlayer()});
 for(const id of ['playerSelect','playerGroup'])byId(id).addEventListener('change',renderPlayer);
 byId('teamTab').onclick=()=>switchTab('team');byId('playerTab').onclick=()=>switchTab('player');
 fillPlayers();renderMatchupButtons();renderTeam();renderPlayer();switchTab('team');
}
function switchTab(which){const onTeam=which==='team';byId('teamPage').classList.toggle('hidden',!onTeam);byId('playerPage').classList.toggle('hidden',onTeam);byId('teamTab').classList.toggle('active',onTeam);byId('playerTab').classList.toggle('active',!onTeam);document.title=(onTeam?'Team':'Player')+' Stats | NBA Stats Explorer'}
function renderMatchupButtons(){Array.from(byId('matchupButtons').children).forEach((b,i)=>b.classList.toggle('active',i===selectedMatchup))}
function renderTeam(){
 const single=byId('teamMode').value==='single';byId('teamSelectWrap').classList.toggle('hidden',!single);
 const m=DATA.matchups[selectedMatchup],leftCode=single?byId('teamSelect').value:m.home,rightCode=m.away;
 const rows=DATA.teamData.entries;const l=rows.find(t=>t.name===names[leftCode]);const r=rows.find(t=>t.name===names[rightCode]);
 byId('teamHeading').textContent=single?names[leftCode]:names[leftCode]+' vs '+names[rightCode];
 byId('teamSubheading').textContent=single?'24 stats · 4 boxes':'24 stats per team · 4 boxes';
 const container=byId('teamBoxes');container.replaceChildren();
 if(!l||(!single&&!r)){container.append(el('div','empty','Team not found in this workbook'));return}
 for(const {box,i} of filteredBoxes(DATA.teamData.boxes,byId('teamGroup').value)){
  container.append(boxCard(box,l.groups[i],single?null:r.groups[i],leftCode,rightCode,single));
 }
}
function fillPlayers(){
 const chosen=byId('playerSelect').value;const team=byId('playerTeam').value,q=byId('playerSearch').value.trim().toLowerCase();
 const players=DATA.playerData.entries.filter(p=>(team==='all'||team===p.team)&&(!q||p.name.toLowerCase().includes(q)));
 const options=players.map(p=>[p.team+'|'+p.name,p.name+' · '+p.team]);setupOptions(byId('playerSelect'),options);
 if(players.some(p=>p.team+'|'+p.name===chosen))byId('playerSelect').value=chosen;
 byId('playerSubheading').textContent=players.length+' matching players · 54 stats per player';
}
function renderPlayer(){
 const player=DATA.playerData.entries.find(p=>p.team+'|'+p.name===byId('playerSelect').value);const wrap=byId('playerBoxes');wrap.replaceChildren();
 if(!player){byId('playerHeading').textContent='No matching players';wrap.append(el('div','empty','Try a different name or team.'));return}
 byId('playerHeading').textContent=player.name+' · '+player.team;
 for(const {box,i} of filteredBoxes(DATA.playerData.boxes,byId('playerGroup').value)){wrap.append(boxCard(box,player.groups[i],null,player.team,null,true))}
}
window.addEventListener('DOMContentLoaded',init);
</script></body></html>'''


def handler_factory(data: dict):
    payload = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
    # Encode serialized data separately from HTML. Escaping prevents HTML-script injection.
    safe = payload.replace('<', '\\u003c').replace('>', '\\u003e').replace('&', '\\u0026')
    page = HTML.replace('</head>', '<script>window.__NBA_DATA__=' + safe + ';</script></head>')
    page_bytes = page.encode('utf-8')

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            if self.path in ('/', '/index.html'):
                content,ctype,status=page_bytes,'text/html; charset=utf-8',200
            elif self.path == '/api/data':
                content,ctype,status=payload.encode('utf-8'),'application/json; charset=utf-8',200
            elif self.path == '/favicon.ico':
                content,ctype,status=b'','image/x-icon',204
            else:
                content,ctype,status=b'Not Found','text/plain; charset=utf-8',404
            self.send_response(status)
            self.send_header('Content-Type',ctype)
            self.send_header('Content-Length',str(len(content)))
            self.send_header('Cache-Control','no-store')
            self.end_headers()
            self.wfile.write(content)
        def log_message(self, fmt, *args):
            if args and str(args[1]) not in ('200','204'):
                super().log_message(fmt,*args)

    return Handler


def main():
    parser=argparse.ArgumentParser(description='View NBA team/player six-stat boxes in your browser, offline.')
    parser.add_argument('--file',type=Path,default=Path(__file__).resolve().parent / DEFAULT_FILENAME,help='Path to Excel workbook')
    parser.add_argument('--port',type=int,default=8765,help='Local server port (default 8765)')
    parser.add_argument('--no-browser',action='store_true',help='Do not open a browser automatically')
    args=parser.parse_args()
    try:
        data=load_workbook(args.file.expanduser().resolve())
        server=ThreadingHTTPServer(('127.0.0.1',args.port),handler_factory(data))
    except (OSError,ValueError,FileNotFoundError,zipfile.BadZipFile,KeyError) as exc:
        print('Error: '+str(exc),file=sys.stderr)
        return 1
    url=f'http://127.0.0.1:{server.server_port}/'
    print(f'NBA Stats Explorer: {url}')
    print(f'Loaded {len(data["teamData"]["entries"])} teams, {len(data["playerData"]["entries"])} players.')
    print('Team Stats opens first. Press Ctrl+C here to close the viewer.')
    if not args.no_browser:
        threading.Timer(0.5,lambda:webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nViewer closed.')
    finally:
        server.server_close()
    return 0


if __name__=='__main__':
    raise SystemExit(main())
