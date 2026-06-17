#!/usr/bin/env python3
"""
Playbook Dashboard Updater
--------------------------
Reads all playbook-monthly-YYYY-MM.xlsx files from the /data folder,
combines them, and rebuilds index.html with full history.

Usage:
    python update_dashboard.py

Run this after dropping a new monthly Excel file into /data.
"""

import os
import re
import json
import sys
import pandas as pd
from pathlib import Path

# ── Paths ────────────────────────────────────────────────────────────────────
SCRIPT_DIR  = Path(__file__).parent
DATA_DIR    = SCRIPT_DIR / "data"
OUTPUT_FILE = SCRIPT_DIR / "index.html"

# ── Playbook name normalisation ───────────────────────────────────────────────
PLAYBOOK_MAP = {
    "dx_playbook":                        "DX Playbook",
    "dx_competencies_leadership_drivers": "DX Playbook",
    "healthcare_vertical_playbook":       "Healthcare Playbook",
    "legal_vertical_playbook":            "Legal Playbook",
    "salesforce_playbook":                "Salesforce Playbook",
    "public_sector_playbook":             "Public Sector Playbook",
    "accelerate_sales_playbook":          "Accelerate",
    "iq501":                              "IQ501",
    "gc_ip_sales_playbook":               "GC/IP Sales Playbook",
    "road_to_dx":                         "Road to DX",
}

def get_playbook(url):
    url = str(url)
    m = re.search(r'/playbooks/([^/]+)/', url)
    if m:
        key = m.group(1).lower()
        return PLAYBOOK_MAP.get(key, key.replace('_', ' ').title())
    return "Accelerate"

def get_page(url):
    url = str(url)
    parts = url.rstrip('/').split('/')
    last = parts[-1]
    if last in ('', 'index.html'):
        return 'Home'
    return last.replace('.html', '').replace('_', ' ').replace('-', ' ').title()

def load_excel(path: Path, month_label: str) -> pd.DataFrame:
    df = pd.read_excel(path)
    df = df.drop(columns=[c for c in ['Uid','Email','Employee Id','Market','Branch'] if c in df.columns])
    df['Playbook'] = df['Url'].apply(get_playbook)
    df['Page']     = df['Url'].apply(get_page)
    df['Date']     = pd.to_datetime(df['Date']).dt.strftime('%Y-%m-%d')
    df['Month']    = month_label
    df = df.rename(columns={
        'First Name':      'FirstName',
        'Last Name':       'LastName',
        'Employee/Dealer': 'Type',
    })
    return df[['FirstName','LastName','Region','Type','Date','Month','Playbook','Page']]

# ── Collect all monthly files ─────────────────────────────────────────────────
pattern = re.compile(r'^playbook-monthly-(\d{4}-\d{2})\.xlsx$')
files   = sorted([
    (m.group(1), p)
    for p in DATA_DIR.glob('*.xlsx')
    if (m := pattern.match(p.name))
])

if not files:
    print(f"No files matching playbook-monthly-YYYY-MM.xlsx found in {DATA_DIR}")
    sys.exit(1)

print(f"Found {len(files)} monthly file(s):")
for label, path in files:
    print(f"  {label}  →  {path.name}")

frames = []
for label, path in files:
    df = load_excel(path, label)
    frames.append(df)
    print(f"  Loaded {len(df):,} rows from {label}")

combined = pd.concat(frames, ignore_index=True)
combined['Region'] = combined['Region'].where(combined['Region'].notna(), None)
records = json.loads(combined.to_json(orient='records'))

total_rows = len(records)
months     = sorted(set(r['Month'] for r in records))
print(f"\nTotal rows combined: {total_rows:,}")
print(f"Months covered: {', '.join(months)}")

# ── HTML template ─────────────────────────────────────────────────────────────
html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Playbook Traffic Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2"></script>
<style>
  :root {{
    --bg:#0f1117; --surface:#1a1d27; --surface2:#22263a; --border:#2e3350;
    --accent:#4f8ef7; --accent2:#7c5cfc; --accent3:#f7c94f;
    --text:#e8ecf4; --muted:#7b82a0; --green:#3ecf8e; --red:#f76f6f;
    --teal:#2dd4bf; --teal-subtle:#2dd4bf18;
    --font:'Segoe UI',system-ui,sans-serif;
    --pill-emp-bg:#1a2a4a; --pill-emp-color:#4f8ef7;
    --pill-dlr-bg:#2a1a3a; --pill-dlr-color:#cf5cf7;
  }}
  body.light-mode {{
    --bg:#f4f6fb; --surface:#ffffff; --surface2:#eef1f7; --border:#d0d7e8;
    --accent:#2563eb; --accent2:#6d28d9; --accent3:#d97706;
    --text:#1a1d27; --muted:#475569; --green:#059669; --red:#dc2626;
    --teal:#0f766e; --teal-subtle:#0f766e18;
    --pill-emp-bg:#dbeafe; --pill-emp-color:#1d4ed8;
    --pill-dlr-bg:#f3e8ff; --pill-dlr-color:#7c3aed;
  }}
  body.light-mode select, body.light-mode input[type=date]{{color-scheme:light;}}
  *{{box-sizing:border-box;margin:0;padding:0;}}
  body{{background:var(--bg);color:var(--text);font-family:var(--font);min-height:100vh;transition:background .2s,color .2s;}}

  .header{{padding:20px 28px 16px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;}}
  .header-left{{display:flex;align-items:center;gap:16px;}}
  .header h1{{font-size:18px;font-weight:700;letter-spacing:.3px;}}
  .header h1 span{{color:var(--muted);font-weight:400;}}
  .hamburger{{position:relative;}}
  .hamburger-btn{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:6px 10px;font-size:16px;cursor:pointer;transition:all .15s;line-height:1;}}
  .hamburger-btn:hover,.hamburger-btn.open{{border-color:var(--accent);color:var(--text);}}
  .hamburger-menu{{position:absolute;top:calc(100% + 6px);left:0;background:var(--surface);border:1px solid var(--border);border-radius:8px;min-width:240px;box-shadow:0 4px 24px rgba(0,0,0,0.28);display:none;z-index:200;overflow:hidden;}}
  .hamburger-menu.open{{display:block;}}
  .hamburger-section-label{{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--muted);padding:10px 14px 4px;}}
  .hamburger-item{{display:flex;align-items:center;gap:8px;padding:10px 14px;font-size:13px;color:var(--text);text-decoration:none;transition:background .1s;}}
  .hamburger-item:hover{{background:var(--surface2);}}
  .badges{{display:flex;gap:8px;flex-wrap:wrap;}}
  .badge{{background:var(--surface2);border:1px solid var(--border);border-radius:20px;padding:4px 12px;font-size:12px;color:var(--muted);}}

  .filters{{padding:14px 28px;display:flex;gap:10px;flex-wrap:wrap;align-items:center;border-bottom:1px solid var(--border);background:var(--surface);}}
  .filter-label{{font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.6px;margin-right:4px;}}
  select,input[type=text],input[type=date]{{background:var(--surface2);border:1px solid var(--border);color:var(--text);border-radius:6px;padding:6px 10px;font-size:13px;cursor:pointer;outline:none;color-scheme:dark;}}
  select:focus,input:focus{{border-color:var(--accent);}}
  .btn-reset{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:6px 14px;font-size:12px;cursor:pointer;transition:border-color .15s,color .15s;}}
  .btn-reset:hover{{border-color:var(--accent);color:var(--text);}}
  .btn-tlg{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:6px 14px;font-size:12px;cursor:pointer;transition:all .15s;}}
  .btn-tlg:hover{{border-color:var(--red);color:var(--red);}}
  .btn-tlg.active{{background:#f76f6f22;border-color:var(--red);color:var(--red);}}
  .btn-preset{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:4px 10px;font-size:11px;cursor:pointer;transition:all .15s;}}
  .btn-preset:hover{{border-color:var(--accent);color:var(--text);}}
  .btn-preset.active{{background:#4f8ef722;border-color:var(--accent);color:var(--accent);font-weight:600;}}
  .btn-theme{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:5px 12px;font-size:12px;cursor:pointer;transition:all .15s;}}
  .btn-theme:hover{{border-color:var(--accent);color:var(--text);}}
  .result-count{{margin-left:auto;font-size:12px;color:var(--muted);}}

  .stats{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;padding:20px 28px;}}
  .stat{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:16px 18px;}}
  .stat-label{{font-size:11px;text-transform:uppercase;letter-spacing:.7px;color:var(--muted);margin-bottom:6px;}}
  .stat-value{{font-size:28px;font-weight:700;line-height:1;}}
  .stat-value.blue{{color:var(--accent);}} .stat-value.purple{{color:var(--accent2);}}
  .stat-value.yellow{{color:var(--accent3);}} .stat-value.green{{color:var(--green);}}
  .stat-sub{{font-size:11px;color:var(--muted);margin-top:4px;}}

  .charts-top{{display:grid;grid-template-columns:2fr 1fr;gap:16px;padding:0 28px 16px;}}
  .charts-bottom{{padding:0 28px 20px;}}
  .charts-pages{{display:grid;grid-template-columns:1fr 1fr;gap:16px;padding:0 28px 16px;}}
  @media(max-width:860px){{.charts-top,.charts-pages{{grid-template-columns:1fr;}}}}
  .chart-card{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:18px;}}
  .chart-title{{font-size:13px;font-weight:600;margin-bottom:14px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;}}
  .chart-wrap{{position:relative;height:240px;}}
  .chart-wrap-tall{{position:relative;height:180px;}}
  .chart-wrap-pages{{position:relative;height:320px;}}

  .table-section{{padding:0 28px 32px;}}
  .table-header{{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;flex-wrap:wrap;gap:8px;}}
  .table-title{{font-size:13px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;}}
  .table-outer{{border-radius:10px;border:1px solid var(--border);overflow:hidden;}}
  table{{width:100%;border-collapse:collapse;font-size:13px;}}
  thead tr{{background:var(--surface2);}}
  thead th{{padding:10px 14px;text-align:left;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);cursor:pointer;user-select:none;white-space:nowrap;}}
  thead th:hover{{color:var(--text);}} thead th.sorted{{color:var(--accent);}}
  tbody tr{{border-top:1px solid var(--border);transition:background .1s;}}
  tbody tr:hover{{background:var(--surface2);}}
  tbody td{{padding:9px 14px;vertical-align:middle;}}
  .pill{{display:inline-block;padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600;white-space:nowrap;}}
.no-data{{text-align:center;color:var(--muted);padding:40px;font-size:13px;}}
  .section-hint{{font-size:11px;color:var(--muted);margin-bottom:14px;margin-top:-8px;opacity:0.7;}}
  .info-btn{{display:inline-flex;align-items:center;justify-content:center;width:15px;height:15px;border-radius:50%;background:var(--surface2);border:1px solid var(--border);color:var(--muted);font-size:9px;font-weight:700;cursor:pointer;margin-left:5px;vertical-align:middle;flex-shrink:0;line-height:1;transition:border-color .15s,color .15s;}}
  .info-btn:hover{{border-color:var(--accent);color:var(--accent);}}
  .info-popover{{position:fixed;z-index:9999;background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:12px 14px;font-size:12px;color:var(--text);line-height:1.6;max-width:260px;box-shadow:0 4px 24px rgba(0,0,0,0.5);display:none;}}
  .info-popover.visible{{display:block;}}
  .drilldown-wrap{{display:flex;border:1px solid var(--border);border-radius:10px;overflow:hidden;}}
  .drilldown-left{{width:260px;flex-shrink:0;overflow-y:auto;max-height:820px;border-right:1px solid var(--border);}}
  .drilldown-person{{display:flex;align-items:center;gap:10px;padding:10px 14px;cursor:pointer;border-bottom:1px solid var(--border);transition:background .1s;}}
  .drilldown-person:last-child{{border-bottom:none;}}
  .drilldown-person:hover{{background:var(--surface2);}}
  .drilldown-person.active{{background:#4f8ef711;border-left:3px solid var(--accent);padding-left:11px;}}
  .drilldown-name{{flex:1;font-size:13px;}}
  .drilldown-count{{font-size:11px;font-weight:700;color:var(--teal);background:var(--teal-subtle);border-radius:10px;padding:2px 8px;}}
  .recency-dot{{width:8px;height:8px;border-radius:50%;flex-shrink:0;display:inline-block;}}
  .recency-legend{{display:flex;gap:16px;align-items:center;padding:8px 0 10px;font-size:11px;color:var(--muted);flex-wrap:wrap;}}
  .recency-legend-item{{display:flex;align-items:center;gap:5px;}}
  body.light-mode .pill-pb{{filter:brightness(0.55);}}
  body.light-mode .chart-title,body.light-mode .table-title{{color:var(--text);}}
  body.light-mode .section-hint{{opacity:1;color:var(--muted);}}
  .drilldown-right{{flex:1;overflow-y:auto;max-height:820px;padding:16px 18px;}}
  .drilldown-right-header{{margin-bottom:12px;padding-bottom:12px;border-bottom:1px solid var(--border);font-size:13px;}}
  .drilldown-right table thead th{{padding:8px 12px;text-align:left;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);}}
  .drilldown-right table tbody tr{{border-top:1px solid var(--border);}}
  .drilldown-right table tbody td{{padding:8px 12px;font-size:13px;}}
  @media(max-width:680px){{.drilldown-wrap{{flex-direction:column;}}.drilldown-left{{width:100%;max-height:200px;border-right:none;border-bottom:1px solid var(--border);}}}}
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <div class="hamburger" id="hamburger">
      <button class="hamburger-btn" id="hamburger-btn" onclick="toggleHamburger()" aria-label="Menu">&#9776;</button>
      <div class="hamburger-menu" id="hamburger-menu">
        <div class="hamburger-section-label">Certifications</div>
        <a href="cert-healthcare.html" class="hamburger-item">&#127973; Healthcare Certification Dashboard</a>
      </div>
    </div>
    <h1>Playbook Traffic Dashboard</h1>
  </div>
  <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
    <span class="badge" id="badge-asof"></span>
    <button class="btn-theme" id="btn-theme" onclick="toggleTheme()">☀ Light</button>
  </div>
</div>

<div class="filters">
  <span class="filter-label">Filter</span>
  <button class="btn-preset" id="btn-7d"  onclick="setRange(7,this)">7D</button>
  <button class="btn-preset active" id="btn-30d" onclick="setRange(30,this)">30D</button>
  <button class="btn-preset" id="btn-90d" onclick="setRange(90,this)">90D</button>
  <button class="btn-preset" id="btn-all" onclick="setRange(0,this)">All</button>
  <span class="filter-label" style="margin-right:2px">From</span>
  <input type="date" id="f-date-from">
  <span class="filter-label" style="margin:0 2px">To</span>
  <input type="date" id="f-date-to">
  <select id="f-playbook"><option value="">All Playbooks</option></select>
  <select id="f-region"><option value="">All Regions</option></select>
  <select id="f-type"><option value="">Employee &amp; Dealer</option><option value="Employee">Employee</option><option value="Dealer">Dealer</option></select>
  <button class="btn-reset" onclick="resetFilters()">Reset</button>
  <button class="btn-tlg" id="btn-tlg" onclick="toggleTLG()">Hide TLG</button><span class="info-btn" onclick="showInfo(event,'hide-tlg')">?</span>
  <button class="btn-tlg" id="btn-vertical" onclick="toggleVertical()">Vertical Markets</button><span class="info-btn" onclick="showInfo(event,'vertical-filter')">?</span>
  <span class="result-count" id="result-count"></span>
</div>

<div class="stats" id="stats-row"></div>

<div class="charts-top">
  <div class="chart-card">
    <div class="chart-title">Page Views by Playbook<span class="info-btn" onclick="showInfo(event,'chart-playbook')">?</span></div>
    <div class="section-hint">Hover over a bar to see the view count</div>
    <div class="chart-wrap"><canvas id="barChart"></canvas></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">Views by Region<span class="info-btn" onclick="showInfo(event,'chart-region')">?</span></div>
    <div class="section-hint">Hover over a segment to see the region breakdown</div>
    <div class="chart-wrap"><canvas id="pieChart"></canvas></div>
  </div>
</div>

<div class="charts-bottom">
  <div class="chart-card">
    <div class="chart-title">Monthly Trend — Views Over Time<span class="info-btn" onclick="showInfo(event,'chart-trend')">?</span></div>
    <div class="section-hint">Hover to see monthly totals by playbook — shows top 5 playbooks by volume</div>
    <div class="chart-wrap-tall"><canvas id="trendChart"></canvas></div>
  </div>
</div>

<div class="charts-bottom">
  <div class="chart-card">
    <div class="chart-title" id="pages-chart-title">Top Pages</div>
    <div class="section-hint">Hover over a bar to see views, unique visitors, and avg visits per person</div>
    <div class="chart-wrap-pages"><canvas id="pagesChart"></canvas></div>
  </div>
</div>


<div class="table-section" id="drilldown-section">
  <div class="table-header">
    <span class="table-title">Who's Active<span class="info-btn" onclick="showInfo(event,'whos-active')">?</span></span>
    <input type="text" id="drilldown-search" oninput="filterPersonList()" placeholder="Search name..." style="font-size:12px;padding:4px 10px;width:180px;background:var(--surface2);border:1px solid var(--border);color:var(--text);border-radius:6px;outline:none;">
  </div>
  <div class="recency-legend">
    <span style="font-weight:600;text-transform:uppercase;letter-spacing:.5px;">Last visit:</span>
    <span class="recency-legend-item"><span class="recency-dot" style="background:var(--green)"></span> Within 30 days</span>
    <span class="recency-legend-item"><span class="recency-dot" style="background:var(--accent3)"></span> 31–60 days</span>
    <span class="recency-legend-item"><span class="recency-dot" style="background:var(--red)"></span> 60+ days</span>
  </div>
  <div class="drilldown-wrap">
    <div class="drilldown-left" id="drilldown-left"></div>
    <div class="drilldown-right" id="drilldown-right"><div class="no-data">Select a person to see their activity</div></div>
  </div>
</div>

<script>
if (typeof ChartDataLabels !== 'undefined') Chart.register(ChartDataLabels);
const RAW = {json.dumps(records)};

const PLAYBOOK_COLORS = {{
  "Salesforce Playbook":    "#4f8ef7",
  "Healthcare Playbook":    "#3ecf8e",
  "Public Sector Playbook": "#f7c94f",
  "Accelerate":             "#f76f6f",
  "DX Playbook":            "#7c5cfc",
  "Legal Playbook":         "#f7944f",
  "IQ501":                  "#5cf0f7",
  "GC/IP Sales Playbook":   "#cf5cf7",
  "Road to DX":             "#2dd4bf",
}};
function pbColor(pb){{ return PLAYBOOK_COLORS[pb] || "#7b82a0"; }}
function recencyColor(dateStr){{
  if(!dateStr) return cv('--muted');
  const days = Math.floor((Date.now() - new Date(dateStr).getTime()) / 86400000);
  if(days <= 30) return cv('--green');
  if(days <= 60) return cv('--accent3');
  return cv('--red');
}}

const allMonths   = [...new Set(RAW.map(r=>r.Month))].sort();
const allPlaybooks= [...new Set([...Object.keys(PLAYBOOK_COLORS), ...RAW.map(r=>r.Playbook)])].sort();
const allRegions  = [...new Set(RAW.map(r=>r.Region).filter(Boolean))].sort();

function sel(id){{ return document.getElementById(id); }}
function cv(v){{ return getComputedStyle(document.body).getPropertyValue(v).trim(); }}

// Theme init
(function(){{
  if(localStorage.getItem('pb-theme')==='light') document.body.classList.add('light-mode');
  sel('btn-theme').textContent = document.body.classList.contains('light-mode') ? '🌙 Dark' : '☀ Light';
}})();
function toggleTheme(){{
  const light = document.body.classList.toggle('light-mode');
  localStorage.setItem('pb-theme', light ? 'light' : 'dark');
  sel('btn-theme').textContent = light ? '🌙 Dark' : '☀ Light';
  applyFilters();
}}
allPlaybooks.forEach(p => sel('f-playbook').innerHTML += `<option value="${{p}}">${{p}}</option>`);
allRegions.forEach(r => sel('f-region').innerHTML  += `<option value="${{r}}">${{r}}</option>`);

function daysAgo(n){{ const d=new Date(); d.setDate(d.getDate()-n); return d.toISOString().slice(0,10); }}
function today(){{ return new Date().toISOString().slice(0,10); }}
sel('f-date-from').value = daysAgo(30);
sel('f-date-to').value   = today();

function setRange(days, btn){{
  document.querySelectorAll('.btn-preset').forEach(b=>b.classList.remove('active'));
  if(btn) btn.classList.add('active');
  if(days===0){{ sel('f-date-from').value=''; sel('f-date-to').value=''; }}
  else{{ sel('f-date-from').value=daysAgo(days); sel('f-date-to').value=today(); }}
  applyFilters();
}}

// Header badges
function fmtMonth(m) {{
  const [y,mo] = m.split('-');
  return new Date(y, mo-1).toLocaleString('en-US', {{month:'long',year:'numeric'}});
}}
function fmtDate(d){{
  const [y,m,day]=d.split('-');
  return new Date(+y,+m-1,+day).toLocaleDateString('en-US',{{month:'long',day:'numeric',year:'numeric'}});
}}
const maxDate = RAW.reduce((mx,r)=>r.Date>mx?r.Date:mx,'');
if(maxDate) sel('badge-asof').textContent = 'Data through '+fmtDate(maxDate);

let filtered = [...RAW];
function getFilters(){{
  return {{
    fromDate: sel('f-date-from').value,
    toDate:   sel('f-date-to').value,
    playbook: sel('f-playbook').value,
    region:   sel('f-region').value,
    type:     sel('f-type').value,
  }};
}}

const TLG = new Set(["Jason Ackerman","Bianca Davis","James Parker","Resmie Biba","Chris Curtis","Sara Thompson","Jeremy MacBean","Bradley Pierce","Laura Sefcik","Samantha Maresca","Staci Musco","CJ Homer","Rich Moore","Dale Kinsey"]);
let hideTLG = false;
let lastVisitMap = {{}};

function toggleTLG(){{
  hideTLG = !hideTLG;
  sel('btn-tlg').classList.toggle('active', hideTLG);
  sel('btn-tlg').textContent = hideTLG ? 'Show TLG' : 'Hide TLG';
  applyFilters();
}}

const VERTICAL_PLAYBOOKS = new Set(["Healthcare Playbook","Legal Playbook","Public Sector Playbook"]);
let hideVertical = false;

function toggleVertical(){{
  hideVertical = !hideVertical;
  sel('btn-vertical').classList.toggle('active', hideVertical);
  sel('btn-vertical').textContent = hideVertical ? 'All Playbooks' : 'Vertical Markets';
  applyFilters();
}}

function applyFilters(){{
  const f = getFilters();
  filtered = RAW.filter(r => {{
    if (hideTLG && TLG.has(`${{r.FirstName}} ${{r.LastName}}`)) return false;
    if (hideVertical && !VERTICAL_PLAYBOOKS.has(r.Playbook)) return false;
    if (f.fromDate && r.Date < f.fromDate) return false;
    if (f.toDate   && r.Date > f.toDate)   return false;
    if (f.playbook && r.Playbook !== f.playbook) return false;
    if (f.region   && r.Region   !== f.region)   return false;
    if (f.type     && r.Type     !== f.type)      return false;
    return true;
  }});
  render();
}}

['f-date-from','f-date-to'].forEach(id => sel(id).addEventListener('change', ()=>{{
  document.querySelectorAll('.btn-preset').forEach(b=>b.classList.remove('active'));
  applyFilters();
}}));
['f-playbook','f-region','f-type'].forEach(id => sel(id).addEventListener('change', applyFilters));

function resetFilters(){{
  document.querySelectorAll('.btn-preset').forEach(b=>b.classList.remove('active'));
  sel('btn-30d').classList.add('active');
  sel('f-date-from').value = daysAgo(30);
  sel('f-date-to').value   = today();
  ['f-playbook','f-region','f-type'].forEach(id => sel(id).value = '');
  hideTLG = false;
  sel('btn-tlg').classList.remove('active');
  sel('btn-tlg').textContent = 'Hide TLG';
  hideVertical = false;
  sel('btn-vertical').classList.remove('active');
  sel('btn-vertical').textContent = 'Vertical Markets';
  applyFilters();
}}

function filterForRange(fromStr, toStr){{
  const f=getFilters();
  return RAW.filter(r=>{{
    if(hideTLG && TLG.has(r.FirstName+' '+r.LastName)) return false;
    if(hideVertical && !VERTICAL_PLAYBOOKS.has(r.Playbook)) return false;
    if(fromStr && r.Date<fromStr) return false;
    if(toStr   && r.Date>toStr)   return false;
    if(f.playbook && r.Playbook!==f.playbook) return false;
    if(f.region   && r.Region!==f.region)     return false;
    if(f.type     && r.Type!==f.type)         return false;
    return true;
  }});
}}

function filterPersonList(){{
  const q=(sel('drilldown-search')?.value||'').toLowerCase();
  document.querySelectorAll('.drilldown-person').forEach(el=>{{
    el.style.display=(!q||el.dataset.name.toLowerCase().includes(q))?'':'none';
  }});
}}

function countBy(arr, key){{
  return arr.reduce((acc,r) => {{ const v=r[key]||'(none)'; acc[v]=(acc[v]||0)+1; return acc; }}, {{}});
}}

let barChart, pieChart, trendChart, pagesChart;

function render(){{
  const isLight = document.body.classList.contains('light-mode');
  const chartLabel = isLight ? cv('--text') : cv('--muted');
  // Build per-person visit counts first (needed for sort)
  const visitorMap = {{}};
  lastVisitMap = {{}};
  filtered.forEach(r => {{
    const key = `${{r.FirstName}} ${{r.LastName}}`;
    visitorMap[key] = (visitorMap[key] || 0) + 1;
    if (!lastVisitMap[key] || r.Date > lastVisitMap[key]) lastVisitMap[key] = r.Date;
  }});

  // Stats
  const totalViews  = filtered.length;
  const uniqueUsers = new Set(filtered.map(r=>`${{r.FirstName}} ${{r.LastName}}`)).size;
  const pbCounts    = countBy(filtered, 'Playbook');
  const topPB       = Object.entries(pbCounts).sort((a,b)=>b[1]-a[1])[0] || ['—',0];
  const monthsShown = new Set(filtered.map(r=>r.Month)).size;
  sel('result-count').textContent = `${{totalViews.toLocaleString()}} views`;

  const avgVisits = uniqueUsers > 0 ? (totalViews / uniqueUsers).toFixed(1) : '0';

  // Prior period comparison (same date-range length, immediately before)
  const fromDate=sel('f-date-from').value, toDate=sel('f-date-to').value;
  let viewsDelta=null, usersDelta=null, avgDelta=null;
  if(fromDate && toDate){{
    const dFrom=new Date(fromDate+'T00:00:00'), dTo=new Date(toDate+'T00:00:00');
    const span=Math.round((dTo-dFrom)/86400000)+1;
    const pTo=new Date(dFrom); pTo.setDate(pTo.getDate()-1);
    const pFrom=new Date(pTo); pFrom.setDate(pFrom.getDate()-span+1);
    const prior=filterForRange(pFrom.toISOString().slice(0,10), pTo.toISOString().slice(0,10));
    const pViews=prior.length;
    const pUsers=new Set(prior.map(r=>r.FirstName+' '+r.LastName)).size;
    const pAvg=pUsers>0?(pViews/pUsers):0;
    if(pViews>0) viewsDelta=Math.round((totalViews-pViews)/pViews*100);
    if(pUsers>0) usersDelta=Math.round((uniqueUsers-pUsers)/pUsers*100);
    if(pAvg>0)   avgDelta=Math.round((parseFloat(avgVisits)-pAvg)/pAvg*100);
  }}
  function dHtml(d){{
    if(d===null) return '';
    const arrow=d>=0?'↑':'↓', color=d>=0?'var(--green)':'var(--red)';
    return '<span style="font-size:10px;color:'+color+';margin-left:6px;font-weight:600">'+arrow+Math.abs(d)+'%</span>';
  }}

  sel('stats-row').innerHTML = `
    <div class="stat"><div class="stat-label">Total Page Views<span class="info-btn" onclick="showInfo(event,'total-views')">?</span></div><div class="stat-value blue">${{totalViews.toLocaleString()}}${{dHtml(viewsDelta)}}</div><div class="stat-sub">${{monthsShown}} month${{monthsShown!==1?'s':''}} shown</div></div>
    <div class="stat"><div class="stat-label">Unique Users<span class="info-btn" onclick="showInfo(event,'unique-users')">?</span></div><div class="stat-value purple">${{uniqueUsers}}${{dHtml(usersDelta)}}</div><div class="stat-sub">employees &amp; dealers</div></div>
    <div class="stat"><div class="stat-label">Avg Visits / Person<span class="info-btn" onclick="showInfo(event,'avg-visits')">?</span></div><div class="stat-value" style="color:${{cv('--teal')}}">${{avgVisits}}${{dHtml(avgDelta)}}</div><div class="stat-sub">vs prior period</div></div>
    <div class="stat"><div class="stat-label">Top Playbook<span class="info-btn" onclick="showInfo(event,'top-playbook')">?</span></div><div class="stat-value yellow" style="font-size:16px;padding-top:4px">${{topPB[0]}}</div><div class="stat-sub">${{topPB[1].toLocaleString()}} views</div></div>
    <div class="stat"><div class="stat-label">Playbooks Active<span class="info-btn" onclick="showInfo(event,'playbooks-active')">?</span></div><div class="stat-value green">${{Object.keys(pbCounts).length}}</div><div class="stat-sub">out of 9 total</div></div>
  `;


  // Bar chart — always include all known playbooks (zero if no data)
  const pbAll = {{...pbCounts}};
  Object.keys(PLAYBOOK_COLORS).forEach(pb => {{ if (!(pb in pbAll)) pbAll[pb] = 0; }});
  const pbSorted = Object.entries(pbAll).sort((a,b)=>b[1]-a[1]);
  if (barChart) barChart.destroy();
  barChart = new Chart(sel('barChart'), {{
    type: 'bar',
    data: {{ labels: pbSorted.map(([k])=>k), datasets: [{{ data: pbSorted.map(([,v])=>v), backgroundColor: pbSorted.map(([k])=>pbColor(k)), borderRadius: 5, borderSkipped: false }}] }},
    options: {{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{{
        legend:{{display:false}},
        tooltip:{{callbacks:{{label:c=>` ${{c.raw.toLocaleString()}} views`}}}},
        datalabels: typeof ChartDataLabels !== 'undefined' ? {{
          anchor:'center', align:'center',
          color: document.body.classList.contains('light-mode') ? '#1e293b' : '#fff', font:{{size:11, weight:'700'}},
          formatter: v => v > 0 ? v.toLocaleString() : '',
        }} : {{display:false}}
      }},
      scales:{{
        x:{{grid:{{color:cv('--border')}},ticks:{{color:chartLabel,font:{{size:11}}}}}},
        y:{{grid:{{display:false}},ticks:{{color:cv('--text'),font:{{size:11}}}}}}
      }}
    }}
  }});

  // Pie chart
  const regCounts = countBy(filtered, 'Region');
  const pieLabels = Object.keys(regCounts).sort((a,b)=>regCounts[b]-regCounts[a]);
  const pieColors = ['#4f8ef7','#3ecf8e','#f7c94f','#7c5cfc','#f76f6f','#f7944f','#5cf0f7','#cf5cf7','#7b82a0'];
  if (pieChart) pieChart.destroy();
  pieChart = new Chart(sel('pieChart'), {{
    type: 'doughnut',
    data: {{ labels: pieLabels, datasets: [{{ data: pieLabels.map(k=>regCounts[k]), backgroundColor: pieColors.slice(0,pieLabels.length), borderWidth:0 }}] }},
    options: {{
      responsive:true, maintainAspectRatio:false,
      plugins:{{
        legend:{{position:'bottom',labels:{{color:chartLabel,font:{{size:11}},boxWidth:10,padding:8}}}},
        tooltip:{{callbacks:{{label:c=>` ${{c.label}}: ${{c.raw.toLocaleString()}} views`}}}},
        datalabels:{{display:false}}
      }}
    }}
  }});

  // Top pages chart
  const pbFilter = sel('f-playbook').value;
  const pageMap = {{}};
  filtered.forEach(r => {{
    const label = pbFilter ? r.Page : `${{r.Page}} · ${{r.Playbook}}`;
    if (!pageMap[label]) pageMap[label] = {{count:0, visitors:new Set(), color:pbColor(r.Playbook)}};
    pageMap[label].count++;
    pageMap[label].visitors.add(`${{r.FirstName}} ${{r.LastName}}`);
  }});
  const pagesSorted = Object.entries(pageMap).sort((a,b)=>b[1].count-a[1].count).slice(0,10);
  const pageAvgs     = pagesSorted.map(([,v]) => (v.count/v.visitors.size).toFixed(1));
  const pageVisitors = pagesSorted.map(([,v]) => v.visitors.size);
  sel('pages-chart-title').textContent = pbFilter ? `Top Pages — ${{pbFilter}}` : 'Top Pages — All Playbooks';
  if (pagesChart) pagesChart.destroy();
  pagesChart = new Chart(sel('pagesChart'), {{
    type: 'bar',
    data: {{ labels: pagesSorted.map(([k])=>k), datasets: [{{ data: pagesSorted.map(([,v])=>v.count), backgroundColor: pagesSorted.map(([,v])=>v.color), borderRadius:5, borderSkipped:false }}] }},
    options: {{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{{ legend:{{display:false}}, tooltip:{{callbacks:{{label:c=>` ${{c.raw.toLocaleString()}} views · ${{pageVisitors[c.dataIndex]}} visitors · ${{pageAvgs[c.dataIndex]}} avg visits/person`}}}}, datalabels:{{display:false}} }},
      scales:{{
        x:{{grid:{{color:cv('--border')}},ticks:{{color:chartLabel,font:{{size:11}}}}}},
        y:{{grid:{{display:false}},ticks:{{color:cv('--text'),font:{{size:11}}}}}}
      }}
    }}
  }});

  // Trend chart — views per month per playbook
  const visibleMonths = [...new Set(filtered.map(r=>r.Month))].sort();
  const topPlaybooks  = Object.entries(pbCounts).sort((a,b)=>b[1]-a[1]).slice(0,5).map(([k])=>k);
  const trendDatasets = topPlaybooks.map(pb => {{
    const color = pbColor(pb);
    return {{
      label: pb,
      data: visibleMonths.map(m => filtered.filter(r=>r.Month===m && r.Playbook===pb).length),
      borderColor: color,
      backgroundColor: color + '22',
      borderWidth: 2,
      pointRadius: 4,
      tension: 0.3,
      fill: false,
    }};
  }});
  if (trendChart) trendChart.destroy();
  trendChart = new Chart(sel('trendChart'), {{
    type: 'line',
    data: {{ labels: visibleMonths, datasets: trendDatasets }},
    options: {{
      responsive:true, maintainAspectRatio:false,
      plugins:{{
        legend:{{position:'bottom',labels:{{color:chartLabel,font:{{size:11}},boxWidth:10,padding:8}}}},
        tooltip:{{mode:'index',intersect:false}},
        datalabels:{{display:false}}
      }},
      scales:{{
        x:{{grid:{{color:cv('--border')}},ticks:{{color:chartLabel,font:{{size:11}}}}}},
        y:{{grid:{{color:cv('--border')}},ticks:{{color:chartLabel,font:{{size:11}}}}}}
      }}
    }}
  }});

  // Who's Active — left panel
  const personList = Object.entries(visitorMap).sort((a,b)=>b[1]-a[1]);
  sel('drilldown-left').innerHTML = personList.map(([name, count]) => {{
    const dot = `<span class="recency-dot" style="background:${{recencyColor(lastVisitMap[name])}}"></span>`;
    return `<div class="drilldown-person" onclick="drillSelect(this,'${{name.replace(/'/g,"\\'")}}')" data-name="${{name}}">
       ${{dot}}
       <span class="drilldown-name">${{name}}</span>
       <span class="drilldown-count">${{count}}</span>
     </div>`;
  }}).join('') || `<div class="no-data">No data</div>`;
  filterPersonList();
  if (personList.length) {{
    const first = sel('drilldown-left').querySelector('.drilldown-person:not([style*="none"])') || sel('drilldown-left').querySelector('.drilldown-person');
    if(first) drillSelect(first, first.dataset.name);
  }} else {{
    sel('drilldown-right').innerHTML = `<div class="no-data">No records match your filters.</div>`;
  }}
}}

function drillSelect(el, name) {{
  document.querySelectorAll('.drilldown-person').forEach(p => p.classList.remove('active'));
  el.classList.add('active');
  const visits = filtered.filter(r => `${{r.FirstName}} ${{r.LastName}}` === name);
  const first = visits[0];
  const region = first?.Region || '—';
  const type = first?.Type || '—';
  const typeColor = type==='Employee'?cv('--pill-emp-color'):cv('--pill-dlr-color');
  const typeBg = type==='Employee'?cv('--pill-emp-bg'):cv('--pill-dlr-bg');

  // Group by Date + Playbook + Page, count each combo
  const grouped = {{}};
  visits.forEach(v => {{
    const key = `${{v.Date}}|${{v.Playbook}}|${{v.Page}}`;
    if (!grouped[key]) grouped[key] = {{date:v.Date, playbook:v.Playbook, page:v.Page, count:0}};
    grouped[key].count++;
  }});
  const rows = Object.values(grouped).sort((a,b) => b.date.localeCompare(a.date));

  const lastVisit = lastVisitMap[name] || '';
  const lastVisitColor = recencyColor(lastVisit);
  sel('drilldown-right').innerHTML = `
    <div class="drilldown-right-header">
      <strong style="font-size:14px">${{name}}</strong>
      <span style="color:var(--muted)"> · ${{visits.length}} visit${{visits.length!==1?'s':''}} · ${{region}} · </span>
      <span class="pill" style="background:${{typeBg}};color:${{typeColor}}">${{type}}</span>
      <span style="color:var(--muted)"> · Last visit: </span><span style="color:${{lastVisitColor}};font-weight:600">${{lastVisit||'—'}}</span>
    </div>
    <table style="width:100%;border-collapse:collapse;">
      <thead><tr>
        <th>Date</th><th>Playbook</th><th>Page</th><th style="text-align:right">Visits</th>
      </tr></thead>
      <tbody>
        ${{rows.map(v => {{
          const c = pbColor(v.playbook);
          return `<tr>
            <td style="color:var(--muted)">${{v.date}}</td>
            <td><span class="pill pill-pb" style="background:${{c}}22;color:${{c}}">${{v.playbook}}</span></td>
            <td style="color:var(--muted)">${{v.page}}</td>
            <td style="text-align:right;color:${{cv('--teal')}};font-weight:600">${{v.count}}</td>
          </tr>`;
        }}).join('')}}
      </tbody>
    </table>
  `;
}}

applyFilters();

const INFO = {{
  'total-views':    'Each time someone loads a page in a playbook, that counts as one page view. A single person visiting 3 different pages = 3 views.',
  'unique-users':   'The number of distinct people who accessed a playbook in the selected period. Each person is counted once regardless of how many pages they viewed.',
  'avg-visits':     'Total page views divided by unique users. A score of 3.0 means each person viewed an average of 3 pages during this period. Higher scores indicate people are exploring more content or returning to pages repeatedly.',
  'top-playbook':   'The playbook with the highest total page views in the selected period.',
  'playbooks-active': 'The number of playbooks that received at least one view in the selected period, out of 9 total playbooks tracked.',
  'chart-playbook': 'Total page views per playbook for the selected period. The number shown inside each bar is the exact view count. Hover a bar for more detail.',
  'chart-region':   'Breakdown of total page views by sales region. Hover over a segment to see the exact count for that region.',
  'chart-trend':    'Page view trends over time. Only the top 5 playbooks by total volume are shown — lower-traffic playbooks are not included on this chart.',
  'chart-pages':    'The 10 most visited pages in the selected period. Hover over any bar to see total views, unique visitors, and average visits per person.',
  'hide-tlg':       'Removes the internal L&D team (TLG) from all data — charts, stats, and the Active Users panel. Use this when sharing results with managers or stakeholders outside the team.',
  'vertical-filter': 'Filters to only the three vertical market playbooks — Healthcare, Legal, and Public Sector. All other playbooks are hidden while active. Can be combined with Hide TLG.',
  'whos-active':    'Lists every person who accessed a playbook in the selected period, sorted by total page views. The number next to each name is their total page views — not unique pages visited. Clicking a page 5 times counts as 5.',
}};

function showInfo(e, key) {{
  e.stopPropagation();
  const pop = document.getElementById('info-popover');
  if (pop.dataset.key === key && pop.classList.contains('visible')) {{
    pop.classList.remove('visible');
    return;
  }}
  pop.dataset.key = key;
  pop.textContent = INFO[key] || '';
  pop.classList.add('visible');
  const r = e.target.getBoundingClientRect();
  const left = Math.min(r.left, window.innerWidth - 280);
  pop.style.left = left + 'px';
  pop.style.top = (r.bottom + 8 + window.scrollY) + 'px';
  pop.style.position = 'absolute';
}}
function toggleHamburger(){{
  const menu = sel('hamburger-menu');
  const btn  = sel('hamburger-btn');
  const open = menu.classList.toggle('open');
  btn.classList.toggle('open', open);
}}
document.addEventListener('click', function(e){{
  const h = sel('hamburger');
  if(h && !h.contains(e.target)){{
    sel('hamburger-menu').classList.remove('open');
    sel('hamburger-btn').classList.remove('open');
  }}
  document.getElementById('info-popover')?.classList.remove('visible');
}});
</script>
<div id="info-popover" class="info-popover"></div>
</body>
</html>"""

OUTPUT_FILE.write_text(html, encoding='utf-8')
print(f"\nDashboard written to: {OUTPUT_FILE}")
print(f"Total records: {total_rows:,} across {len(months)} month(s)")
