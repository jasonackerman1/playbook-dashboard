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
<style>
  :root {{
    --bg:#0f1117; --surface:#1a1d27; --surface2:#22263a; --border:#2e3350;
    --accent:#4f8ef7; --accent2:#7c5cfc; --accent3:#f7c94f;
    --text:#e8ecf4; --muted:#7b82a0; --green:#3ecf8e; --red:#f76f6f;
    --font:'Segoe UI',system-ui,sans-serif;
  }}
  *{{box-sizing:border-box;margin:0;padding:0;}}
  body{{background:var(--bg);color:var(--text);font-family:var(--font);min-height:100vh;}}

  .header{{padding:20px 28px 16px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;}}
  .header h1{{font-size:18px;font-weight:700;letter-spacing:.3px;}}
  .header h1 span{{color:var(--muted);font-weight:400;}}
  .badges{{display:flex;gap:8px;flex-wrap:wrap;}}
  .badge{{background:var(--surface2);border:1px solid var(--border);border-radius:20px;padding:4px 12px;font-size:12px;color:var(--muted);}}

  .filters{{padding:14px 28px;display:flex;gap:10px;flex-wrap:wrap;align-items:center;border-bottom:1px solid var(--border);background:var(--surface);}}
  .filter-label{{font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.6px;margin-right:4px;}}
  select,input[type=text]{{background:var(--surface2);border:1px solid var(--border);color:var(--text);border-radius:6px;padding:6px 10px;font-size:13px;cursor:pointer;outline:none;}}
  select:focus,input:focus{{border-color:var(--accent);}}
  .btn-reset{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:6px 14px;font-size:12px;cursor:pointer;transition:border-color .15s,color .15s;}}
  .btn-reset:hover{{border-color:var(--accent);color:var(--text);}}
  .btn-tlg{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:6px 14px;font-size:12px;cursor:pointer;transition:all .15s;}}
  .btn-tlg:hover{{border-color:var(--red);color:var(--red);}}
  .btn-tlg.active{{background:#f76f6f22;border-color:var(--red);color:var(--red);}}
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
  .drilldown-wrap{{display:flex;border:1px solid var(--border);border-radius:10px;overflow:hidden;}}
  .drilldown-left{{width:260px;flex-shrink:0;overflow-y:auto;max-height:820px;border-right:1px solid var(--border);}}
  .drilldown-person{{display:flex;align-items:center;gap:10px;padding:10px 14px;cursor:pointer;border-bottom:1px solid var(--border);transition:background .1s;}}
  .drilldown-person:last-child{{border-bottom:none;}}
  .drilldown-person:hover{{background:var(--surface2);}}
  .drilldown-person.active{{background:#4f8ef711;border-left:3px solid var(--accent);padding-left:11px;}}
  .drilldown-name{{flex:1;font-size:13px;}}
  .drilldown-count{{font-size:11px;font-weight:700;color:#2dd4bf;background:#2dd4bf18;border-radius:10px;padding:2px 8px;}}
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
  <h1>Playbook Traffic Dashboard <span id="header-range"></span></h1>
  <div class="badges">
    <span class="badge" id="badge-views"></span>
    <span class="badge" id="badge-months"></span>
  </div>
</div>

<div class="filters">
  <span class="filter-label">Filter</span>
  <select id="f-month"><option value="">All Months</option></select>
  <select id="f-playbook"><option value="">All Playbooks</option></select>
  <select id="f-region"><option value="">All Regions</option></select>
  <select id="f-type"><option value="">Employee &amp; Dealer</option><option value="Employee">Employee</option><option value="Dealer">Dealer</option></select>
  <button class="btn-reset" onclick="resetFilters()">Reset</button>
  <button class="btn-tlg" id="btn-tlg" onclick="toggleTLG()">Hide TLG</button>
  <span class="result-count" id="result-count"></span>
</div>

<div class="stats" id="stats-row"></div>

<div class="charts-top">
  <div class="chart-card">
    <div class="chart-title">Page Views by Playbook</div>
    <div class="section-hint">Hover over a bar to see the view count</div>
    <div class="chart-wrap"><canvas id="barChart"></canvas></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">Views by Region</div>
    <div class="section-hint">Hover over a segment to see the region breakdown</div>
    <div class="chart-wrap"><canvas id="pieChart"></canvas></div>
  </div>
</div>

<div class="charts-bottom">
  <div class="chart-card">
    <div class="chart-title">Monthly Trend — Views Over Time</div>
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
    <span class="table-title">Who's Active</span>
    <span class="section-hint" style="margin:0">Click a name to see their full page history</span>
  </div>
  <div class="drilldown-wrap">
    <div class="drilldown-left" id="drilldown-left"></div>
    <div class="drilldown-right" id="drilldown-right"><div class="no-data">Select a person to see their activity</div></div>
  </div>
</div>

<script>
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

const allMonths   = [...new Set(RAW.map(r=>r.Month))].sort();
const allPlaybooks= [...new Set([...Object.keys(PLAYBOOK_COLORS), ...RAW.map(r=>r.Playbook)])].sort();
const allRegions  = [...new Set(RAW.map(r=>r.Region).filter(Boolean))].sort();

function sel(id){{ return document.getElementById(id); }}
allMonths.forEach(m => sel('f-month').innerHTML    += `<option value="${{m}}">${{m}}</option>`);
allPlaybooks.forEach(p => sel('f-playbook').innerHTML += `<option value="${{p}}">${{p}}</option>`);
allRegions.forEach(r => sel('f-region').innerHTML  += `<option value="${{r}}">${{r}}</option>`);

// Header badges
function fmtMonth(m) {{
  const [y,mo] = m.split('-');
  return new Date(y, mo-1).toLocaleString('en-US', {{month:'long',year:'numeric'}});
}}
sel('header-range').textContent = allMonths.length > 1
  ? `— ${{fmtMonth(allMonths[0])}} to ${{fmtMonth(allMonths[allMonths.length-1])}}`
  : `— ${{fmtMonth(allMonths[0])}}`;
sel('badge-views').textContent  = `${{RAW.length.toLocaleString()}} total views`;
sel('badge-months').textContent = `${{allMonths.length}} month${{allMonths.length>1?'s':''}}`;

let filtered = [...RAW];
function getFilters(){{
  return {{
    month:    sel('f-month').value,
    playbook: sel('f-playbook').value,
    region:   sel('f-region').value,
    type:     sel('f-type').value,
  }};
}}

const TLG = new Set(["Jason Ackerman","Bianca Davis","James Parker","Resmie Biba","Chris Curtis","Sara Thompson","Jeremy MacBean","Bradley Pierce","Laura Sefcik","Samantha Maresca","Staci Musco","CJ Homer","Rich Moore","Dale Kinsey"]);
let hideTLG = false;

function toggleTLG(){{
  hideTLG = !hideTLG;
  sel('btn-tlg').classList.toggle('active', hideTLG);
  sel('btn-tlg').textContent = hideTLG ? 'Show TLG' : 'Hide TLG';
  applyFilters();
}}

function applyFilters(){{
  const f = getFilters();
  filtered = RAW.filter(r => {{
    if (hideTLG && TLG.has(`${{r.FirstName}} ${{r.LastName}}`)) return false;
    if (f.month    && r.Month    !== f.month)    return false;
    if (f.playbook && r.Playbook !== f.playbook) return false;
    if (f.region   && r.Region   !== f.region)   return false;
    if (f.type     && r.Type     !== f.type)      return false;
    return true;
  }});
  render();
}}

['f-month','f-playbook','f-region','f-type'].forEach(id => sel(id).addEventListener('change', applyFilters));

function resetFilters(){{
  ['f-month','f-playbook','f-region','f-type'].forEach(id => sel(id).value = '');
  hideTLG = false;
  sel('btn-tlg').classList.remove('active');
  sel('btn-tlg').textContent = 'Hide TLG';
  applyFilters();
}}

function countBy(arr, key){{
  return arr.reduce((acc,r) => {{ const v=r[key]||'(none)'; acc[v]=(acc[v]||0)+1; return acc; }}, {{}});
}}

let barChart, pieChart, trendChart, pagesChart;

function render(){{
  // Build per-person visit counts first (needed for sort)
  const visitorMap = {{}};
  filtered.forEach(r => {{
    const key = `${{r.FirstName}} ${{r.LastName}}`;
    visitorMap[key] = (visitorMap[key] || 0) + 1;
  }});

  // Stats
  const totalViews  = filtered.length;
  const uniqueUsers = new Set(filtered.map(r=>`${{r.FirstName}} ${{r.LastName}}`)).size;
  const pbCounts    = countBy(filtered, 'Playbook');
  const topPB       = Object.entries(pbCounts).sort((a,b)=>b[1]-a[1])[0] || ['—',0];
  const monthsShown = new Set(filtered.map(r=>r.Month)).size;
  sel('result-count').textContent = `${{totalViews.toLocaleString()}} views`;

  const avgVisits = uniqueUsers > 0 ? (totalViews / uniqueUsers).toFixed(1) : '0';

  sel('stats-row').innerHTML = `
    <div class="stat"><div class="stat-label">Total Page Views</div><div class="stat-value blue">${{totalViews.toLocaleString()}}</div><div class="stat-sub">${{monthsShown}} month${{monthsShown!==1?'s':''}} shown</div></div>
    <div class="stat"><div class="stat-label">Unique Users</div><div class="stat-value purple">${{uniqueUsers}}</div><div class="stat-sub">employees &amp; dealers</div></div>
    <div class="stat"><div class="stat-label">Avg Visits / Person</div><div class="stat-value" style="color:#2dd4bf">${{avgVisits}}</div><div class="stat-sub">this period</div></div>
    <div class="stat"><div class="stat-label">Top Playbook</div><div class="stat-value yellow" style="font-size:16px;padding-top:4px">${{topPB[0]}}</div><div class="stat-sub">${{topPB[1].toLocaleString()}} views</div></div>
    <div class="stat"><div class="stat-label">Playbooks Active</div><div class="stat-value green">${{Object.keys(pbCounts).length}}</div><div class="stat-sub">out of 9 total</div></div>
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
      plugins:{{ legend:{{display:false}}, tooltip:{{callbacks:{{label:c=>` ${{c.raw.toLocaleString()}} views`}}}} }},
      scales:{{
        x:{{grid:{{color:'#2e3350'}},ticks:{{color:'#7b82a0',font:{{size:11}}}}}},
        y:{{grid:{{display:false}},ticks:{{color:'#e8ecf4',font:{{size:11}}}}}}
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
        legend:{{position:'bottom',labels:{{color:'#7b82a0',font:{{size:11}},boxWidth:10,padding:8}}}},
        tooltip:{{callbacks:{{label:c=>` ${{c.label}}: ${{c.raw.toLocaleString()}} views`}}}}
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
      plugins:{{ legend:{{display:false}}, tooltip:{{callbacks:{{label:c=>` ${{c.raw.toLocaleString()}} views · ${{pageVisitors[c.dataIndex]}} visitors · ${{pageAvgs[c.dataIndex]}} avg visits/person`}}}} }},
      scales:{{
        x:{{grid:{{color:'#2e3350'}},ticks:{{color:'#7b82a0',font:{{size:11}}}}}},
        y:{{grid:{{display:false}},ticks:{{color:'#e8ecf4',font:{{size:11}}}}}}
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
        legend:{{position:'bottom',labels:{{color:'#7b82a0',font:{{size:11}},boxWidth:10,padding:8}}}},
        tooltip:{{mode:'index',intersect:false}}
      }},
      scales:{{
        x:{{grid:{{color:'#2e3350'}},ticks:{{color:'#7b82a0',font:{{size:11}}}}}},
        y:{{grid:{{color:'#2e3350'}},ticks:{{color:'#7b82a0',font:{{size:11}}}}}}
      }}
    }}
  }});

  // Who's Active — left panel
  const personList = Object.entries(visitorMap).sort((a,b)=>b[1]-a[1]);
  sel('drilldown-left').innerHTML = personList.map(([name, count]) =>
    `<div class="drilldown-person" onclick="drillSelect(this,'${{name.replace(/'/g,"\\'")}}')" data-name="${{name}}">
       <span class="drilldown-name">${{name}}</span>
       <span class="drilldown-count">${{count}}</span>
     </div>`
  ).join('') || `<div class="no-data">No data</div>`;
  if (personList.length) {{
    const first = sel('drilldown-left').querySelector('.drilldown-person');
    drillSelect(first, personList[0][0]);
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
  const typeColor = type==='Employee'?'#4f8ef7':'#cf5cf7';
  const typeBg = type==='Employee'?'#1a2a4a':'#2a1a3a';

  // Group by Date + Playbook + Page, count each combo
  const grouped = {{}};
  visits.forEach(v => {{
    const key = `${{v.Date}}|${{v.Playbook}}|${{v.Page}}`;
    if (!grouped[key]) grouped[key] = {{date:v.Date, playbook:v.Playbook, page:v.Page, count:0}};
    grouped[key].count++;
  }});
  const rows = Object.values(grouped).sort((a,b) => b.date.localeCompare(a.date));

  sel('drilldown-right').innerHTML = `
    <div class="drilldown-right-header">
      <strong style="font-size:14px">${{name}}</strong>
      <span style="color:var(--muted)"> · ${{visits.length}} visit${{visits.length!==1?'s':''}} · ${{region}} · </span>
      <span class="pill" style="background:${{typeBg}};color:${{typeColor}}">${{type}}</span>
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
            <td><span class="pill" style="background:${{c}}22;color:${{c}}">${{v.playbook}}</span></td>
            <td style="color:var(--muted)">${{v.page}}</td>
            <td style="text-align:right;color:#2dd4bf;font-weight:600">${{v.count}}</td>
          </tr>`;
        }}).join('')}}
      </tbody>
    </table>
  `;
}}

applyFilters();
</script>
</body>
</html>"""

OUTPUT_FILE.write_text(html, encoding='utf-8')
print(f"\nDashboard written to: {OUTPUT_FILE}")
print(f"Total records: {total_rows:,} across {len(months)} month(s)")
