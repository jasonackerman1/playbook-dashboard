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
  .search-wrap input{{width:220px;}}
  .table-outer{{border-radius:10px;border:1px solid var(--border);overflow:hidden;}}
  table{{width:100%;border-collapse:collapse;font-size:13px;}}
  thead tr{{background:var(--surface2);}}
  thead th{{padding:10px 14px;text-align:left;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);cursor:pointer;user-select:none;white-space:nowrap;}}
  thead th:hover{{color:var(--text);}} thead th.sorted{{color:var(--accent);}}
  tbody tr{{border-top:1px solid var(--border);transition:background .1s;}}
  tbody tr:hover{{background:var(--surface2);}}
  tbody td{{padding:9px 14px;vertical-align:middle;}}
  .pill{{display:inline-block;padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600;white-space:nowrap;}}
  .pagination{{display:flex;justify-content:center;gap:6px;margin-top:16px;flex-wrap:wrap;}}
  .page-btn{{background:var(--surface2);border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:5px 10px;font-size:12px;cursor:pointer;}}
  .page-btn:hover,.page-btn.active{{border-color:var(--accent);color:var(--accent);}}
  .no-data{{text-align:center;color:var(--muted);padding:40px;font-size:13px;}}
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
  <span class="result-count" id="result-count"></span>
</div>

<div class="stats" id="stats-row"></div>

<div class="charts-top">
  <div class="chart-card">
    <div class="chart-title">Page Views by Playbook</div>
    <div class="chart-wrap"><canvas id="barChart"></canvas></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">Views by Region</div>
    <div class="chart-wrap"><canvas id="pieChart"></canvas></div>
  </div>
</div>

<div class="charts-bottom">
  <div class="chart-card">
    <div class="chart-title">Monthly Trend — Views Over Time</div>
    <div class="chart-wrap-tall"><canvas id="trendChart"></canvas></div>
  </div>
</div>

<div class="charts-pages">
  <div class="chart-card">
    <div class="chart-title" id="pages-chart-title">Top Pages — Total Views</div>
    <div class="chart-wrap-pages"><canvas id="pagesChart"></canvas></div>
  </div>
  <div class="chart-card">
    <div class="chart-title" id="avg-pages-chart-title">Top Pages — Avg Visits / Person</div>
    <div class="chart-wrap-pages"><canvas id="avgPagesChart"></canvas></div>
  </div>
</div>

<div class="table-section" id="table-section">
  <div class="table-header">
    <span class="table-title">Activity Log</span>
    <div class="search-wrap">
      <input type="text" id="search" placeholder="Search name, page..." oninput="applyFilters()">
    </div>
  </div>
  <div class="table-outer">
    <table>
      <thead>
        <tr>
          <th onclick="sortBy('name')">Name</th>
          <th onclick="sortBy('Region')">Region</th>
          <th onclick="sortBy('Type')">Type</th>
          <th onclick="sortBy('Month')">Month</th>
          <th onclick="sortBy('Date')">Date</th>
          <th onclick="sortBy('Playbook')">Playbook</th>
          <th onclick="sortBy('Page')">Page</th>
          <th>Total Visits</th>
        </tr>
      </thead>
      <tbody id="table-body"></tbody>
    </table>
  </div>
  <div class="pagination" id="pagination"></div>
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
const allPlaybooks= [...new Set(RAW.map(r=>r.Playbook))].sort();
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
let sortCol = 'Date', sortDir = -1, page = 1;
const PAGE_SIZE = 30;

function getFilters(){{
  return {{
    month:    sel('f-month').value,
    playbook: sel('f-playbook').value,
    region:   sel('f-region').value,
    type:     sel('f-type').value,
    search:   sel('search').value.toLowerCase().trim(),
  }};
}}

function applyFilters(){{
  const f = getFilters();
  filtered = RAW.filter(r => {{
    if (f.month    && r.Month    !== f.month)    return false;
    if (f.playbook && r.Playbook !== f.playbook) return false;
    if (f.region   && r.Region   !== f.region)   return false;
    if (f.type     && r.Type     !== f.type)      return false;
    if (f.search){{
      const hay = `${{r.FirstName}} ${{r.LastName}} ${{r.Page}} ${{r.Playbook}} ${{r.Region||''}} ${{r.Month}}`.toLowerCase();
      if (!hay.includes(f.search)) return false;
    }}
    return true;
  }});
  page = 1;
  render();
}}

['f-month','f-playbook','f-region','f-type'].forEach(id => sel(id).addEventListener('change', applyFilters));

function resetFilters(){{
  ['f-month','f-playbook','f-region','f-type'].forEach(id => sel(id).value = '');
  sel('search').value = '';
  applyFilters();
}}

function sortBy(col){{
  if (sortCol === col) sortDir *= -1;
  else {{ sortCol = col; sortDir = 1; }}
  render();
}}

function countBy(arr, key){{
  return arr.reduce((acc,r) => {{ const v=r[key]||'(none)'; acc[v]=(acc[v]||0)+1; return acc; }}, {{}});
}}

let barChart, pieChart, trendChart, pagesChart, avgPagesChart;

function render(){{
  const sorted = [...filtered].sort((a,b) => {{
    let av = sortCol==='name' ? `${{a.FirstName}} ${{a.LastName}}` : a[sortCol]||'';
    let bv = sortCol==='name' ? `${{b.FirstName}} ${{b.LastName}}` : b[sortCol]||'';
    return av < bv ? -sortDir : av > bv ? sortDir : 0;
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

  // Build per-person visit counts for activity log
  const visitorMap = {{}};
  filtered.forEach(r => {{
    const key = `${{r.FirstName}} ${{r.LastName}}`;
    if (!visitorMap[key]) visitorMap[key] = 0;
    visitorMap[key]++;
  }});

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
    if (!pageMap[label]) pageMap[label] = {{count:0, color:pbColor(r.Playbook)}};
    pageMap[label].count++;
  }});
  const pagesSorted = Object.entries(pageMap).sort((a,b)=>b[1].count-a[1].count).slice(0,10);
  sel('pages-chart-title').textContent = pbFilter ? `Top Pages — ${{pbFilter}}` : 'Top Pages — All Playbooks';
  if (pagesChart) pagesChart.destroy();
  pagesChart = new Chart(sel('pagesChart'), {{
    type: 'bar',
    data: {{ labels: pagesSorted.map(([k])=>k), datasets: [{{ data: pagesSorted.map(([,v])=>v.count), backgroundColor: pagesSorted.map(([,v])=>v.color), borderRadius:5, borderSkipped:false }}] }},
    options: {{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{{ legend:{{display:false}}, tooltip:{{callbacks:{{label:c=>` ${{c.raw.toLocaleString()}} views`}}}} }},
      scales:{{
        x:{{grid:{{color:'#2e3350'}},ticks:{{color:'#7b82a0',font:{{size:11}}}}}},
        y:{{grid:{{display:false}},ticks:{{color:'#e8ecf4',font:{{size:11}}}}}}
      }}
    }}
  }});

  // Avg visits per page chart
  const avgPageMap = {{}};
  filtered.forEach(r => {{
    const label = pbFilter ? r.Page : `${{r.Page}} · ${{r.Playbook}}`;
    if (!avgPageMap[label]) avgPageMap[label] = {{views:0, visitors:new Set(), color:pbColor(r.Playbook)}};
    avgPageMap[label].views++;
    avgPageMap[label].visitors.add(`${{r.FirstName}} ${{r.LastName}}`);
  }});
  const avgPagesSorted = Object.entries(avgPageMap)
    .map(([k,v]) => [k, {{avg:parseFloat((v.views/v.visitors.size).toFixed(2)), color:v.color}}])
    .sort((a,b)=>b[1].avg-a[1].avg).slice(0,10);
  sel('avg-pages-chart-title').textContent = pbFilter ? `Avg Visits / Person — ${{pbFilter}}` : 'Avg Visits / Person — By Page';
  if (avgPagesChart) avgPagesChart.destroy();
  avgPagesChart = new Chart(sel('avgPagesChart'), {{
    type: 'bar',
    data: {{ labels: avgPagesSorted.map(([k])=>k), datasets: [{{ data: avgPagesSorted.map(([,v])=>v.avg), backgroundColor: avgPagesSorted.map(([,v])=>v.color), borderRadius:5, borderSkipped:false }}] }},
    options: {{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{{ legend:{{display:false}}, tooltip:{{callbacks:{{label:c=>` ${{c.raw}} avg visits/person`}}}} }},
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

  // Table
  const pageData   = sorted.slice((page-1)*PAGE_SIZE, page*PAGE_SIZE);
  const totalPages = Math.max(1, Math.ceil(sorted.length/PAGE_SIZE));
  const tbody = sel('table-body');
  if (!pageData.length) {{
    tbody.innerHTML = `<tr><td colspan="8" class="no-data">No records match your filters.</td></tr>`;
  }} else {{
    tbody.innerHTML = pageData.map(r => {{
      const color = pbColor(r.Playbook);
      const totalVisits = visitorMap[`${{r.FirstName}} ${{r.LastName}}`] || 0;
      return `<tr>
        <td>${{r.FirstName}} ${{r.LastName}}</td>
        <td>${{r.Region||'<span style="color:var(--muted)">—</span>'}}</td>
        <td><span class="pill" style="background:${{r.Type==='Employee'?'#1a2a4a':'#2a1a3a'}};color:${{r.Type==='Employee'?'#4f8ef7':'#cf5cf7'}}">${{r.Type}}</span></td>
        <td style="color:var(--muted)">${{r.Month}}</td>
        <td style="color:var(--muted)">${{r.Date}}</td>
        <td><span class="pill" style="background:${{color}}22;color:${{color}}">${{r.Playbook}}</span></td>
        <td style="color:var(--muted);max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${{r.Page}}</td>
        <td style="color:#2dd4bf;font-weight:600">${{totalVisits}}</td>
      </tr>`;
    }}).join('');
  }}

  // Pagination
  const pg = sel('pagination');
  pg.innerHTML = '';
  if (totalPages <= 1) return;
  const makeBtn = (label, p, active=false) => {{
    const b = document.createElement('button');
    b.className = 'page-btn' + (active?' active':'');
    b.textContent = label;
    b.onclick = () => {{ page = p; render(); sel('table-section').scrollIntoView({{behavior:'smooth'}}); }};
    pg.appendChild(b);
  }};
  if (page > 1) makeBtn('← Prev', page-1);
  const start=Math.max(1,page-2), end=Math.min(totalPages,page+2);
  for (let i=start; i<=end; i++) makeBtn(i, i, i===page);
  if (page < totalPages) makeBtn('Next →', page+1);
}}

applyFilters();
</script>
</body>
</html>"""

OUTPUT_FILE.write_text(html, encoding='utf-8')
print(f"\nDashboard written to: {OUTPUT_FILE}")
print(f"Total records: {total_rows:,} across {len(months)} month(s)")
