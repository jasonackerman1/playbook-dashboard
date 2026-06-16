#!/usr/bin/env python3
"""update_cert_dashboard.py — generates cert-{vertical}.html from LMS Excel exports in cert-data/.
Accumulates all monthly files per vertical; deduplicates by Email so each person appears once."""

import openpyxl
import os
import re
import json
import warnings
warnings.filterwarnings('ignore')

TLG = {
    "Jason Ackerman","Bianca Davis","James Parker","Resmie Biba",
    "Chris Curtis","Sara Thompson","Jeremy MacBean","Bradley Pierce",
    "Laura Sefcik","Samantha Maresca","Staci Musco","CJ Homer","Rich Moore","Dale Kinsey"
}

VERTICAL_MAP = {
    'healthcare':   'Healthcare',
    'publicsector': 'Public Sector',
    'legal':        'Legal',
    'salesforce':   'Salesforce',
    'dx':           'DX Playbook',
    'gcip':         'GC/IP Sales',
    'iq501':        'IQ501',
    'roadtodx':     'Road to DX',
}

# Column indices (0-based) matching Resmie's LMS export format
COL_FIRST        = 2
COL_LAST         = 3
COL_EMAIL        = 4
COL_JOBTITLE     = 5
COL_REGION       = 6
COL_MGR_FIRST    = 9
COL_MGR_LAST     = 10
COL_MGR_EMAIL    = 11
COL_MGR_TITLE    = 12
COL_COMPLETE     = 19   # "Yes" / "No"
COL_DATE         = 20   # Curriculum Assignment Date (completion date)
COL_QTR          = 21   # Qtr Certified
COL_LAYERED_SEC  = 22   # Layered Security Certified Status
COL_HEALTHCARE   = 23   # Healthcare Certified Status
COL_AMBULATORY   = 24   # Ambulatory Certified
COL_EXTENDED     = 25   # Extended Care Certified

SUB_CERTS = [
    ('LayeredSec', 'Layered Security'),
    ('Healthcare', 'Healthcare'),
    ('Ambulatory', 'Ambulatory'),
    ('Extended',   'Extended Care'),
]


def km_fiscal_quarter(date):
    """KM fiscal year starts April. Q1=Apr-Jun, Q2=Jul-Sep, Q3=Oct-Dec, Q4=Jan-Mar."""
    if date is None:
        return ''
    import math
    m = date.month
    # Shift so April=1 ... March=12, then divide into 4 equal quarters
    q = math.ceil(((m - 4) % 12 + 1) / 3)
    fy = date.year if m >= 4 else date.year - 1
    return f'Q{q} FY{fy}'


def extract_file_date(fname):
    """Extract YYYY-MM from filename for chronological ordering of monthly files."""
    m = re.search(r'(\d{4}-\d{2})', os.path.basename(fname))
    return m.group(1) if m else '0000-00'


def person_key(r):
    """Unique identifier for deduplication — Email preferred, name as fallback."""
    return r['Email'].lower() if r['Email'] else f"{r['FirstName']} {r['LastName']}".lower()


def detect_vertical(fname):
    """Extract vertical slug from filename — supports both naming conventions."""
    fn = os.path.basename(fname).lower()
    # cert-healthcare-2026-05.xlsx
    m = re.match(r'cert-([a-z0-9]+)-', fn)
    if m:
        return m.group(1)
    # FY26-Healthcare-Certification-Courses_2026-05.xlsx
    m = re.match(r'fy\d+-([a-z]+)-', fn)
    if m:
        return m.group(1).lower()
    return None


def load_rows(filepath):
    """Load non-blank rows, return list of dicts."""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active
    rows = []
    for raw in ws.iter_rows(min_row=2, values_only=True):
        if raw[COL_FIRST] is None:
            continue
        rows.append({
            'FirstName': str(raw[COL_FIRST]).strip(),
            'LastName':  str(raw[COL_LAST]).strip(),
            'Email':     str(raw[COL_EMAIL]).strip() if raw[COL_EMAIL] else '',
            'JobTitle':  str(raw[COL_JOBTITLE]).strip() if raw[COL_JOBTITLE] else '',
            'Region':    str(raw[COL_REGION]).strip() if raw[COL_REGION] else '',
            'Manager':   ((str(raw[COL_MGR_FIRST]).strip() + ' ' + str(raw[COL_MGR_LAST]).strip()).strip()) if raw[COL_MGR_FIRST] else '',
            'MgrEmail':  str(raw[COL_MGR_EMAIL]).strip() if raw[COL_MGR_EMAIL] else '',
            'MgrTitle':  str(raw[COL_MGR_TITLE]).strip() if raw[COL_MGR_TITLE] else '',
            'Complete':   str(raw[COL_COMPLETE]).strip() if raw[COL_COMPLETE] else 'No',
            'Date':       raw[COL_DATE].strftime('%Y-%m-%d') if raw[COL_DATE] else '',
            'Qtr':        str(raw[COL_QTR]).strip() if raw[COL_QTR] else km_fiscal_quarter(raw[COL_DATE]),
            'LayeredSec': str(raw[COL_LAYERED_SEC]).strip() if raw[COL_LAYERED_SEC] else 'No',
            'Healthcare': str(raw[COL_HEALTHCARE]).strip() if raw[COL_HEALTHCARE] else 'No',
            'Ambulatory': str(raw[COL_AMBULATORY]).strip() if raw[COL_AMBULATORY] else 'No',
            'Extended':   str(raw[COL_EXTENDED]).strip() if raw[COL_EXTENDED] else 'No',
        })
    wb.close()
    return rows


def generate_html(slug, name, rows):
    raw_json = json.dumps(rows)
    tlg_json = json.dumps(sorted(TLG))

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{name} Certification Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root {{
    --bg:#0f1117; --surface:#1a1d27; --surface2:#22263a; --border:#2e3350;
    --accent:#4f8ef7; --accent2:#7c5cfc; --accent3:#f7c94f;
    --text:#e8ecf4; --muted:#7b82a0; --green:#3ecf8e; --red:#f76f6f;
    --teal:#2dd4bf; --green-subtle:#3ecf8e22; --red-subtle:#f76f6f22;
    --font:'Segoe UI',system-ui,sans-serif;
  }}
  body.light-mode {{
    --bg:#f4f6fb; --surface:#ffffff; --surface2:#eef1f7; --border:#d0d7e8;
    --accent:#2563eb; --accent2:#6d28d9; --accent3:#d97706;
    --text:#1a1d27; --muted:#475569; --green:#059669; --red:#dc2626;
    --teal:#0f766e; --green-subtle:#05966922; --red-subtle:#dc262622;
  }}
  body.light-mode select,body.light-mode input[type=date]{{color-scheme:light;}}
  *{{box-sizing:border-box;margin:0;padding:0;}}
  body{{background:var(--bg);color:var(--text);font-family:var(--font);min-height:100vh;transition:background .2s,color .2s;}}

  .header{{padding:20px 28px 16px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;}}
  .header-left{{display:flex;align-items:center;gap:16px;}}
  .hamburger{{position:relative;}}
  .hamburger-btn{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:6px 10px;font-size:16px;cursor:pointer;transition:all .15s;line-height:1;}}
  .hamburger-btn:hover,.hamburger-btn.open{{border-color:var(--accent);color:var(--text);}}
  .hamburger-menu{{position:absolute;top:calc(100% + 6px);left:0;background:var(--surface);border:1px solid var(--border);border-radius:8px;min-width:220px;box-shadow:0 4px 24px rgba(0,0,0,0.28);display:none;z-index:200;overflow:hidden;}}
  .hamburger-menu.open{{display:block;}}
  .hamburger-section-label{{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--muted);padding:10px 14px 4px;}}
  .hamburger-item{{display:flex;align-items:center;gap:8px;padding:10px 14px;font-size:13px;color:var(--text);text-decoration:none;transition:background .1s;}}
  .hamburger-item:hover{{background:var(--surface2);}}
  .info-btn{{display:inline-flex;align-items:center;justify-content:center;width:15px;height:15px;border-radius:50%;background:var(--surface2);border:1px solid var(--border);color:var(--muted);font-size:9px;font-weight:700;cursor:pointer;margin-left:5px;vertical-align:middle;flex-shrink:0;line-height:1;transition:border-color .15s,color .15s;}}
  .info-btn:hover{{border-color:var(--accent);color:var(--accent);}}
  .info-popover{{position:fixed;z-index:9999;background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:12px 14px;font-size:12px;color:var(--text);line-height:1.6;max-width:260px;box-shadow:0 4px 24px rgba(0,0,0,0.5);display:none;}}
  .info-popover.visible{{display:block;}}
  .header h1{{font-size:18px;font-weight:700;letter-spacing:.3px;}}
  .header h1 span{{color:var(--muted);font-weight:400;}}
  .btn-theme{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:5px 12px;font-size:12px;cursor:pointer;transition:all .15s;}}
  .btn-theme:hover{{border-color:var(--accent);color:var(--text);}}

  .filters{{padding:14px 28px;display:flex;gap:10px;flex-wrap:wrap;align-items:center;border-bottom:1px solid var(--border);background:var(--surface);}}
  .filter-label{{font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.6px;margin-right:4px;}}
  select,input[type=date],input[type=text]{{background:var(--surface2);border:1px solid var(--border);color:var(--text);border-radius:6px;padding:6px 10px;font-size:13px;cursor:pointer;outline:none;color-scheme:dark;}}
  select:focus,input:focus{{border-color:var(--accent);}}
  .btn-reset{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:6px 14px;font-size:12px;cursor:pointer;transition:border-color .15s,color .15s;}}
  .btn-reset:hover{{border-color:var(--accent);color:var(--text);}}
  .btn-tlg{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:6px 14px;font-size:12px;cursor:pointer;transition:all .15s;}}
  .btn-tlg:hover{{border-color:var(--red);color:var(--red);}}
  .btn-tlg.active{{background:var(--red-subtle);border-color:var(--red);color:var(--red);}}
  .result-count{{margin-left:auto;font-size:12px;color:var(--muted);}}

  .stats{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;padding:20px 28px;}}
  .stat{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:16px 18px;}}
  .stat-label{{font-size:11px;text-transform:uppercase;letter-spacing:.7px;color:var(--muted);margin-bottom:6px;}}
  .stat-value{{font-size:28px;font-weight:700;line-height:1;}}
  .stat-value.green{{color:var(--green);}}
  .stat-value.red{{color:var(--red);}}
  .stat-value.teal{{color:var(--teal);}}
  .stat-sub{{font-size:11px;color:var(--muted);margin-top:4px;}}

  .charts{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;padding:0 28px 16px;}}
  @media(max-width:1100px){{.charts{{grid-template-columns:1fr 1fr;}}}}
  @media(max-width:680px){{.charts{{grid-template-columns:1fr;}}}}
  .chart-card{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:18px;}}
  .chart-title{{font-size:13px;font-weight:600;margin-bottom:14px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;}}
  body.light-mode .chart-title{{color:var(--text);}}
  .chart-wrap{{position:relative;height:260px;}}

  .section{{padding:0 28px 32px;}}
  .section-header{{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;flex-wrap:wrap;gap:8px;}}
  .section-title{{font-size:13px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;}}
  body.light-mode .section-title{{color:var(--text);}}
  .section-hint{{font-size:11px;color:var(--muted);margin-top:3px;}}

  .roster-search{{width:180px;}}
  .roster-wrap{{display:flex;border:1px solid var(--border);border-radius:10px;overflow:hidden;}}
  .roster-left{{width:300px;flex-shrink:0;overflow-y:auto;max-height:820px;border-right:1px solid var(--border);}}
  .roster-person{{display:flex;align-items:center;gap:10px;padding:10px 14px;cursor:pointer;border-bottom:1px solid var(--border);transition:background .1s;}}
  .roster-person:last-child{{border-bottom:none;}}
  .roster-person:hover{{background:var(--surface2);}}
  .roster-person.active{{background:#4f8ef711;border-left:3px solid var(--accent);padding-left:11px;}}
  .roster-name{{flex:1;font-size:13px;}}
  .cert-dot{{width:8px;height:8px;border-radius:50%;flex-shrink:0;}}
  .cert-dot.yes{{background:var(--green);}}
  .cert-dot.no{{background:var(--red);}}
  .cert-badge{{font-size:11px;font-weight:700;border-radius:10px;padding:2px 8px;white-space:nowrap;}}
  .cert-badge.yes{{color:var(--green);background:var(--green-subtle);}}
  .cert-badge.no{{color:var(--red);background:var(--red-subtle);}}
  .roster-right{{flex:1;overflow-y:auto;max-height:820px;padding:16px 20px;}}
  .roster-right-header{{margin-bottom:14px;padding-bottom:12px;border-bottom:1px solid var(--border);}}
  .no-data{{text-align:center;color:var(--muted);padding:40px;font-size:13px;}}
  @media(max-width:680px){{.roster-wrap{{flex-direction:column;}}.roster-left{{width:100%;max-height:220px;border-right:none;border-bottom:1px solid var(--border);}}}}

  .detail-grid{{display:grid;grid-template-columns:1fr 1fr;gap:12px 28px;margin-top:14px;}}
  .detail-label{{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:3px;}}
  .detail-value{{font-size:14px;font-weight:500;}}
  .badge-status{{display:inline-block;padding:3px 12px;border-radius:20px;font-size:12px;font-weight:700;}}
  .badge-status.certified{{background:var(--green-subtle);color:var(--green);}}
  .badge-status.not-certified{{background:var(--red-subtle);color:var(--red);}}
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <div class="hamburger" id="hamburger">
      <button class="hamburger-btn" id="hamburger-btn" onclick="toggleHamburger()" aria-label="Menu">&#9776;</button>
      <div class="hamburger-menu" id="hamburger-menu">
        <div class="hamburger-section-label">Dashboards</div>
        <a href="index.html" class="hamburger-item">&#128202; Playbook Dashboard</a>
      </div>
    </div>
    <h1>{name} <span>Certification Dashboard</span></h1>
  </div>
  <button class="btn-theme" id="btn-theme" onclick="toggleTheme()">&#9728; Light</button>
</div>

<div class="filters">
  <span class="filter-label">Certification</span>
  <select id="f-cert">
    <option value="">All Certifications</option>
    <option value="LayeredSec">Layered Security</option>
    <option value="Healthcare">Healthcare</option>
    <option value="Ambulatory">Ambulatory</option>
    <option value="Extended">Extended Care</option>
  </select>
  <span class="filter-label">Status</span>
  <select id="f-status">
    <option value="">All</option>
    <option value="Yes">Certified</option>
    <option value="No">Not Certified</option>
  </select>
  <span class="filter-label">Region</span>
  <select id="f-region"><option value="">All Regions</option></select>
  <span class="filter-label" style="margin-right:2px">From</span>
  <input type="date" id="f-date-from">
  <span class="filter-label" style="margin:0 2px">To</span>
  <input type="date" id="f-date-to">
  <button class="btn-reset" onclick="resetFilters()">Reset</button>
  <button class="btn-tlg" id="btn-tlg" onclick="toggleTLG()">Hide TLG</button>
  <span class="result-count" id="result-count"></span>
</div>

<div class="stats">
  <div class="stat">
    <div class="stat-label">Total Assigned <span class="info-btn" onclick="showInfo(event,'total-assigned')">?</span></div>
    <div class="stat-value" id="s-total">&#8212;</div>
  </div>
  <div class="stat">
    <div class="stat-label">Certified <span class="info-btn" onclick="showInfo(event,'certified')">?</span></div>
    <div class="stat-value green" id="s-certified">&#8212;</div>
  </div>
  <div class="stat">
    <div class="stat-label">Not Yet Certified <span class="info-btn" onclick="showInfo(event,'not-certified')">?</span></div>
    <div class="stat-value red" id="s-not">&#8212;</div>
  </div>
  <div class="stat">
    <div class="stat-label">Completion Rate <span class="info-btn" onclick="showInfo(event,'completion-rate')">?</span></div>
    <div class="stat-value teal" id="s-rate">&#8212;</div>
    <div class="stat-sub" id="s-rate-sub"></div>
  </div>
</div>

<div class="charts">
  <div class="chart-card">
    <div class="chart-title">Certified vs Not Certified by Region <span class="info-btn" onclick="showInfo(event,'by-region')">?</span></div>
    <div class="chart-wrap"><canvas id="regionChart"></canvas></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">Certifications Over Time <span class="info-btn" onclick="showInfo(event,'over-time')">?</span></div>
    <div class="chart-wrap"><canvas id="trendChart"></canvas></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">Sub-Certification Breakdown <span class="info-btn" onclick="showInfo(event,'subcert')">?</span></div>
    <div class="chart-wrap"><canvas id="subcertChart"></canvas></div>
  </div>
</div>

<div class="section">
  <div class="section-header">
    <div>
      <div class="section-title">Certification Roster <span class="info-btn" onclick="showInfo(event,'roster')">?</span></div>
      <div class="section-hint">Click a person to see their details &mdash; sorted by status then name</div>
    </div>
    <input type="text" class="roster-search" id="roster-search" placeholder="Search by name&hellip;" oninput="filterRoster()">
  </div>
  <div class="roster-wrap">
    <div class="roster-left" id="roster-left"></div>
    <div class="roster-right" id="roster-right">
      <div class="no-data">Select a person to see their details</div>
    </div>
  </div>
</div>

<script>
const RAW = {raw_json};
const TLG_SET = new Set({tlg_json});

let filtered = [];
let hideTLG = false;
let regionChart, trendChart, subcertChart;

function sel(id){{ return document.getElementById(id); }}
function cv(v){{ return getComputedStyle(document.body).getPropertyValue(v).trim(); }}

// Theme
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

// Populate filter dropdowns from data
const allRegions = [...new Set(RAW.map(r=>r.Region).filter(Boolean))].sort();
allRegions.forEach(r => sel('f-region').innerHTML += `<option value="${{r}}">${{r}}</option>`);

['f-cert','f-status','f-region','f-date-from','f-date-to'].forEach(id => {{
  sel(id).addEventListener('change', applyFilters);
}});

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
  if(!e.target.classList.contains('info-btn')){{
    sel('info-popover').classList.remove('visible');
  }}
}});

const INFO_MSGS = {{
  'total-assigned':  'Total number of people assigned this certification curriculum.',
  'certified':       'People who have completed and passed the curriculum certification.',
  'not-certified':   'People assigned the curriculum who have not yet completed it.',
  'completion-rate': 'Percentage of assigned people who have earned the certification. Updates when a specific Certification is selected.',
  'by-region':       'Certified vs. not certified broken down by sales region. Responds to the Certification and Status filters.',
  'over-time':       'Number of overall curriculum certifications earned per fiscal quarter (Q1 = Apr–Jun, Q2 = Jul–Sep, Q3 = Oct–Dec, Q4 = Jan–Mar).',
  'subcert':         'How many people have earned each of the four sub-certifications within this curriculum.',
  'roster':          'All assigned people with their certification status. Click a name to see details, sub-certification badges, and manager contact info.',
}};
function showInfo(e, key){{
  const pop = sel('info-popover');
  pop.textContent = INFO_MSGS[key] || '';
  pop.classList.add('visible');
  const r = e.target.getBoundingClientRect();
  pop.style.top  = (r.bottom + 6) + 'px';
  pop.style.left = Math.min(r.left, window.innerWidth - 280) + 'px';
  e.stopPropagation();
}}

function toggleTLG(){{
  hideTLG = !hideTLG;
  sel('btn-tlg').classList.toggle('active', hideTLG);
  sel('btn-tlg').textContent = hideTLG ? 'Show TLG' : 'Hide TLG';
  applyFilters();
}}

function resetFilters(){{
  ['f-cert','f-status','f-region','f-date-from','f-date-to'].forEach(id => sel(id).value = '');
  if(hideTLG){{ hideTLG=false; sel('btn-tlg').classList.remove('active'); sel('btn-tlg').textContent='Hide TLG'; }}
  applyFilters();
}}

function applyFilters(){{
  const cert   = sel('f-cert').value;
  const certField = cert || 'Complete';
  const status = sel('f-status').value;
  const region = sel('f-region').value;
  const from   = sel('f-date-from').value;
  const to     = sel('f-date-to').value;
  const effectiveStatus = (cert && !status) ? 'Yes' : status;
  filtered = RAW.filter(r => {{
    if(hideTLG && TLG_SET.has(r.FirstName+' '+r.LastName)) return false;
    if(effectiveStatus && r[certField] !== effectiveStatus) return false;
    if(region && r.Region   !== region) return false;
    if(from && r.Date && r.Date < from) return false;
    if(to   && r.Date && r.Date > to)   return false;
    return true;
  }});
  sel('result-count').textContent = `${{filtered.length}} person${{filtered.length!==1?'s':''}}`;
  render();
}}

function render(){{
  const isLight    = document.body.classList.contains('light-mode');
  const chartLabel = isLight ? cv('--text') : cv('--muted');
  const certField  = sel('f-cert').value || 'Complete';

  // Stat cards
  const total    = filtered.length;
  const certified = filtered.filter(r=>r[certField]==='Yes').length;
  const notCert  = total - certified;
  const rate     = total > 0 ? Math.round(certified/total*100) : 0;

  sel('s-total').textContent     = total;
  sel('s-certified').textContent = certified;
  sel('s-not').textContent       = notCert;
  sel('s-rate').textContent      = rate + '%';
  sel('s-rate-sub').textContent  = total > 0 ? `${{certified}} of ${{total}} assigned` : '';

  // Region chart — grouped bars
  const regions         = [...new Set(filtered.map(r=>r.Region).filter(Boolean))].sort();
  const certByRegion    = regions.map(rg => filtered.filter(r=>r.Region===rg&&r[certField]==='Yes').length);
  const notCertByRegion = regions.map(rg => filtered.filter(r=>r.Region===rg&&r[certField]==='No').length);

  if(regionChart) regionChart.destroy();
  regionChart = new Chart(sel('regionChart'), {{
    type: 'bar',
    data: {{
      labels: regions,
      datasets: [
        {{ label:'Certified',     data:certByRegion,    backgroundColor:cv('--green')+'cc', borderRadius:4, borderSkipped:false }},
        {{ label:'Not Certified', data:notCertByRegion, backgroundColor:cv('--red')+'cc',   borderRadius:4, borderSkipped:false }},
      ]
    }},
    options: {{
      responsive:true, maintainAspectRatio:false,
      plugins:{{
        legend:{{position:'bottom',labels:{{color:chartLabel,font:{{size:11}},boxWidth:10,padding:8}}}},
        tooltip:{{mode:'index',intersect:false}},
        datalabels:{{display:false}}
      }},
      scales:{{
        x:{{grid:{{color:cv('--border')}},ticks:{{color:chartLabel,font:{{size:11}}}}}},
        y:{{grid:{{color:cv('--border')}},ticks:{{color:chartLabel,font:{{size:11}},stepSize:1}}}}
      }}
    }}
  }});

  // Trend chart — certified people per fiscal quarter
  const yesRows = filtered.filter(r=>r.Complete==='Yes'&&r.Qtr);
  const qtrMap = {{}};
  yesRows.forEach(r => {{ qtrMap[r.Qtr] = (qtrMap[r.Qtr]||0)+1; }});
  const trendMonths = Object.keys(qtrMap).sort((a,b) => {{
    const parse = s => {{ const m=s.match(/Q(\d)\s+FY(\d+)/); return m ? parseInt(m[2])*10+parseInt(m[1]) : 0; }};
    return parse(a)-parse(b);
  }});
  const trendCounts = trendMonths.map(q=>qtrMap[q]);

  if(trendChart) trendChart.destroy();
  trendChart = new Chart(sel('trendChart'), {{
    type: 'bar',
    data: {{ labels:trendMonths, datasets:[{{
      data:trendCounts,
      backgroundColor:cv('--accent')+'99',
      borderRadius:4,
      borderSkipped:false
    }}]}},
    options: {{
      responsive:true, maintainAspectRatio:false,
      plugins:{{
        legend:{{display:false}},
        tooltip:{{callbacks:{{label:c=>` ${{c.raw}} certified`}}}},
        datalabels:{{display:false}}
      }},
      scales:{{
        x:{{grid:{{color:cv('--border')}},ticks:{{color:chartLabel,font:{{size:10}},maxRotation:45}}}},
        y:{{grid:{{color:cv('--border')}},ticks:{{color:chartLabel,font:{{size:11}},stepSize:1}}}}
      }}
    }}
  }});

  // Sub-cert breakdown chart
  const subLabels = ['Layered Security','Healthcare','Ambulatory','Extended Care'];
  const subKeys   = ['LayeredSec','Healthcare','Ambulatory','Extended'];
  const subColors = [cv('--accent'),cv('--green'),cv('--accent2'),cv('--teal')];
  const subCounts = subKeys.map(k => filtered.filter(r=>r[k]==='Yes').length);
  const subNotCounts = subKeys.map(k => filtered.filter(r=>r[k]==='No').length);

  if(subcertChart) subcertChart.destroy();
  subcertChart = new Chart(sel('subcertChart'), {{
    type: 'bar',
    data: {{
      labels: subLabels,
      datasets: [
        {{ label:'Certified',     data:subCounts,    backgroundColor:subColors.map(c=>c+'cc'), borderRadius:4, borderSkipped:false }},
        {{ label:'Not Certified', data:subNotCounts, backgroundColor:cv('--surface2'), borderRadius:4, borderSkipped:false, borderWidth:1, borderColor:cv('--border') }},
      ]
    }},
    options: {{
      responsive:true, maintainAspectRatio:false,
      plugins:{{
        legend:{{position:'bottom',labels:{{color:chartLabel,font:{{size:11}},boxWidth:10,padding:8}}}},
        tooltip:{{mode:'index',intersect:false}},
        datalabels:{{display:false}}
      }},
      scales:{{
        x:{{grid:{{color:cv('--border')}},ticks:{{color:chartLabel,font:{{size:10}}}},stacked:true}},
        y:{{grid:{{color:cv('--border')}},ticks:{{color:chartLabel,font:{{size:11}},stepSize:1}},stacked:true}}
      }}
    }}
  }});

  // Roster — certified first, then alphabetical by last name
  const sorted = [...filtered].sort((a,b)=>{{
    if(a[certField]!==b[certField]) return a[certField]==='Yes'?-1:1;
    return (a.LastName+a.FirstName).localeCompare(b.LastName+b.FirstName);
  }});
  sel('roster-left').innerHTML = sorted.map(p=>{{
    const fullName = `${{p.FirstName}} ${{p.LastName}}`;
    const isCert   = p[certField]==='Yes';
    return `<div class="roster-person" onclick="rosterSelect(this)" data-name="${{fullName}}">
      <div class="cert-dot ${{isCert?'yes':'no'}}"></div>
      <span class="roster-name">${{fullName}}</span>
      <span class="cert-badge ${{isCert?'yes':'no'}}">${{isCert?'Certified':'Not Yet'}}</span>
    </div>`;
  }}).join('') || `<div class="no-data">No people match filters</div>`;

  filterRoster();

  const first = sel('roster-left').querySelector('.roster-person:not([style*="none"])') ||
                sel('roster-left').querySelector('.roster-person');
  if(first) rosterSelect(first);
}}

function rosterSelect(el){{
  document.querySelectorAll('.roster-person').forEach(p=>p.classList.remove('active'));
  el.classList.add('active');
  const name = el.dataset.name;
  const p = filtered.find(r=>`${{r.FirstName}} ${{r.LastName}}`===name);
  if(!p) return;
  const isCert = p.Complete==='Yes';
  sel('roster-right').innerHTML = `
    <div class="roster-right-header">
      <div style="font-size:15px;font-weight:700;margin-bottom:6px">${{p.FirstName}} ${{p.LastName}}</div>
      <span class="badge-status ${{isCert?'certified':'not-certified'}}">${{isCert?'Certified':'Not Yet Certified'}}</span>
    </div>
    <div class="detail-grid">
      <div><div class="detail-label">Job Title</div><div class="detail-value">${{p.JobTitle||'&#8212;'}}</div></div>
      <div><div class="detail-label">Region</div><div class="detail-value">${{p.Region||'&#8212;'}}</div></div>
      <div><div class="detail-label">Email</div><div class="detail-value"><a href="mailto:${{p.Email}}" style="color:var(--accent);text-decoration:none">${{p.Email||'&#8212;'}}</a></div></div>
      <div><div class="detail-label">Certification Date</div><div class="detail-value">${{p.Date||'&#8212;'}}</div></div>
      <div><div class="detail-label">Quarter Certified</div><div class="detail-value">${{p.Qtr||'&#8212;'}}</div></div>
    </div>
    <div style="margin-top:14px;padding-top:12px;border-top:1px solid var(--border)">
      <div style="font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px">Sub-Certifications</div>
      <div style="display:flex;flex-wrap:wrap;gap:8px">
        ${{[['LayeredSec','Layered Security'],['Healthcare','Healthcare'],['Ambulatory','Ambulatory'],['Extended','Extended Care']].map(([k,label])=>{{
          const yes = p[k]==='Yes';
          return `<span style="font-size:12px;font-weight:600;padding:4px 12px;border-radius:20px;background:${{yes?'var(--green-subtle)':'var(--surface2)'}};color:${{yes?'var(--green)':'var(--muted)'}};border:1px solid ${{yes?'var(--green)':'var(--border)'}}">${{yes?'&#10003;':'&#8212;'}} ${{label}}</span>`;
        }}).join('')}}
      </div>
    </div>
    ${{p.Manager ? `<div style="margin-top:16px;padding-top:14px;border-top:1px solid var(--border)">
      <div style="font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px">Manager</div>
      <div class="detail-grid">
        <div><div class="detail-label">Name</div><div class="detail-value">${{p.Manager}}</div></div>
        <div><div class="detail-label">Title</div><div class="detail-value">${{p.MgrTitle||'&#8212;'}}</div></div>
        <div><div class="detail-label">Email</div><div class="detail-value"><a href="mailto:${{p.MgrEmail}}" style="color:var(--accent);text-decoration:none">${{p.MgrEmail||'&#8212;'}}</a></div></div>
      </div>
    </div>` : ''}}
  `;
}}

function filterRoster(){{
  const q = (sel('roster-search')?.value||'').toLowerCase();
  document.querySelectorAll('.roster-person').forEach(el=>{{
    el.style.display = (!q||el.dataset.name.toLowerCase().includes(q)) ? '' : 'none';
  }});
}}

applyFilters();
</script>
<div class="info-popover" id="info-popover"></div>
</body>
</html>"""


def main():
    cert_dir = 'cert-data'
    if not os.path.isdir(cert_dir):
        print(f'No {cert_dir}/ directory found.')
        return

    files = sorted(f for f in os.listdir(cert_dir) if f.endswith('.xlsx'))
    if not files:
        print('No .xlsx files found in cert-data/')
        return

    # Group files by vertical slug
    vert_files = {}
    for fname in files:
        slug = detect_vertical(fname)
        if not slug:
            print(f'  Skipping {fname} — could not detect vertical from filename')
            continue
        vert_files.setdefault(slug, []).append(fname)

    generated = []
    for slug in sorted(vert_files):
        # Sort files chronologically by YYYY-MM in filename
        fnames    = sorted(vert_files[slug], key=extract_file_date)
        vert_name = VERTICAL_MAP.get(slug, slug.title())
        print(f'\n{slug} ({vert_name}) — {len(fnames)} file(s):')

        # Load all files in order, tag each row with its file date
        all_rows = []
        for fname in fnames:
            file_date = extract_file_date(fname)
            filepath  = os.path.join(cert_dir, fname)
            print(f'  {fname}  [{file_date}]')
            rows = load_rows(filepath)
            for r in rows:
                r['_file_date'] = file_date
            all_rows.extend(rows)
            print(f'    {len(rows)} rows  ({sum(1 for r in rows if r["Complete"]=="Yes")} certified)')

        # Deduplicate: later files overwrite earlier ones for the same person
        all_rows.sort(key=lambda r: r['_file_date'])
        seen = {}
        for r in all_rows:
            seen[person_key(r)] = r
        deduped = list(seen.values())
        for r in deduped:
            del r['_file_date']

        certified = sum(1 for r in deduped if r['Complete'] == 'Yes')
        print(f'  → {len(deduped)} unique people  ({certified} certified, {len(deduped)-certified} not yet)')

        html = generate_html(slug, vert_name, deduped)
        out  = f'cert-{slug}.html'
        with open(out, 'w', encoding='utf-8') as fh:
            fh.write(html)
        print(f'  Written → {out}')
        generated.append(out)

    print(f'\nDone. {len(generated)} dashboard(s) generated.')


if __name__ == '__main__':
    main()
