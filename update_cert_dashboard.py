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

# Column indices (0-based) matching Resmie's LMS export format.
# CAUTION: verify against each new file before regenerating — Resmie may adjust columns.
# Last verified: FY26-Healthcare-Certification-Courses_2026-05.xlsx
COL_FIRST        = 2   # First Name
COL_LAST         = 3   # Last Name
COL_EMAIL        = 4   # Email Address
COL_JOBTITLE     = 5   # Job Title
COL_REGION       = 6   # Region (in data; Market used for grouping/filtering)
COL_MARKET       = 7   # Market
COL_MGR_FIRST    = 9   # ManagerFirstName
COL_MGR_LAST     = 10  # ManagerLastName
COL_MGR_EMAIL    = 11  # SUPEMAILADDR (manager email)
COL_MGR_TITLE    = 12  # Manager JobTitle
COL_HIRE_DATE    = 14  # Hire Date
COL_COMPLETE     = 19  # Curriculum Complete — "Yes"/"No" (LMS-driven)
COL_DATE         = 20  # Curriculum Assignment Date (enrollment date; used for date-range filter)
COL_HC_DATE      = 21  # Healthcare Certification Date (compute quarter from this)
# col 22 = HC Qtr Certified — Excel formula string, ignored; quarter computed via km_fiscal_quarter()
COL_HEALTHCARE   = 23  # HC Certified (manager sign-off)
COL_ACUTE_DATE   = 24  # Acute Care Certification Date
# col 25 = Acute Care Qtr Certified — Excel formula string, ignored
COL_ACUTE        = 26  # Acute Care Certified
COL_AMB_DATE     = 27  # Ambulatory Certification Date
# col 28 = Ambulatory Qtr Certified — Excel formula string, ignored
COL_AMBULATORY   = 29  # Ambulatory Certified
COL_EXT_DATE     = 30  # Extended Care Certification Date
# col 31 = Extended Care Qtr Certified — Excel formula string, ignored
COL_EXTENDED     = 32  # Extended Care Certified
# Layered Security: completely removed from new file

SUB_CERTS = [
    ('Healthcare', 'Healthcare'),
    ('AcuteCare',  'Acute Care'),
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
    return f'FY{fy} Q{q}'


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
        def _date(v): return v.strftime('%Y-%m-%d') if v and hasattr(v, 'strftime') else ''
        def _qtr(v):  return str(v).strip() if v and str(v).strip() not in ('None','') else ''
        def _cert(v): return str(v).strip() if v else 'No'
        def _qtr_from_date(v):
            return km_fiscal_quarter(v) if v and hasattr(v, 'month') else ''
        rows.append({
            'FirstName':  str(raw[COL_FIRST]).strip(),
            'LastName':   str(raw[COL_LAST]).strip(),
            'Email':      str(raw[COL_EMAIL]).strip() if raw[COL_EMAIL] else '',
            'JobTitle':   str(raw[COL_JOBTITLE]).strip() if raw[COL_JOBTITLE] else '',
            'Market':     str(raw[COL_MARKET]).strip() if raw[COL_MARKET] else '',
            'Manager':    ((str(raw[COL_MGR_FIRST]).strip() + ' ' + str(raw[COL_MGR_LAST]).strip()).strip()) if raw[COL_MGR_FIRST] else '',
            'MgrEmail':   str(raw[COL_MGR_EMAIL]).strip() if raw[COL_MGR_EMAIL] else '',
            'MgrTitle':   str(raw[COL_MGR_TITLE]).strip() if raw[COL_MGR_TITLE] else '',
            'HireDate':   _date(raw[COL_HIRE_DATE]),
            'Complete':   _cert(raw[COL_COMPLETE]),
            'Date':       _date(raw[COL_DATE]),
            'HCDate':     _date(raw[COL_HC_DATE]),
            'HCQtr':      _qtr_from_date(raw[COL_HC_DATE]),
            'Healthcare': _cert(raw[COL_HEALTHCARE]),
            'AcuteDate':  _date(raw[COL_ACUTE_DATE]),
            'AcuteQtr':   _qtr_from_date(raw[COL_ACUTE_DATE]),
            'AcuteCare':  _cert(raw[COL_ACUTE]),
            'AmbDate':    _date(raw[COL_AMB_DATE]),
            'AmbQtr':     _qtr_from_date(raw[COL_AMB_DATE]),
            'Ambulatory': _cert(raw[COL_AMBULATORY]),
            'ExtDate':    _date(raw[COL_EXT_DATE]),
            'ExtQtr':     _qtr_from_date(raw[COL_EXT_DATE]),
            'Extended':   _cert(raw[COL_EXTENDED]),
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
  .stat-value.blue{{color:var(--accent);}}
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

  .sort-btn{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:4px 10px;font-size:11px;cursor:pointer;transition:all .15s;white-space:nowrap;}}
  .sort-btn:hover{{border-color:var(--accent);color:var(--text);}}
  .sort-btn.active{{border-color:var(--accent);color:var(--accent);background:var(--accent)11;}}
  .mgr-group-hdr{{display:flex;align-items:center;justify-content:space-between;padding:10px 14px;cursor:pointer;background:var(--surface2);border-bottom:1px solid var(--border);user-select:none;}}
  .mgr-group-hdr:hover{{background:var(--border);}}
  .mgr-group-hdr.open .mgr-chevron{{transform:rotate(180deg);}}
  .mgr-chevron{{transition:transform .15s;font-size:10px;color:var(--muted);flex-shrink:0;}}
  .mgr-team{{display:none;}}
  .mgr-team.open{{display:block;}}
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
  <span class="filter-label">Curriculum</span>
  <select id="f-cert">
    <option value="">All Curriculums</option>
    <option value="Healthcare">Healthcare Foundational</option>
    <option value="AcuteCare">Acute Care</option>
    <option value="Ambulatory">Ambulatory</option>
    <option value="Extended">Extended Care</option>
  </select>
  <span class="filter-label">Status</span>
  <select id="f-status">
    <option value="">All</option>
    <option value="Yes">Certified</option>
    <option value="No">Not Certified</option>
  </select>
  <span class="filter-label">Market</span>
  <select id="f-market"><option value="">All Markets</option></select>
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
    <div class="stat-label">Curriculum Complete <span class="info-btn" onclick="showInfo(event,'curriculum-complete')">?</span></div>
    <div class="stat-value blue" id="s-curriculum">&#8212;</div>
    <div class="stat-sub" id="s-curriculum-sub"></div>
  </div>
  <div class="stat">
    <div class="stat-label">HC Certified <span class="info-btn" onclick="showInfo(event,'certified')">?</span></div>
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
    <div class="chart-title">Certification Pipeline <span class="info-btn" onclick="showInfo(event,'pipeline')">?</span></div>
    <div class="chart-wrap"><canvas id="regionChart"></canvas></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">Certifications Over Time <span class="info-btn" onclick="showInfo(event,'over-time')">?</span></div>
    <div class="chart-wrap"><canvas id="trendChart"></canvas></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">Curriculum Breakdown <span class="info-btn" onclick="showInfo(event,'subcert')">?</span></div>
    <div class="chart-wrap"><canvas id="subcertChart"></canvas></div>
  </div>
</div>

<div class="section">
  <div class="section-header">
    <div>
      <div class="section-title">Certification Roster <span class="info-btn" onclick="showInfo(event,'roster')">?</span></div>
      <div class="section-hint">Click a person to see their details &mdash; sorted by status then name</div>
    </div>
    <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;">
      <span style="font-size:11px;color:var(--muted);margin-right:2px;">View:</span>
      <button class="sort-btn active" id="view-individual" onclick="setRosterView('individual')">Individual</button>
      <button class="sort-btn" id="view-manager" onclick="setRosterView('manager')">By Manager</button>
      <span id="sort-controls" style="display:flex;align-items:center;gap:6px;margin-left:6px;">
        <span style="font-size:11px;color:var(--muted);margin-right:2px;">Sort:</span>
        <button class="sort-btn active" data-sort="status" data-label="Status" onclick="setRosterSort('status')">Status ↓</button>
        <button class="sort-btn" data-sort="name"   data-label="Name"   onclick="setRosterSort('name')">Name</button>
        <button class="sort-btn" data-sort="market" data-label="Market" onclick="setRosterSort('market')">Market</button>
      </span>
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
let rosterSortField = 'status';
let rosterSortDir   = 'desc';
let rosterView      = 'individual';

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
const allMarkets = [...new Set(RAW.map(r=>r.Market).filter(Boolean))].sort();
allMarkets.forEach(m => sel('f-market').innerHTML += `<option value="${{m}}">${{m}}</option>`);

['f-cert','f-status','f-market','f-date-from','f-date-to'].forEach(id => {{
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
  'total-assigned':      'Total number of people assigned this certification curriculum.',
  'curriculum-complete': 'People who have completed all required LMS coursework. This is the first step — manager-confirmed HC certification is tracked separately.',
  'certified':           'People who have received full Healthcare certification, confirmed by their manager. The final step after completing LMS coursework.',
  'not-certified':       'People assigned the curriculum who have not yet received manager-confirmed HC certification.',
  'completion-rate':     'Percentage of assigned people who have earned manager-confirmed HC certification.',
  'pipeline':            'Where people stand in the certification journey: Not Started (no LMS work), Curriculum Complete (LMS done, awaiting manager sign-off), HC Certified (fully certified).',
  'over-time':           'Sub-certifications earned per fiscal quarter, stacked by type. FY2026 Q1 = Apr-Jun, Q2 = Jul-Sep, Q3 = Oct-Dec, Q4 = Jan-Mar.',
  'subcert':             'How many people have completed each curriculum. Healthcare Foundational is required first — Acute Care, Ambulatory, and Extended Care build on top of it.',
  'roster':              'All assigned people with their certification status. Click a name to see details, sub-certification badges, and manager contact info.',
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
  ['f-cert','f-status','f-market','f-date-from','f-date-to'].forEach(id => sel(id).value = '');
  if(hideTLG){{ hideTLG=false; sel('btn-tlg').classList.remove('active'); sel('btn-tlg').textContent='Hide TLG'; }}
  applyFilters();
}}

function applyFilters(){{
  const cert   = sel('f-cert').value;
  const certField = cert || 'Healthcare';
  const status = sel('f-status').value;
  const market = sel('f-market').value;
  const from   = sel('f-date-from').value;
  const to     = sel('f-date-to').value;
  const effectiveStatus = (cert && !status) ? 'Yes' : status;
  filtered = RAW.filter(r => {{
    if(hideTLG && TLG_SET.has(r.FirstName+' '+r.LastName)) return false;
    if(effectiveStatus && r[certField] !== effectiveStatus) return false;
    if(market && r.Market !== market) return false;
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
  const certField  = sel('f-cert').value || 'Healthcare';

  // Stat cards
  const total          = filtered.length;
  const curriculumDone = filtered.filter(r=>r.Complete==='Yes').length;
  const certified      = filtered.filter(r=>r[certField]==='Yes').length;
  const notCert        = total - certified;
  const rate           = total > 0 ? Math.round(certified/total*100) : 0;

  sel('s-total').textContent          = total;
  sel('s-curriculum').textContent     = curriculumDone;
  sel('s-curriculum-sub').textContent = total > 0 ? `${{curriculumDone}} of ${{total}} assigned` : '';
  sel('s-certified').textContent      = certified;
  sel('s-not').textContent            = notCert;
  sel('s-rate').textContent           = rate + '%';
  sel('s-rate-sub').textContent       = total > 0 ? `${{certified}} of ${{total}} assigned` : '';

  // Pipeline chart — three stages of the HC certification journey
  const pipelineNotStarted   = filtered.filter(r=>r.Complete!=='Yes'&&r.Healthcare!=='Yes').length;
  const pipelineCurrComplete = filtered.filter(r=>r.Complete==='Yes'&&r.Healthcare!=='Yes').length;
  const pipelineHCCertified  = filtered.filter(r=>r.Healthcare==='Yes').length;

  if(regionChart) regionChart.destroy();
  regionChart = new Chart(sel('regionChart'), {{
    type: 'bar',
    data: {{
      labels: ['Not Started','Curriculum Complete','HC Certified'],
      datasets: [{{
        data: [pipelineNotStarted, pipelineCurrComplete, pipelineHCCertified],
        backgroundColor: [cv('--red')+'cc', cv('--accent')+'cc', cv('--green')+'cc'],
        borderRadius: 4,
        borderSkipped: false,
      }}]
    }},
    options: {{
      indexAxis: 'y',
      responsive:true, maintainAspectRatio:false,
      plugins:{{
        legend:{{display:false}},
        tooltip:{{
          callbacks:{{
            label: ctx => ` ${{ctx.raw}} person${{ctx.raw!==1?'s':''}}`
          }}
        }},
        datalabels:{{display:false}}
      }},
      scales:{{
        x:{{grid:{{color:cv('--border')}},ticks:{{color:chartLabel,font:{{size:11}},stepSize:1}}}},
        y:{{grid:{{display:false}},ticks:{{color:chartLabel,font:{{size:12}}}}}}
      }}
    }}
  }});

  // Trend chart — stacked bars by sub-cert per fiscal quarter
  const TREND_SUBCERTS = [
    {{key:'Healthcare', qtrKey:'HCQtr',    label:'Healthcare Foundational', color:cv('--green')}},
    {{key:'AcuteCare',  qtrKey:'AcuteQtr', label:'Acute Care',    color:cv('--accent3')}},
    {{key:'Ambulatory', qtrKey:'AmbQtr',   label:'Ambulatory',    color:cv('--accent2')}},
    {{key:'Extended',   qtrKey:'ExtQtr',   label:'Extended Care', color:cv('--teal')}},
  ];
  const parseQtr = s => {{ const m=s.match(/FY(\d+)\s+Q(\d)/); return m ? parseInt(m[1])*10+parseInt(m[2]) : 0; }};
  const allTrendQtrs = new Set();
  TREND_SUBCERTS.forEach(sc => {{
    filtered.forEach(r => {{ if(r[sc.key]==='Yes' && r[sc.qtrKey]) allTrendQtrs.add(r[sc.qtrKey]); }});
  }});
  const trendQtrs = [...allTrendQtrs].sort((a,b)=>parseQtr(a)-parseQtr(b));
  const trendDatasets = TREND_SUBCERTS.map(sc => ({{
    label: sc.label,
    data:  trendQtrs.map(q => filtered.filter(r=>r[sc.key]==='Yes'&&r[sc.qtrKey]===q).length),
    backgroundColor: sc.color+'bb',
    borderRadius: 3,
    borderSkipped: false,
  }}));

  if(trendChart) trendChart.destroy();
  trendChart = new Chart(sel('trendChart'), {{
    type: 'bar',
    data: {{ labels:trendQtrs, datasets:trendDatasets }},
    options: {{
      responsive:true, maintainAspectRatio:false,
      plugins:{{
        legend:{{position:'bottom',labels:{{color:chartLabel,font:{{size:10}},boxWidth:8,padding:6}}}},
        tooltip:{{mode:'index',intersect:false}},
        datalabels:{{display:false}}
      }},
      scales:{{
        x:{{grid:{{color:cv('--border')}},ticks:{{color:chartLabel,font:{{size:10}},maxRotation:45}},stacked:true}},
        y:{{grid:{{color:cv('--border')}},ticks:{{color:chartLabel,font:{{size:11}},stepSize:1}},stacked:true}}
      }}
    }}
  }});

  // Sub-cert breakdown chart
  const subLabels = ['Healthcare Foundational','Acute Care','Ambulatory','Extended Care'];
  const subKeys   = ['Healthcare','AcuteCare','Ambulatory','Extended'];
  const subColors = [cv('--green'),cv('--accent3'),cv('--accent2'),cv('--teal')];
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

  renderRosterList();
}}

function subCertCount(r){{
  return ['Healthcare','AcuteCare','Ambulatory','Extended'].filter(k=>r[k]==='Yes').length;
}}

function setRosterSort(field){{
  if(rosterSortField === field) {{
    rosterSortDir = rosterSortDir === 'desc' ? 'asc' : 'desc';
  }} else {{
    rosterSortField = field;
    rosterSortDir   = 'desc';
  }}
  renderRosterList();
}}

function pipelineSteps(p) {{
  const bothDone = p.Healthcare==='Yes';
  const lmsOnly  = !bothDone && p.Complete==='Yes';
  const amber    = '#f59e0b';
  const blue     = cv('--accent');
  function dot(color, tip, check) {{
    const bg = check ? color+'33' : 'var(--surface2)';
    const border = check ? color : 'var(--border)';
    const txt    = check ? color : 'var(--muted)';
    return `<div title="${{tip}}" style="width:16px;height:16px;border-radius:50%;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:9px;font-weight:700;background:${{bg}};border:1.5px solid ${{border}};color:${{txt}}">${{check?'&#10003;':''}}</div>`;
  }}
  function miniDot(done, tip) {{
    return `<div title="${{tip}}" style="width:10px;height:10px;border-radius:50%;flex-shrink:0;background:${{done?blue+'33':'var(--surface2)'}};border:1.5px solid ${{done?blue:'var(--border)'}}"></div>`;
  }}
  const lmsColor = bothDone ? cv('--green') : lmsOnly ? amber : 'var(--border)';
  const lmsCheck = p.Complete==='Yes';
  const hcColor  = cv('--green');
  const hcCheck  = p.Healthcare==='Yes';
  const secDots  = [
    miniDot(p.AcuteCare==='Yes',  'Acute Care'),
    miniDot(p.Ambulatory==='Yes', 'Ambulatory'),
    miniDot(p.Extended==='Yes',   'Extended Care'),
  ].join('');
  return `<div style="display:flex;align-items:center;gap:3px;flex-shrink:0;">
    ${{dot(lmsColor,'LMS Complete',lmsCheck)}}
    <div style="width:8px;height:1px;background:var(--border);flex-shrink:0;"></div>
    ${{dot(hcColor,'HC Certified',hcCheck)}}
    <div style="width:8px;height:1px;background:var(--border);flex-shrink:0;"></div>
    <div style="display:flex;align-items:center;gap:2px;">${{secDots}}</div>
  </div>`;
}}

function tierStyle(p, extraIndent) {{
  extraIndent = extraIndent || 0;
  const base = 14 + extraIndent;
  const green = cv('--green');
  const amber = '#f59e0b';
  if(p.Healthcare==='Yes') return `border-left:3px solid ${{green}};padding-left:${{base-3}}px;`;
  if(p.Complete==='Yes')   return `border-left:3px solid ${{amber}};padding-left:${{base-3}}px;`;
  return extraIndent ? `padding-left:${{base}}px;` : '';
}}

function setRosterView(v) {{
  rosterView = v;
  sel('view-individual').classList.toggle('active', v==='individual');
  sel('view-manager').classList.toggle('active', v==='manager');
  sel('sort-controls').style.display = v==='individual' ? 'flex' : 'none';
  sel('roster-right').innerHTML = '<div class="no-data">Select a person to see their details</div>';
  renderRosterList();
}}

function toggleMgrGroup(el) {{
  el.classList.toggle('open');
  el.nextElementSibling.classList.toggle('open');
}}

function renderRosterList(){{
  const certField = sel('f-cert').value || 'Healthcare';

  if(rosterView === 'manager') {{
    const groups = {{}};
    filtered.forEach(p => {{
      const mgr = p.Manager || 'No Manager';
      if(!groups[mgr]) groups[mgr] = [];
      groups[mgr].push(p);
    }});
    const mgrsSorted = Object.keys(groups).sort((a,b) => groups[b].length - groups[a].length);
    sel('roster-left').innerHTML = mgrsSorted.map(mgr => {{
      const team = groups[mgr];
      const cert = team.filter(p=>p.Healthcare==='Yes').length;
      const lms  = team.filter(p=>p.Complete==='Yes').length;
      const peopleHtml = team
        .sort((a,b)=>(a.LastName+a.FirstName).localeCompare(b.LastName+b.FirstName))
        .map(p => {{
          const fullName = `${{p.FirstName}} ${{p.LastName}}`;
          return `<div class="roster-person" onclick="rosterSelect(this)" data-name="${{fullName}}" style="${{tierStyle(p,10)}}">
            <span class="roster-name">${{fullName}}</span>
            ${{pipelineSteps(p)}}
          </div>`;
        }}).join('');
      return `<div>
        <div class="mgr-group-hdr open" onclick="toggleMgrGroup(this)">
          <div>
            <div style="font-size:12px;font-weight:600;color:var(--text)">${{mgr}}</div>
            <div style="font-size:11px;color:var(--muted);margin-top:2px;">${{team.length}} rep${{team.length!==1?'s':''}} &middot; ${{lms}} LMS &middot; ${{cert}} certified</div>
          </div>
          <span class="mgr-chevron">&#9660;</span>
        </div>
        <div class="mgr-team open">${{peopleHtml}}</div>
      </div>`;
    }}).join('') || `<div class="no-data">No people match filters</div>`;
    filterRoster();
    return;
  }}

  // Individual view
  const d = rosterSortDir === 'desc' ? -1 : 1;
  const sorted = [...filtered].sort((a,b) => {{
    if(rosterSortField === 'name') {{
      return d * (a.LastName+a.FirstName).localeCompare(b.LastName+b.FirstName);
    }} else if(rosterSortField === 'market') {{
      const rc = d * (a.Market||'').localeCompare(b.Market||'');
      return rc !== 0 ? rc : (a.LastName+a.FirstName).localeCompare(b.LastName+b.FirstName);
    }} else {{
      function score(p) {{ return p[certField]==='Yes' ? 2 : p.Complete==='Yes' ? 1 : 0; }}
      const diff = score(a) - score(b);
      if(diff !== 0) return d * diff;
      return subCertCount(b) - subCertCount(a);
    }}
  }});

  document.querySelectorAll('.sort-btn').forEach(btn => {{
    if(!btn.dataset.sort) return;
    const active = btn.dataset.sort === rosterSortField;
    btn.classList.toggle('active', active);
    const arrow = active ? (rosterSortDir === 'desc' ? ' ↓' : ' ↑') : '';
    btn.textContent = btn.dataset.label + arrow;
  }});

  sel('roster-left').innerHTML = sorted.map(p => {{
    const fullName = `${{p.FirstName}} ${{p.LastName}}`;
    return `<div class="roster-person" onclick="rosterSelect(this)" data-name="${{fullName}}" style="${{tierStyle(p)}}">
      <span class="roster-name">${{fullName}}</span>
      ${{pipelineSteps(p)}}
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
  const isCert = p.Healthcare==='Yes';
  function fmtDate(d) {{
    if (!d) return '&#8212;';
    const [y,mo,dy] = d.split('-');
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return months[parseInt(mo)-1] + ' ' + parseInt(dy) + ', ' + y;
  }}
  const SUBCERT_DEFS = [
    {{key:'Healthcare', dateKey:'HCDate',    qtrKey:'HCQtr',    label:'Healthcare Foundational'}},
    {{key:'AcuteCare',  dateKey:'AcuteDate', qtrKey:'AcuteQtr', label:'Acute Care'}},
    {{key:'Ambulatory', dateKey:'AmbDate',   qtrKey:'AmbQtr',   label:'Ambulatory'}},
    {{key:'Extended',   dateKey:'ExtDate',   qtrKey:'ExtQtr',   label:'Extended Care'}},
  ];
  sel('roster-right').innerHTML = `
    <div class="roster-right-header">
      <div style="font-size:15px;font-weight:700;margin-bottom:6px">${{p.FirstName}} ${{p.LastName}}</div>
      <span class="badge-status ${{isCert?'certified':'not-certified'}}">${{isCert?'HC Certified':'Not Yet Certified'}}</span>
      ${{isCert && p.HCDate ? `<span style="font-size:12px;color:var(--muted);margin-left:8px">${{fmtDate(p.HCDate)}}</span>` : ''}}
    </div>
    <div class="detail-grid">
      <div><div class="detail-label">Job Title</div><div class="detail-value">${{p.JobTitle||'&#8212;'}}</div></div>
      <div><div class="detail-label">Market</div><div class="detail-value">${{p.Market||'&#8212;'}}</div></div>
      <div><div class="detail-label">Healthcare Foundational</div><div class="detail-value">${{p.Complete==='Yes' ? '&#10003; LMS Complete' : 'Not Complete'}}</div></div>
      <div><div class="detail-label">Hire Date</div><div class="detail-value">${{fmtDate(p.HireDate)}}</div></div>
      <div style="grid-column:1/-1"><div class="detail-label">Email</div><div class="detail-value"><a href="mailto:${{p.Email}}" style="color:var(--accent);text-decoration:none">${{p.Email||'&#8212;'}}</a></div></div>
    </div>
    <div style="margin-top:14px;padding-top:12px;border-top:1px solid var(--border)">
      <div style="font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px">Secondary Curriculums</div>
      <div style="display:flex;flex-direction:column;gap:8px">
        ${{SUBCERT_DEFS.filter(sc=>sc.key!=='Healthcare').map(sc => {{
          const yes = p[sc.key]==='Yes';
          const dateStr = yes && p[sc.dateKey] ? fmtDate(p[sc.dateKey]) + (p[sc.qtrKey] ? ' &middot; ' + p[sc.qtrKey] : '') : '';
          return `<div style="display:flex;align-items:center;gap:12px">
            <span style="font-size:12px;font-weight:600;padding:4px 12px;border-radius:20px;min-width:140px;text-align:center;background:${{yes?'var(--green-subtle)':'var(--surface2)'}};color:${{yes?'var(--green)':'var(--muted)'}};border:1px solid ${{yes?'var(--green)':'var(--border)'}}">${{yes?'&#10003;':'&#8212;'}} ${{sc.label}}</span>
            ${{dateStr ? `<span style="font-size:12px;color:var(--muted)">${{dateStr}}</span>` : ''}}
          </div>`;
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
            print(f'    {len(rows)} rows  ({sum(1 for r in rows if r["Healthcare"]=="Yes")} HC certified, {sum(1 for r in rows if r["Complete"]=="Yes")} curriculum complete)')

        # Deduplicate: later files overwrite earlier ones for the same person
        all_rows.sort(key=lambda r: r['_file_date'])
        seen = {}
        for r in all_rows:
            seen[person_key(r)] = r
        deduped = list(seen.values())
        for r in deduped:
            del r['_file_date']

        hc_certified  = sum(1 for r in deduped if r['Healthcare'] == 'Yes')
        curr_complete = sum(1 for r in deduped if r['Complete'] == 'Yes')
        print(f'  → {len(deduped)} unique people  ({hc_certified} HC certified, {curr_complete} curriculum complete, {len(deduped)-hc_certified} not yet)')

        html = generate_html(slug, vert_name, deduped)
        out  = f'cert-{slug}.html'
        with open(out, 'w', encoding='utf-8') as fh:
            fh.write(html)
        print(f'  Written → {out}')
        generated.append(out)

    print(f'\nDone. {len(generated)} dashboard(s) generated.')


if __name__ == '__main__':
    main()
