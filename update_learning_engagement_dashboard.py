"""
Learning Engagement Dashboard generator.

Reads a LinkedIn Learning engagement dataset from learning-engagement-data/ and
produces learning-engagement.html, following the same house style/architecture as
every other dashboard in this repo: KM Academy logo, "Data through" header,
triple-click-to-home nav (on the <h1>), ? info tooltips, light/dark theme
(shared pb-theme localStorage key), PDF/Excel export dropdown, and a real
Hide/Show TLG toggle button (not silent exclusion).

DATA SOURCE STATUS (first build, 2026-09-17): Resmie has not yet delivered the
real source files (a LinkedIn Learning learner-detail export + an active-
employee roster export). The loader below reads a JSON file matching the exact
per-record schema her own prototype (`learning_engagement_dashboard (1).html`)
already used:
    {"total_active": <int>,
     "records": [{"u","c","m","d","a","ct","cn","h","p","sk","mo","dt"}, ...]}
A stand-in copy of her prototype's own embedded data has been placed at
learning-engagement-data/learning_engagement_data.json for local testing only —
NOT real data, just a way to prove this script produces the same look/output
her prototype does. If her real deliverable turns out to be raw Excel exports
instead of this JSON shape, this loader's file-detection/parsing will need to
be rebuilt against the real column layout at that time — same caution as every
other LMS loader in this repo (never guess column positions blind).

Known gap, flagged not fixed: Resmie's data has no name field, only an
anonymized-looking employee ID ("u"). Every other dashboard's TLG toggle
matches by name against the hardcoded 19-person TLG list. TLG_IDS below starts
empty because there's no ID mapping yet — the button is fully wired and
defaults to "hide," but filters nobody until Resmie's real file adds either a
name field or a confirmed employee-ID mapping for the 19 TLG members.
"""
import json
import os
import glob
import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(SCRIPT_DIR, 'learning-engagement-data')

MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
          'July', 'August', 'September', 'October', 'November', 'December']

# See "Known gap" note above — populate with real employee IDs once Resmie
# confirms them for the 19 TLG members.
TLG_IDS = set([])


def load_learning_engagement_data():
    """Find the current data file in learning-engagement-data/ and return
    (records, total_active, date_label)."""
    candidates = sorted(glob.glob(os.path.join(DATA_DIR, '*.json')))
    if not candidates:
        raise FileNotFoundError(
            f"No data file found in {DATA_DIR}. Drop a learning_engagement_data.json "
            "(or Resmie's real export, once its format is known/converted) there and rerun."
        )
    path = candidates[-1]
    with open(path, encoding='utf-8') as f:
        payload = json.load(f)
    records = payload['records']
    total_active = payload.get('total_active', 0)

    dates = [r['dt'] for r in records if r.get('dt')]
    if dates:
        latest = max(dates)
        y, m, d = latest.split('-')
        date_label = f"{MONTHS[int(m) - 1]} {int(d)}, {y}"
    else:
        date_label = datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime('%B %-d, %Y')

    return records, total_active, date_label


def generate_html(records, total_active, date_label):
    data_json = json.dumps(records, separators=(',', ':'))
    tlg_ids_json = json.dumps(sorted(TLG_IDS))

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Learning Engagement Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
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
  *{{box-sizing:border-box;margin:0;padding:0;}}
  body{{background:var(--bg);color:var(--text);font-family:var(--font);min-height:100vh;transition:background .2s,color .2s;}}

  /* ── Header ─────────────────────────────────────────────────────── */
  .header{{padding:20px 28px 16px;border-bottom:1px solid var(--border);display:grid;grid-template-columns:1fr auto 1fr;align-items:center;gap:12px;}}
  .header-center{{display:flex;justify-content:center;align-items:center;}}
  .kma-logo{{height:38px;width:auto;display:block;}}
  .kma-logo-light{{display:none;}}
  .light-mode .kma-logo-dark{{display:none;}}
  .light-mode .kma-logo-light{{display:block;}}
  .header h1{{font-size:18px;font-weight:700;letter-spacing:.3px;cursor:pointer;}}
  .header h1 span{{color:var(--muted);font-weight:400;}}
  .header-date{{font-size:11px;color:var(--muted);margin-top:4px;}}
  .btn-theme{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:5px 12px;font-size:12px;cursor:pointer;transition:all .15s;}}
  .btn-theme:hover{{border-color:var(--accent);color:var(--text);}}

  /* ── Filters ────────────────────────────────────────────────────── */
  .filters{{padding:14px 28px;display:flex;gap:10px;flex-wrap:wrap;align-items:center;border-bottom:1px solid var(--border);background:var(--surface);}}
  .filter-label{{font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.6px;margin-right:4px;}}
  .sort-btn{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:5px 14px;font-size:12px;cursor:pointer;transition:all .15s;white-space:nowrap;}}
  .sort-btn:hover{{border-color:var(--accent);color:var(--text);}}
  .sort-btn.active{{border-color:var(--accent);color:var(--accent);background:var(--accent)11;}}
  .sort-btn.active[data-cat="BCA"]{{border-color:var(--teal);color:var(--teal);background:var(--teal)11;}}
  .sort-btn.active[data-cat="All"]{{border-color:var(--accent2);color:var(--accent2);background:var(--accent2)11;}}
  .btn-tlg{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:5px 14px;font-size:12px;cursor:pointer;transition:all .15s;}}
  .btn-tlg:hover{{border-color:var(--red);color:var(--red);}}
  .btn-tlg.active{{background:#f76f6f22;border-color:var(--red);color:var(--red);}}
  .result-count{{margin-left:auto;font-size:12px;color:var(--muted);}}

  /* ── Info tooltip ───────────────────────────────────────────────── */
  .info-btn{{display:inline-flex;align-items:center;justify-content:center;width:15px;height:15px;border-radius:50%;background:var(--surface2);border:1px solid var(--border);color:var(--muted);font-size:9px;font-weight:700;cursor:pointer;margin-left:5px;vertical-align:middle;flex-shrink:0;line-height:1;transition:border-color .15s,color .15s;}}
  .info-btn:hover{{border-color:var(--accent);color:var(--accent);}}
  .info-popover{{position:fixed;z-index:9999;background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:12px 14px;font-size:12px;color:var(--text);line-height:1.6;max-width:260px;box-shadow:0 4px 24px rgba(0,0,0,0.5);display:none;}}
  .info-popover.visible{{display:block;}}

  /* ── Export dropdown ───────────────────────────────────────────── */
  .btn-export{{background:var(--accent);border:1px solid var(--accent);color:#fff;border-radius:6px;padding:5px 12px;font-size:12px;cursor:pointer;transition:all .15s;font-weight:600;}}
  .btn-export:hover{{opacity:0.88;}}
  .export-drop{{position:relative;}}
  .export-menu{{position:absolute;top:calc(100% + 6px);right:0;background:var(--surface);border:1px solid var(--border);border-radius:8px;min-width:150px;box-shadow:0 4px 24px rgba(0,0,0,0.28);display:none;z-index:200;overflow:hidden;}}
  .export-menu.open{{display:block;}}
  .export-item{{display:block;width:100%;text-align:left;padding:10px 14px;font-size:13px;color:var(--text);background:transparent;border:none;cursor:pointer;transition:background .1s;}}
  .export-item:hover{{background:var(--surface2);}}

  /* ── Engagement / Learners overview ────────────────────────────── */
  .engagement-section{{padding:18px 28px 4px;}}
  .eng-card{{background:var(--surface);border:1px solid var(--border);border-radius:10px;display:grid;grid-template-columns:1fr 1px 1fr 1px 1fr;}}
  .eng-tile{{padding:20px 24px;min-width:0;}}
  .eng-divider{{background:var(--border);}}
  .eng-label{{font-size:12px;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);margin-bottom:8px;}}
  .eng-value{{font-size:34px;font-weight:700;line-height:1;color:var(--text);}}
  .eng-value.accent{{color:var(--accent);}}
  .eng-desc{{font-size:12px;color:var(--muted);margin-top:8px;line-height:1.5;max-width:440px;}}
  @media(max-width:920px){{.eng-card{{grid-template-columns:1fr;}}.eng-divider{{height:1px;}}}}

  /* ── Stat cards ─────────────────────────────────────────────────── */
  .stats{{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;padding:20px 28px;}}
  .stat{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:16px 18px;}}
  .stat-label{{font-size:11px;text-transform:uppercase;letter-spacing:.7px;color:var(--muted);margin-bottom:6px;}}
  .stat-value{{font-size:28px;font-weight:700;line-height:1;}}
  .stat-value.blue{{color:var(--accent);}}
  .stat-value.teal{{color:var(--teal);}}
  .stat-value.purple{{color:var(--accent2);}}
  .stat-value.green{{color:var(--green);}}
  .stat-value sup{{font-size:13px;font-weight:600;color:var(--muted);margin-left:2px;}}
  .stat-sub{{font-size:11px;color:var(--muted);margin-top:4px;}}

  /* ── Compare card ───────────────────────────────────────────────── */
  .compare-wrap{{padding:0 28px 4px;}}
  .compare{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:18px;display:flex;align-items:center;gap:22px;flex-wrap:wrap;}}
  .compare-title{{font-size:13px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;min-width:170px;}}
  .cbar{{flex:1;min-width:220px;display:flex;height:30px;border-radius:6px;overflow:hidden;background:var(--surface2);}}
  .cseg{{display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;color:#0f1117;transition:width .5s cubic-bezier(.4,0,.2,1);white-space:nowrap;overflow:hidden;}}
  .cseg.bus-seg{{background:var(--accent);}}
  .cseg.bca-seg{{background:var(--teal);}}
  .clegend{{display:flex;gap:16px;font-size:12px;color:var(--muted);}}
  .clegend .dot{{display:inline-block;width:9px;height:9px;border-radius:2px;margin-right:6px;}}
  .clegend .dot.bus{{background:var(--accent);}}
  .clegend .dot.bca{{background:var(--teal);}}

  /* ── Chart cards ────────────────────────────────────────────────── */
  .charts{{display:grid;grid-template-columns:repeat(3, 1fr);gap:16px;padding:16px 28px 8px;}}
  @media(max-width:1080px){{.charts{{grid-template-columns:1fr 1fr;}}}}
  @media(max-width:820px){{.charts{{grid-template-columns:1fr;}}}}
  .chart-card{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:18px;}}
  .chart-card.full{{grid-column:1/-1;}}
  .chart-title{{font-size:13px;font-weight:600;margin-bottom:2px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;}}
  body.light-mode .chart-title{{color:var(--text);}}
  .chart-sub{{font-size:11px;color:var(--muted);margin-bottom:14px;}}
  .chart-wrap{{position:relative;height:240px;margin-top:8px;}}
  .chart-legend{{display:flex;gap:18px;font-size:11px;color:var(--muted);margin-top:10px;}}
  .chart-legend .dot{{display:inline-block;width:9px;height:9px;border-radius:2px;margin-right:6px;}}

  .barrow{{display:grid;grid-template-columns:150px 1fr 74px;align-items:center;gap:12px;padding:8px 0;border-top:1px solid var(--border);}}
  .barrow:first-of-type{{border-top:none;}}
  .barrow .rlabel{{font-size:13px;color:var(--text);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}}
  .barrow .rtrack{{position:relative;height:6px;background:var(--surface2);border-radius:4px;overflow:hidden;}}
  .barrow .rfill{{position:absolute;left:0;top:0;bottom:0;border-radius:4px;width:0%;transition:width .6s cubic-bezier(.4,0,.2,1),background .3s ease;}}
  .barrow .rval{{font-size:12px;color:var(--muted);text-align:right;white-space:nowrap;font-weight:600;}}
  .barrow .rval b{{color:var(--text);font-weight:700;}}

  .chart-foot{{margin-top:14px;padding-top:12px;border-top:1px solid var(--border);}}
  .toggle-link{{background:none;border:none;color:var(--muted);font-size:12px;font-family:var(--font);cursor:pointer;padding:0;text-decoration:underline;text-decoration-color:var(--border);text-underline-offset:3px;}}
  .toggle-link:hover{{color:var(--text);}}

  .section{{padding:16px 28px 32px;}}
  .section-hint{{font-size:11px;color:var(--muted);line-height:1.6;}}

  .print-hide{{}}
  #print-header,#print-summary{{display:none;}}
  @media print{{
    body{{background:#fff!important;color:#111!important;}}
    .header,.filters,.stats,.engagement-section,.compare-wrap,.charts,.section,.print-hide{{display:none!important;}}
    #print-header{{display:block!important;}}
    #print-summary{{display:block!important;}}
    table{{width:100%;border-collapse:collapse;font-size:11px;margin-bottom:16px;}}
    th{{background:#f0f4ff;color:#111;font-weight:700;padding:5px 8px;border:1px solid #ccc;text-align:left;}}
    td{{padding:5px 8px;border:1px solid #ddd;}}
  }}

  @media(max-width:520px){{
    .stats{{grid-template-columns:1fr 1fr;}}
  }}
</style>
</head>
<body>

  <div class="header">
    <div>
      <h1 id="dash-title">Learning Engagement <span>Dashboard</span></h1>
      <div class="header-date">Data through {date_label}</div>
    </div>
    <div class="header-center">
      <img src="https://jasonackerman1.github.io/playbook-dashboard/KMA-wht.svg" class="kma-logo kma-logo-dark" alt="KM Academy">
      <img src="https://jasonackerman1.github.io/playbook-dashboard/KMA-drk.svg" class="kma-logo kma-logo-light" alt="KM Academy">
    </div>
    <div style="display:flex;justify-content:flex-end;gap:8px;align-items:center;">
      <div class="export-drop print-hide" id="export-drop">
        <button class="btn-export" onclick="toggleExportDrop()">&#128438; Export &#9660;</button>
        <div class="export-menu" id="export-menu">
          <button class="export-item" onclick="runExport()">PDF</button>
          <button class="export-item" onclick="runExportXLSX()">Excel</button>
        </div>
      </div><span class="info-btn print-hide" onclick="showInfo(event,'export')">?</span>
      <button class="btn-theme print-hide" id="btn-theme" onclick="toggleTheme()">&#9728; Light</button>
    </div>
  </div>

  <div class="filters">
    <span class="filter-label">Group</span>
    <div id="filterSeg" style="display:flex;gap:6px;">
      <button class="sort-btn active" data-cat="All">All</button>
      <button class="sort-btn" data-cat="BUS">BUS</button>
      <button class="sort-btn" data-cat="BCA">BCA</button>
    </div>
    <button class="btn-tlg active" id="btn-tlg" onclick="toggleTLG()">Show TLG</button><span class="info-btn" onclick="showInfo(event,'hide-tlg')">?</span>
    <span class="result-count" id="resultCount">–</span>
  </div>

  <div class="engagement-section">
    <div class="eng-card">
      <div class="eng-tile">
        <div class="eng-label">Engagement by content type</div>
        <div id="contentTypeList" style="margin-top:6px;"></div>
      </div>
      <div class="eng-divider"></div>
      <div class="eng-tile">
        <div class="eng-label">Unique learners<span class="info-btn" onclick="showInfo(event,'unique-learners')">?</span></div>
        <div class="eng-value" id="learnersTotal">–</div>
        <div class="eng-desc">Distinct people who engaged with at least one piece of content — counted once each, no matter how many courses, articles, or paths they touched.</div>
      </div>
      <div class="eng-divider"></div>
      <div class="eng-tile">
        <div class="eng-label">Active workforce reach<span class="info-btn" onclick="showInfo(event,'reach')">?</span></div>
        <div class="eng-value accent" id="reachPct">–</div>
        <div class="eng-desc" id="reachDesc">Share of the active workforce with a recorded hire date (per the active-employee master roster) that engaged with LinkedIn Learning at all during this period.</div>
      </div>
    </div>
  </div>

  <div class="stats">
    <div class="stat">
      <div class="stat-label">Avg. engagements / learner<span class="info-btn" onclick="showInfo(event,'avg-eng')">?</span></div>
      <div class="stat-value blue" id="kpiAvgEng">–</div>
      <div class="stat-sub" id="kpiAvgEngSub">&nbsp;</div>
    </div>
    <div class="stat">
      <div class="stat-label">Avg. time spent / learner<span class="info-btn" onclick="showInfo(event,'avg-time')">?</span></div>
      <div class="stat-value teal" id="kpiAvgTime">–</div>
      <div class="stat-sub" id="kpiAvgTimeSub">&nbsp;</div>
    </div>
    <div class="stat">
      <div class="stat-label">Total hours viewed<span class="info-btn" onclick="showInfo(event,'total-hours')">?</span></div>
      <div class="stat-value purple" id="kpiTotalHours">–</div>
      <div class="stat-sub" id="kpiTotalHoursSub">&nbsp;</div>
    </div>
  </div>

  <div class="compare-wrap">
    <div class="compare">
      <div class="compare-title">BUS vs BCA<br>unique learners</div>
      <div class="cbar">
        <div class="cseg bus-seg" id="busSeg">BUS</div>
        <div class="cseg bca-seg" id="bcaSeg">BCA</div>
      </div>
      <div class="clegend">
        <span><span class="dot bus"></span><span id="busLegend">–</span></span>
        <span><span class="dot bca"></span><span id="bcaLegend">–</span></span>
      </div>
    </div>
  </div>

  <div class="charts">
    <div class="chart-card full">
      <div class="chart-title">Month-over-month engagement<span class="info-btn" onclick="showInfo(event,'trend-chart')">?</span></div>
      <div class="chart-sub">Unique learners and hours viewed, by month last viewed</div>
      <div class="chart-wrap"><canvas id="trendChart"></canvas></div>
      <div class="chart-legend">
        <span><span class="dot" id="trendDotLearners"></span>Unique learners</span>
        <span><span class="dot" id="trendDotHours"></span>Hours viewed</span>
      </div>
    </div>

    <div class="chart-card">
      <div class="chart-title">Most popular topics</div>
      <div class="chart-sub">Ranked by number of unique learners engaging with the skill tag</div>
      <div id="topicsList"></div>
      <div class="chart-foot"><button class="toggle-link" id="topicsToggle">Show more topics</button></div>
    </div>

    <div class="chart-card">
      <div class="chart-title">Market engagement</div>
      <div class="chart-sub">Ranked by number of unique learners per market</div>
      <div id="marketsList"></div>
      <div class="chart-foot"><button class="toggle-link" id="marketsToggle">Show all markets</button></div>
    </div>

    <div class="chart-card">
      <div class="chart-title">Department engagement</div>
      <div class="chart-sub">Ranked by number of unique learners per department code</div>
      <div id="deptsList"></div>
      <div class="chart-foot"><button class="toggle-link" id="deptsToggle">Show all departments</button></div>
    </div>
  </div>

  <div class="section">
    <div class="section-hint" id="footerNote"></div>
  </div>

  <div id="info-popover" class="info-popover"></div>

  <div id="print-header">
    <div style="font-size:20px;font-weight:700;margin-bottom:4px;" id="ph-title">Learning Engagement Report</div>
    <div style="font-size:12px;color:#555;margin-bottom:2px;" id="ph-date"></div>
    <div style="font-size:12px;color:#555;margin-bottom:10px;" id="ph-filters"></div>
  </div>
  <div id="print-summary"></div>

<script>
const DATA = {data_json};
const TOTAL_ACTIVE = {total_active};
const TLG_IDS = new Set({tlg_ids_json});
</script>
<script>
(function(){{
  const accentSolid = {{All:'var(--accent2)', BUS:'var(--accent)', BCA:'var(--teal)'}};

  let topicsExpanded = false;
  let marketsExpanded = false;
  let deptsExpanded = false;
  let currentCat = 'All';
  let hideTLG = true;

  function sel(id){{ return document.getElementById(id); }}
  function cv(v){{ return getComputedStyle(document.body).getPropertyValue(v).trim(); }}

  // ── Theme (shared pb-theme convention) ──────────────────────────────────
  (function(){{
    if(localStorage.getItem('pb-theme') !== 'dark') document.body.classList.add('light-mode');
    sel('btn-theme').textContent = document.body.classList.contains('light-mode') ? '\\u{{1F319}} Dark' : '\\u2600 Light';
  }})();
  window.toggleTheme = function(){{
    const light = document.body.classList.toggle('light-mode');
    localStorage.setItem('pb-theme', light ? 'light' : 'dark');
    sel('btn-theme').textContent = light ? '\\u{{1F319}} Dark' : '\\u2600 Light';
    render();
  }};

  // ── Info tooltip ─────────────────────────────────────────────────────────
  document.addEventListener('click', function(e){{
    if(!e.target.classList.contains('info-btn')) sel('info-popover').classList.remove('visible');
  }});
  const INFO_MSGS = {{
    'export': 'Download a report based on whoever is currently shown on screen (respects the Group filter and Hide TLG toggle above). PDF opens a print-ready summary; Excel gives the same numbers as a spreadsheet.',
    'hide-tlg': 'Hides internal TLG/admin accounts from every stat, chart, and export on this page. Known gap: this data set only carries an employee ID, not a name, so this toggle currently has no one to exclude until a name or ID mapping for the TLG list is confirmed.',
    'unique-learners': 'Distinct people who engaged with at least one piece of content in this period, counted once each no matter how many items they touched.',
    'reach': 'Share of the active workforce (per the active-employee master roster) that engaged with LinkedIn Learning at all during this period.',
    'avg-eng': 'Total content items engaged with, divided by the number of unique learners.',
    'avg-time': 'Total hours viewed, divided by the number of unique learners.',
    'total-hours': 'Total hours of content viewed across every engagement in the current filtered view.',
    'trend-chart': 'Unique learners and total hours viewed, grouped by the month each engagement was last viewed.'
  }};
  window.showInfo = function(e, key){{
    const pop = sel('info-popover');
    pop.textContent = INFO_MSGS[key] || '';
    pop.classList.add('visible');
    const r = e.target.getBoundingClientRect();
    pop.style.top  = (r.bottom + 6) + 'px';
    pop.style.left = Math.min(r.left, window.innerWidth - 280) + 'px';
    e.stopPropagation();
  }};

  // ── Export dropdown ──────────────────────────────────────────────────────
  window.toggleExportDrop = function(){{ sel('export-menu').classList.toggle('open'); }};
  document.addEventListener('click', function(e){{
    const d = sel('export-drop');
    if(d && !d.contains(e.target)) sel('export-menu').classList.remove('open');
  }});

  // ── TLG toggle ────────────────────────────────────────────────────────────
  window.toggleTLG = function(){{
    hideTLG = !hideTLG;
    sel('btn-tlg').classList.toggle('active', hideTLG);
    sel('btn-tlg').textContent = hideTLG ? 'Show TLG' : 'Hide TLG';
    render();
  }};

  const uniq = arr => new Set(arr).size;

  function fmtHours(h){{
    if(h < 1) return Math.round(h*60) + 'm';
    const whole = Math.floor(h);
    const mins = Math.round((h - whole)*60);
    return mins ? whole + 'h ' + mins + 'm' : whole + 'h';
  }}
  function fmtNum(n){{ return n.toLocaleString('en-US'); }}

  function baseData(){{
    let rows = DATA.filter(r => r.a);
    if(hideTLG) rows = rows.filter(r => !TLG_IDS.has(r.u));
    return rows;
  }}

  function filterRows(cat){{
    const active = baseData();
    if(cat === 'All') return active;
    return active.filter(r => r.c === cat);
  }}

  function computeKPIs(rows){{
    const learners = uniq(rows.map(r=>r.u));
    const totalHours = rows.reduce((a,r)=>a+r.h,0);
    const avgPerLearner = learners ? totalHours/learners : 0;
    return {{learners, totalHours, avgPerLearner, engagements: rows.length}};
  }}

  function computeContentTypes(rows){{
    const map = new Map();
    rows.forEach(r=>{{
      const ct = r.ct || 'Other';
      if(!map.has(ct)) map.set(ct, {{count:0, users:new Set()}});
      const o = map.get(ct);
      o.count += 1;
      o.users.add(r.u);
    }});
    const arr = Array.from(map.entries()).map(([type,o])=>({{type, count:o.count, users:o.users.size}}));
    arr.sort((a,b)=> b.count - a.count);
    return arr;
  }}

  function computeTopics(rows){{
    const map = new Map();
    rows.forEach(r=>{{
      (r.sk||[]).forEach(s=>{{
        if(!map.has(s)) map.set(s, {{users:new Set(), hours:0}});
        const o = map.get(s);
        o.users.add(r.u);
        o.hours += r.h;
      }});
    }});
    const arr = Array.from(map.entries()).map(([skill,o])=>({{skill, users:o.users.size, hours:o.hours}}));
    arr.sort((a,b)=> b.users - a.users);
    return arr;
  }}

  const MONTH_NAMES = {{'01':'Jan','02':'Feb','03':'Mar','04':'Apr','05':'May','06':'Jun','07':'Jul','08':'Aug','09':'Sep','10':'Oct','11':'Nov','12':'Dec'}};
  function monthLabel(mo){{
    const [y,m] = mo.split('-');
    return MONTH_NAMES[m] + ' ' + y;
  }}

  function computeTrend(rows){{
    const map = new Map();
    rows.forEach(r=>{{
      if(!r.mo) return;
      if(!map.has(r.mo)) map.set(r.mo, {{users:new Set(), hours:0}});
      const o = map.get(r.mo);
      o.users.add(r.u);
      o.hours += r.h;
    }});
    const months = Array.from(map.keys()).sort();
    return {{
      labels: months.map(monthLabel),
      learners: months.map(mo=>map.get(mo).users.size),
      hours: months.map(mo=>Math.round(map.get(mo).hours))
    }};
  }}

  let trendChart = null;
  function drawTrend(rows){{
    const t = computeTrend(rows);
    const learnerColor = cv('--accent');
    const hoursColor = cv('--accent3');
    const gridColor = cv('--border');
    const textColor = cv('--muted');
    sel('trendDotLearners').style.background = learnerColor;
    sel('trendDotHours').style.background = hoursColor;

    const ctx = sel('trendChart').getContext('2d');
    if(trendChart) trendChart.destroy();
    trendChart = new Chart(ctx, {{
      type: 'line',
      data: {{
        labels: t.labels,
        datasets: [
          {{
            label: 'Unique learners', data: t.learners,
            borderColor: learnerColor, backgroundColor: learnerColor,
            yAxisID: 'y', tension: 0.3, pointRadius: 3, borderWidth: 2,
          }},
          {{
            label: 'Hours viewed', data: t.hours,
            borderColor: hoursColor, backgroundColor: hoursColor,
            yAxisID: 'y1', tension: 0.3, pointRadius: 3, borderWidth: 2, borderDash: [4,3],
          }}
        ]
      }},
      options: {{
        responsive: true, maintainAspectRatio: false,
        interaction: {{mode:'index', intersect:false}},
        plugins: {{ legend: {{ display:false }} }},
        scales: {{
          x: {{ grid:{{color:gridColor}}, ticks:{{color:textColor, font:{{size:11}}}} }},
          y: {{ position:'left', grid:{{color:gridColor}}, ticks:{{color:textColor, font:{{size:11}}}}, title:{{display:true, text:'Learners', color:textColor, font:{{size:11}}}} }},
          y1:{{ position:'right', grid:{{display:false}}, ticks:{{color:textColor, font:{{size:11}}}}, title:{{display:true, text:'Hours', color:textColor, font:{{size:11}}}} }},
        }}
      }}
    }});
  }}

  function computeMarkets(rows){{
    const map = new Map();
    rows.forEach(r=>{{
      if(!map.has(r.m)) map.set(r.m, {{users:new Set(), hours:0}});
      const o = map.get(r.m);
      o.users.add(r.u);
      o.hours += r.h;
    }});
    const arr = Array.from(map.entries()).map(([market,o])=>({{market, users:o.users.size, hours:o.hours}}));
    arr.sort((a,b)=> b.users - a.users);
    return arr;
  }}

  function computeDepartments(rows){{
    const map = new Map();
    rows.forEach(r=>{{
      if(!map.has(r.d)) map.set(r.d, {{users:new Set(), hours:0}});
      const o = map.get(r.d);
      o.users.add(r.u);
      o.hours += r.h;
    }});
    const arr = Array.from(map.entries()).map(([dept,o])=>({{dept, users:o.users.size, hours:o.hours}}));
    arr.sort((a,b)=> b.users - a.users);
    return arr;
  }}

  function renderBarList(container, items, limit, expanded, valueKey, labelKey, color){{
    const list = expanded ? items : items.slice(0, limit);
    const max = items.length ? items[0][valueKey] : 1;
    container.innerHTML = list.map(it => {{
      const pct = max ? Math.max(4, Math.round(it[valueKey]/max*100)) : 0;
      const label = it[labelKey].length > 26 ? it[labelKey].slice(0,25)+'\\u2026' : it[labelKey];
      return `<div class="barrow">
        <div class="rlabel" title="${{it[labelKey].replace(/"/g,'&quot;')}}">${{label}}</div>
        <div class="rtrack"><div class="rfill" style="width:${{pct}}%; background:${{color}};"></div></div>
        <div class="rval"><b>${{fmtNum(it[valueKey])}}</b></div>
      </div>`;
    }}).join('');
  }}

  function render(){{
    const rows = filterRows(currentCat);
    const k = computeKPIs(rows);
    const col = accentSolid[currentCat];

    sel('learnersTotal').textContent = fmtNum(k.learners);

    const reachPct = TOTAL_ACTIVE ? (k.learners / TOTAL_ACTIVE * 100) : 0;
    sel('reachPct').textContent = reachPct.toFixed(1) + '%';
    sel('reachDesc').textContent = currentCat === 'All'
      ? `${{fmtNum(k.learners)}} of ${{fmtNum(TOTAL_ACTIVE)}} active employees engaged with LinkedIn Learning at all during this period.`
      : `${{fmtNum(k.learners)}} ${{currentCat}}-tagged learners engaged, out of ${{fmtNum(TOTAL_ACTIVE)}} total active employees org-wide — the master roster doesn't split by BUS/BCA, so this is measured against the whole workforce, not just ${{currentCat}} headcount.`;

    sel('kpiAvgEng').textContent = k.learners ? (k.engagements/k.learners).toFixed(1) : '0';
    sel('kpiAvgEngSub').textContent = 'content items per learner';

    sel('kpiAvgTime').textContent = fmtHours(k.avgPerLearner);
    sel('kpiAvgTimeSub').textContent = 'per learner across the period';

    sel('kpiTotalHours').innerHTML = Math.round(k.totalHours).toLocaleString() + '<sup>hrs</sup>';
    sel('kpiTotalHoursSub').textContent = fmtHours(k.totalHours) + ' viewed total';

    sel('resultCount').textContent = fmtNum(k.engagements) + ' engagements \\u00b7 ' + fmtNum(k.learners) + ' learners';

    const ctData = computeContentTypes(rows);
    renderBarList(sel('contentTypeList'), ctData, ctData.length, true, 'count', 'type', col);

    drawTrend(rows);

    const topicData = computeTopics(rows);
    renderBarList(sel('topicsList'), topicData, 8, topicsExpanded, 'users', 'skill', col);
    sel('topicsToggle').textContent = topicsExpanded ? 'Show fewer topics' : `Show more topics (${{topicData.length - 8}} more)`;
    sel('topicsToggle').style.display = topicData.length > 8 ? 'inline' : 'none';

    const marketData = computeMarkets(rows);
    renderBarList(sel('marketsList'), marketData, 8, marketsExpanded, 'users', 'market', col);
    sel('marketsToggle').textContent = marketsExpanded ? 'Show fewer markets' : `Show all markets (${{marketData.length}})`;
    sel('marketsToggle').style.display = marketData.length > 8 ? 'inline' : 'none';

    const deptData = computeDepartments(rows);
    renderBarList(sel('deptsList'), deptData, 8, deptsExpanded, 'users', 'dept', col);
    sel('deptsToggle').textContent = deptsExpanded ? 'Show fewer departments' : `Show all departments (${{deptData.length}})`;
    sel('deptsToggle').style.display = deptData.length > 8 ? 'inline' : 'none';

    const activeAll = baseData();
    const busUsers = uniq(activeAll.filter(r=>r.c==='BUS').map(r=>r.u));
    const bcaUsers = uniq(activeAll.filter(r=>r.c==='BCA').map(r=>r.u));
    const totalCmp = busUsers + bcaUsers;
    const busPct = totalCmp ? busUsers/totalCmp*100 : 0;
    const bcaPct = totalCmp ? bcaUsers/totalCmp*100 : 0;
    sel('busSeg').style.width = busPct + '%';
    sel('bcaSeg').style.width = bcaPct + '%';
    sel('busSeg').textContent = busPct > 12 ? busUsers.toLocaleString() : '';
    sel('bcaSeg').textContent = bcaPct > 12 ? bcaUsers.toLocaleString() : '';
    sel('busLegend').textContent = 'BUS \\u2014 ' + busUsers.toLocaleString() + ' learners';
    sel('bcaLegend').textContent = 'BCA \\u2014 ' + bcaUsers.toLocaleString() + ' learners';

    sel('footerNote').innerHTML =
      `Data sources: LinkedIn Learning learner detail export joined to the active-employee roster on employee ID / email to resolve market, department, and active status. ` +
      `This dashboard (${{activeAll.length.toLocaleString()}} engagement records, ${{uniq(activeAll.map(r=>r.u))}} learners) is limited to people confirmed in the active-employee roster. ` +
      `Records without a resolvable market or department are grouped as "Unmatched". A small number of learners (${{uniq(activeAll.filter(r=>r.c==='Unknown').map(r=>r.u))}}) had no BUS/BCA group tag on record and are excluded from the group filter and compare strip.`;
  }}

  // ── Export: PDF ───────────────────────────────────────────────────────────
  window.runExport = function(){{
    sel('export-menu').classList.remove('open');
    const rows = filterRows(currentCat);
    const k = computeKPIs(rows);
    const dateStr = new Date().toLocaleDateString('en-US',{{year:'numeric',month:'long',day:'numeric'}});
    const parts = [];
    if(currentCat !== 'All') parts.push('Group: ' + currentCat);
    if(hideTLG) parts.push('TLG Hidden');
    sel('ph-date').textContent = 'Generated: ' + dateStr + '  |  ' + fmtNum(k.engagements) + ' engagements \\u00b7 ' + fmtNum(k.learners) + ' learners';
    sel('ph-filters').textContent = parts.length ? parts.join('  |  ') : 'No filters active \\u2014 showing all data';

    const ctData = computeContentTypes(rows);
    const topicData = computeTopics(rows);
    const marketData = computeMarkets(rows);
    const deptData = computeDepartments(rows);
    const miniTh = 'style="text-align:left;padding:3px 6px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:#888;border-bottom:1px solid #ddd;"';
    const miniTd = 'style="padding:3px 6px;font-size:11px;"';
    const miniTdR = 'style="padding:3px 6px;font-size:11px;text-align:right;"';

    function table(title, rows_, cols){{
      return `<div style="margin-bottom:16px;">
        <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#333;margin-bottom:6px;">${{title}}</div>
        <table><tr>${{cols.map(c=>`<th ${{miniTh}}>${{c}}</th>`).join('')}}</tr>${{rows_}}</table>
      </div>`;
    }}

    sel('print-summary').innerHTML =
      `<div style="display:flex;gap:24px;flex-wrap:wrap;padding:10px 0 14px;border-bottom:2px solid #ddd;margin-bottom:14px;">
        <div><div style="font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:#777;">Unique Learners</div><div style="font-size:24px;font-weight:700;color:#4f8ef7;">${{fmtNum(k.learners)}}</div></div>
        <div><div style="font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:#777;">Avg Engagements / Learner</div><div style="font-size:24px;font-weight:700;color:#7c5cfc;">${{k.learners?(k.engagements/k.learners).toFixed(1):'0'}}</div></div>
        <div><div style="font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:#777;">Total Hours Viewed</div><div style="font-size:24px;font-weight:700;color:#2dd4bf;">${{Math.round(k.totalHours).toLocaleString()}}</div></div>
      </div>` +
      table('Engagement by Content Type', ctData.map(o=>`<tr><td ${{miniTd}}>${{o.type}}</td><td ${{miniTdR}}>${{o.count}}</td><td ${{miniTdR}}>${{o.users}}</td></tr>`).join(''), ['Type','Engagements','Learners']) +
      table('Most Popular Topics', topicData.slice(0,25).map(o=>`<tr><td ${{miniTd}}>${{o.skill}}</td><td ${{miniTdR}}>${{o.users}}</td></tr>`).join(''), ['Topic','Learners']) +
      table('Market Engagement', marketData.map(o=>`<tr><td ${{miniTd}}>${{o.market}}</td><td ${{miniTdR}}>${{o.users}}</td></tr>`).join(''), ['Market','Learners']) +
      table('Department Engagement', deptData.map(o=>`<tr><td ${{miniTd}}>${{o.dept}}</td><td ${{miniTdR}}>${{o.users}}</td></tr>`).join(''), ['Department','Learners']);

    window.print();
  }};

  // ── Export: Excel ─────────────────────────────────────────────────────────
  window.runExportXLSX = function(){{
    sel('export-menu').classList.remove('open');
    const rows = filterRows(currentCat);
    const k = computeKPIs(rows);
    const ctData = computeContentTypes(rows);
    const topicData = computeTopics(rows);
    const marketData = computeMarkets(rows);
    const deptData = computeDepartments(rows);
    const dateStr = new Date().toLocaleDateString('en-US',{{year:'numeric',month:'long',day:'numeric'}});
    const parts = [];
    if(currentCat !== 'All') parts.push('Group: ' + currentCat);
    if(hideTLG) parts.push('TLG Hidden');

    function makeSheet(rows_, colWidths){{
      const ws = XLSX.utils.aoa_to_sheet(rows_);
      ws['!cols'] = colWidths.map(w => ({{wch: w}}));
      return ws;
    }}

    const wb = XLSX.utils.book_new();
    const summaryRows = [
      ['Learning Engagement Report'],
      ['Generated: ' + dateStr],
      ['Filters: ' + (parts.length ? parts.join(' | ') : 'No filters active \\u2014 showing all data')],
      [],
      ['SUMMARY'],
      ['Unique Learners','Avg Engagements / Learner','Total Hours Viewed'],
      [k.learners, k.learners?(k.engagements/k.learners).toFixed(1):'0', Math.round(k.totalHours)],
      [],
      ['ENGAGEMENT BY CONTENT TYPE'],
      ['Type','Engagements','Learners'],
      ...ctData.map(o=>[o.type,o.count,o.users]),
      [],
      ['MOST POPULAR TOPICS'],
      ['Topic','Learners'],
      ...topicData.map(o=>[o.skill,o.users]),
      [],
      ['MARKET ENGAGEMENT'],
      ['Market','Learners'],
      ...marketData.map(o=>[o.market,o.users]),
      [],
      ['DEPARTMENT ENGAGEMENT'],
      ['Department','Learners'],
      ...deptData.map(o=>[o.dept,o.users]),
    ];
    XLSX.utils.book_append_sheet(wb, makeSheet(summaryRows,[32,26,18]), 'Summary');
    XLSX.writeFile(wb, 'learning-engagement-report.xlsx');
  }};

  document.getElementById('filterSeg').addEventListener('click', (e)=>{{
    const btn = e.target.closest('button');
    if(!btn) return;
    currentCat = btn.dataset.cat;
    document.querySelectorAll('#filterSeg button').forEach(b=>b.classList.toggle('active', b===btn));
    render();
  }});
  document.getElementById('topicsToggle').addEventListener('click', ()=>{{ topicsExpanded = !topicsExpanded; render(); }});
  document.getElementById('marketsToggle').addEventListener('click', ()=>{{ marketsExpanded = !marketsExpanded; render(); }});
  document.getElementById('deptsToggle').addEventListener('click', ()=>{{ deptsExpanded = !deptsExpanded; render(); }});

  render();

  // ── Triple-click title → home ────────────────────────────────────────────
  (function(){{
    let n = 0, t;
    const h = document.querySelector('.header h1');
    if(h) h.addEventListener('click', function(){{
      n++; clearTimeout(t);
      if(n >= 3){{ n = 0; window.location.href = 'index.html'; }}
      else t = setTimeout(function(){{ n = 0; }}, 1500);
    }});
  }})();
}})();
</script>
</body>
</html>"""
    return html


def main():
    print("Generating Learning Engagement dashboard...")
    records, total_active, date_label = load_learning_engagement_data()
    print(f"    {len(records)} engagement records, data through {date_label}")
    html = generate_html(records, total_active, date_label)
    out = os.path.join(SCRIPT_DIR, 'learning-engagement.html')
    with open(out, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"Written to: {out}")


if __name__ == '__main__':
    main()
