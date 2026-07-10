#!/usr/bin/env python3
"""update_leaderboard_dashboard.py — generates leaderboard.html from LMS + Salesforce exports."""

import os, re, json, datetime, glob, openpyxl
from html.parser import HTMLParser

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
LEADERBOARD_DIR = 'leaderboard-data'
OUTPUT_FILE     = 'leaderboard.html'
WINDOW_DAYS     = 45

# LMS column indices (0-based)
COL_FIRST           = 3
COL_LAST            = 4
COL_EMAIL           = 5
COL_JOBTITLE        = 6
COL_REGION          = 7
COL_MARKET          = 8
COL_BRANCH          = 9
COL_HIRE_DATE       = 15
COL_CURRIC_COMPLETE = 20
COL_ASSIGN_DATE     = 21

# ---------------------------------------------------------------------------
# HTML-as-XLS parser (Salesforce exports .xls files that are actually HTML)
# ---------------------------------------------------------------------------
class _TableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows = []
        self._row = []
        self._cell = ''
        self._in = False

    def handle_starttag(self, tag, attrs):
        if tag in ('td', 'th'):
            self._in = True
            self._cell = ''
        elif tag == 'tr':
            self._row = []

    def handle_endtag(self, tag):
        if tag in ('td', 'th'):
            self._row.append(self._cell.strip())
            self._in = False
        elif tag == 'tr':
            if self._row:
                self.rows.append(self._row)

    def handle_data(self, data):
        if self._in:
            self._cell += data


def _parse_html_xls(path):
    with open(path, encoding='utf-8', errors='ignore') as f:
        content = f.read()
    p = _TableParser()
    p.feed(content)
    if not p.rows:
        return []
    headers = p.rows[0]
    return [dict(zip(headers, row)) for row in p.rows[1:] if row]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _date(val):
    if not val:
        return None
    if hasattr(val, 'strftime'):
        return val.strftime('%Y-%m-%d')
    s = str(val).strip()
    if not s or s.lower() == 'none':
        return None
    m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{4})', s)
    if m:
        return f'{int(m.group(3)):04d}-{int(m.group(1)):02d}-{int(m.group(2)):02d}'
    if re.match(r'\d{4}-\d{2}-\d{2}', s):
        return s[:10]
    return None


def _norm(s):
    """Title-case normalize a name for comparison."""
    return ' '.join(w.capitalize() for w in str(s or '').split())


def _file_dt(fname):
    """Return a datetime derived from a filename date pattern, falling back to mtime."""
    base = os.path.basename(str(fname))
    m = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', base)
    if m:
        return datetime.datetime(int(m.group(3)), int(m.group(1)), int(m.group(2)))
    m = re.search(r'(\d{2})\.(\d{2})(?!\.\d)', base)
    if m:
        return datetime.datetime(datetime.datetime.today().year, int(m.group(1)), int(m.group(2)))
    m = re.search(r'(\d{4})-(\d{2})', base)
    if m:
        return datetime.datetime(int(m.group(1)), int(m.group(2)), 1)
    return datetime.datetime.fromtimestamp(os.path.getmtime(str(fname)))


# ---------------------------------------------------------------------------
# File detection
# ---------------------------------------------------------------------------
def _find_files():
    xlsx_files = glob.glob(os.path.join(LEADERBOARD_DIR, '*.xlsx'))
    xls_files  = glob.glob(os.path.join(LEADERBOARD_DIR, '*.xls'))

    if not xlsx_files:
        raise FileNotFoundError('No .xlsx LMS file in leaderboard-data/')
    lms_path = sorted(xlsx_files)[-1]

    cw_path = sh_path = None
    for p in xls_files:
        with open(p, encoding='utf-8', errors='ignore') as f:
            chunk = f.read(2000)
        if 'Opportunity Owner Email' in chunk:
            cw_path = p
        elif 'From Stage' in chunk:
            sh_path = p

    if not cw_path:
        raise FileNotFoundError('Closed Won .xls not found in leaderboard-data/')
    if not sh_path:
        raise FileNotFoundError('Stage History .xls not found in leaderboard-data/')

    return lms_path, cw_path, sh_path


# ---------------------------------------------------------------------------
# Parse LMS
# ---------------------------------------------------------------------------
def _parse_lms(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    hires = {}
    cohort_start = None

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) <= COL_CURRIC_COMPLETE:
            continue
        first = str(row[COL_FIRST] or '').strip()
        last  = str(row[COL_LAST]  or '').strip()
        if not first and not last:
            continue
        email     = str(row[COL_EMAIL] or '').strip().lower()
        hire_date = _date(row[COL_HIRE_DATE])
        assign_dt = _date(row[COL_ASSIGN_DATE]) if len(row) > COL_ASSIGN_DATE else None
        if assign_dt and (cohort_start is None or assign_dt < cohort_start):
            cohort_start = assign_dt
        hires[email] = {
            'name':              f'{first} {last}',
            'jobTitle':          str(row[COL_JOBTITLE] or '').strip(),
            'region':            str(row[COL_REGION]   or '').strip(),
            'market':            str(row[COL_MARKET]   or '').strip(),
            'branch':            str(row[COL_BRANCH]   or '').strip(),
            'email':             email,
            'hireDate':          hire_date,
            'assignDate':        assign_dt,
            'curriculumComplete': 'Yes' if str(row[COL_CURRIC_COMPLETE] or '').strip().lower() == 'yes' else 'No',
        }

    return list(hires.values()), cohort_start


# ---------------------------------------------------------------------------
# Stage index  {opp_id: {salesQualifiedBy, engageBy}}
# ---------------------------------------------------------------------------
def _stage_index(stage_rows):
    idx = {}
    for row in stage_rows:
        opp_id   = row.get('Opportunity ID', '').strip()
        to_stage = row.get('To Stage', '').strip()
        mod_by   = _norm(row.get('Last Modified By', ''))
        mod_date = row.get('Last Modified', '') or ''
        if to_stage not in ('Sales Qualified', 'Engage') or not opp_id:
            continue
        key = 'salesQualifiedBy' if to_stage == 'Sales Qualified' else 'engageBy'
        if opp_id not in idx:
            idx[opp_id] = {}
        prev = idx[opp_id].get(key)
        if prev is None or mod_date > prev[0]:
            idx[opp_id][key] = (mod_date, mod_by or None)

    return {
        oid: {
            'salesQualifiedBy': v.get('salesQualifiedBy', (None, None))[1],
            'engageBy':         v.get('engageBy',         (None, None))[1],
        }
        for oid, v in idx.items()
    }


# ---------------------------------------------------------------------------
# Build DEALS + VERIFICATION
# ---------------------------------------------------------------------------
def _build_data(cw_rows, stage_idx, hires, cohort_start):
    from datetime import date as _d

    by_email      = {h['email']: h for h in hires}
    name_to_email = {_norm(h['name']): h['email'] for h in hires}

    verif = {h['email']: {
        'name':             h['name'],
        'jobTitle':         h['jobTitle'],
        'market':           h['market'],
        'hireDate':         h['hireDate'],
        'isOwner':          False,
        'isCreatedBy':      False,
        'isLastModifiedBy': False,
        'ownedCount':       0,
        'createdCount':     0,
        'lastModCount':     0,
    } for h in hires}

    deals = []

    for row in cw_rows:
        if row.get('Stage', '').strip() != 'Closed Won':
            continue

        owner_email = row.get('Opportunity Owner Email', '').strip().lower()
        created_by  = _norm(row.get('Created By', ''))
        last_mod_by = _norm(row.get('Last Modified By', ''))

        # Verification counts across all Closed Won rows
        if owner_email in verif:
            verif[owner_email]['ownedCount'] += 1
            verif[owner_email]['isOwner'] = True
        cb_email = name_to_email.get(created_by)
        if cb_email and cb_email in verif:
            verif[cb_email]['createdCount'] += 1
            verif[cb_email]['isCreatedBy'] = True
        lm_email = name_to_email.get(last_mod_by)
        if lm_email and lm_email in verif:
            verif[lm_email]['lastModCount'] += 1
            verif[lm_email]['isLastModifiedBy'] = True

        # Deals — cohort members, close date >= cohort_start
        if owner_email not in by_email:
            continue
        close_date = _date(row.get('Close Date', ''))
        if cohort_start and close_date and close_date < cohort_start:
            continue

        hire   = by_email[owner_email]
        opp_id = row.get('Opportunity ID', '').strip()

        try:
            amount = float(str(row.get('Amount', '0') or '0').replace(',', '').strip())
        except ValueError:
            amount = 0.0

        try:
            htc = (_d.fromisoformat(close_date) - _d.fromisoformat(hire['hireDate'])).days
        except Exception:
            htc = None

        try:
            ad = hire.get('assignDate')
            atc = (_d.fromisoformat(close_date) - _d.fromisoformat(ad)).days if ad and close_date else None
        except Exception:
            atc = None

        si = stage_idx.get(opp_id, {})
        deals.append({
            'name':               hire['name'],
            'jobTitle':           hire['jobTitle'],
            'market':             hire['market'],
            'branch':             hire['branch'],
            'accountDesignation': row.get('Account Designation', '').strip(),
            'revenueType':        row.get('Revenue Type', '').strip(),
            'accountName':        row.get('Account Name', '').strip(),
            'hireDate':           hire['hireDate'],
            'assignDate':         hire.get('assignDate'),
            'closeDate':          close_date,
            'hireToCloseDays':    htc,
            'assignToCloseDays':  atc,
            'curriculumComplete': hire['curriculumComplete'],
            'amount':             round(amount, 2),
            'oppId':              opp_id,
            'salesQualifiedBy':   si.get('salesQualifiedBy'),
            'engageBy':           si.get('engageBy'),
        })

    verification = [verif[h['email']] for h in hires]
    return deals, verification


# ---------------------------------------------------------------------------
# Generate HTML
# ---------------------------------------------------------------------------
def _generate_html(hires, deals, verification, source_as_of, cohort_start):
    hires_pub = [{k: v for k, v in h.items() if k != 'email'} for h in hires]

    hires_json = json.dumps(hires_pub,    ensure_ascii=False)
    deals_json = json.dumps(deals,        ensure_ascii=False)
    verif_json = json.dumps(verification, ensure_ascii=False)

    src_dt = source_as_of
    source_date_js    = f'new Date({src_dt.year}, {src_dt.month - 1}, {src_dt.day})'
    file_date_label   = f'{src_dt.strftime("%B")} {src_dt.day}, {src_dt.year}'

    try:
        from datetime import date as _d
        cs = _d.fromisoformat(cohort_start)
        cohort_label_long  = cs.strftime('%B %-d, %Y')
        cohort_label_short = cs.strftime('%b %-d, %Y')
    except Exception:
        cohort_label_long  = cohort_start or 'Jun 4, 2026'
        cohort_label_short = cohort_label_long

    html = _HTML_TEMPLATE
    html = html.replace('__HIRES_JSON__',         hires_json)
    html = html.replace('__DEALS_JSON__',         deals_json)
    html = html.replace('__VERIF_JSON__',         verif_json)
    html = html.replace('__SOURCE_DATE_JS__',     source_date_js)
    html = html.replace('__FILE_DATE_LABEL__',    file_date_label)
    html = html.replace('__COHORT_LABEL_LONG__',  cohort_label_long)
    html = html.replace('__COHORT_LABEL_SHORT__', cohort_label_short)
    html = html.replace('__WINDOW_DAYS__',        str(WINDOW_DAYS))

    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'Written → {OUTPUT_FILE}')
    print(f'  {len(hires)} hires · {len(deals)} cohort deals · cohort start {cohort_start}')


# ---------------------------------------------------------------------------
# HTML Template
# Resmie's sections/charts/content preserved exactly.
# CSS variables and header swapped to match the playbook-dashboard design system.
# ---------------------------------------------------------------------------
_HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Accelerate Leaderboard</title>
<style>
  :root {
    --bg:#0f1117; --surface:#1a1d27; --surface2:#22263a; --border:#2e3350;
    --accent:#4f8ef7; --accent2:#7c5cfc; --accent3:#f7c94f;
    --text:#e8ecf4; --muted:#7b82a0; --green:#3ecf8e; --red:#f76f6f;
    --teal:#2dd4bf; --green-subtle:#3ecf8e22; --red-subtle:#f76f6f22;
    --font:'Segoe UI',system-ui,sans-serif;
  }
  body.light-mode {
    --bg:#f4f6fb; --surface:#ffffff; --surface2:#eef1f7; --border:#d0d7e8;
    --accent:#2563eb; --accent2:#6d28d9; --accent3:#d97706;
    --text:#1a1d27; --muted:#475569; --green:#059669; --red:#dc2626;
    --teal:#0f766e; --green-subtle:#05966922; --red-subtle:#dc262622;
  }
  body.light-mode select { color-scheme:light; }
  * { box-sizing:border-box; margin:0; padding:0; }
  body { background:var(--bg); color:var(--text); font-family:var(--font); min-height:100vh; transition:background .2s,color .2s; }

  /* ---- HEADER ---- */
  .header { padding:20px 28px 16px; border-bottom:1px solid var(--border); display:grid; grid-template-columns:1fr auto 1fr; align-items:center; gap:12px; background:var(--surface); }
  .header-left { display:flex; flex-direction:column; }
  .header-center { display:flex; justify-content:center; align-items:center; }
  .header-right { display:flex; justify-content:flex-end; align-items:center; gap:8px; }
  .header h1 { font-size:18px; font-weight:700; letter-spacing:.3px; cursor:default; }
  .header h1 span { color:var(--muted); font-weight:400; }
  .header-date { font-size:11px; color:var(--muted); margin-top:3px; }
  .kma-logo { height:38px; width:auto; display:block; }
  .kma-logo-light { display:none; }
  .light-mode .kma-logo-dark { display:none; }
  .light-mode .kma-logo-light { display:block; }
  .btn-theme { background:transparent; border:1px solid var(--border); color:var(--muted); border-radius:6px; padding:5px 12px; font-size:12px; cursor:pointer; transition:all .15s; }
  .btn-theme:hover { border-color:var(--accent); color:var(--text); }

  /* ---- WRAP ---- */
  .wrap { padding:28px 28px 80px; }

  /* ---- DATA NOTE ---- */
  .data-note { display:flex; align-items:flex-start; gap:10px; background:var(--surface); border:1px solid var(--border); border-radius:8px; padding:12px 16px; margin-bottom:24px; font-size:13px; color:var(--muted); line-height:1.55; }

  /* ---- STAT STRIP ---- */
  .stats { display:grid; grid-template-columns:repeat(auto-fit, minmax(170px, 1fr)); gap:14px; margin-bottom:28px; }
  .stat { background:var(--surface); border:1px solid var(--border); border-radius:10px; padding:16px 18px; }
  .stat-label { font-size:11px; text-transform:uppercase; letter-spacing:.7px; color:var(--muted); margin-bottom:6px; }
  .stat-value { font-size:28px; font-weight:700; line-height:1; color:var(--text); }
  .stat-value.green { color:var(--green); }
  .stat-value.teal  { color:var(--teal);  }
  .stat-value.accent{ color:var(--accent);}
  .stat-sub { font-size:11px; color:var(--muted); margin-top:4px; }

  /* ---- SECTION CARD ---- */
  .section { background:var(--surface); border:1px solid var(--border); border-radius:10px; margin-bottom:24px; overflow:hidden; }
  .section-head { padding:18px 22px 14px; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:flex-start; gap:16px; flex-wrap:wrap; }
  .section-head h2 { font-size:14px; font-weight:700; text-transform:uppercase; letter-spacing:.5px; color:var(--muted); margin:0 0 4px; display:flex; align-items:center; gap:8px; }
  body.light-mode .section-head h2 { color:var(--text); }
  .section-head h2 .count { font-size:11px; font-weight:700; color:var(--accent); background:var(--surface2); padding:2px 8px; border-radius:20px; text-transform:none; letter-spacing:0; }
  .section-head p { margin:0; font-size:13px; color:var(--muted); line-height:1.5; max-width:640px; }
  .search-box { background:var(--surface2); border:1px solid var(--border); color:var(--text); border-radius:20px; padding:6px 12px; font-size:13px; outline:none; width:190px; font-family:inherit; }
  .search-box:focus { border-color:var(--accent); }

  /* ---- MARKET BARS ---- */
  .market-chart { padding:20px 22px 24px; }
  .market-row { display:grid; grid-template-columns:130px 1fr 110px; align-items:center; gap:14px; padding:8px 0; }
  .market-row + .market-row { border-top:1px solid var(--border); }
  .market-label { font-size:14px; font-weight:600; color:var(--text); }
  .market-sub   { font-size:11px; color:var(--muted); margin-top:2px; }
  .market-bar-track { position:relative; height:20px; background:var(--surface2); border-radius:4px; overflow:hidden; }
  .market-bar-fill  { position:absolute; left:0; top:0; bottom:0; border-radius:4px; transition:width .5s ease; }
  .market-amount { font-size:14px; font-weight:700; text-align:right; color:var(--text); font-variant-numeric:tabular-nums; }

  /* ---- TABLE ---- */
  .table-scroll { overflow-x:auto; }
  table { width:100%; border-collapse:collapse; font-size:13px; }
  thead th { text-align:left; font-size:11px; font-weight:600; text-transform:uppercase; letter-spacing:.5px; color:var(--muted); background:var(--surface2); padding:10px 14px; border-bottom:1px solid var(--border); white-space:nowrap; cursor:pointer; user-select:none; position:sticky; top:0; }
  thead th:hover { color:var(--text); }
  thead th.sorted-asc::after  { content:' \2191'; color:var(--accent); }
  thead th.sorted-desc::after { content:' \2193'; color:var(--accent); }
  tbody td { padding:10px 14px; border-bottom:1px solid var(--border); vertical-align:middle; }
  tbody tr:last-child td { border-bottom:none; }
  tbody tr:hover { background:var(--surface2); }
  .td-name   { font-weight:600; }
  .td-muted  { color:var(--muted); }
  .td-amount { text-align:right; color:var(--teal); font-weight:700; font-variant-numeric:tabular-nums; }
  .td-num    { text-align:right; font-variant-numeric:tabular-nums; }
  .td-self   { color:var(--green); font-weight:600; }

  /* ---- BADGES ---- */
  .badge { display:inline-block; font-size:11px; font-weight:600; padding:2px 9px; border-radius:20px; }
  .badge-yes      { background:var(--green-subtle); color:var(--green); }
  .badge-no       { background:var(--red-subtle);   color:var(--red); }
  .badge-growth   { background:#4f8ef722; color:var(--accent); }
  .badge-retention{ background:#f7c94f22; color:var(--accent3); }
  .badge-in       { color:var(--teal); }
  .badge-out      { color:var(--muted); }

  /* ---- WINDOW BAR ---- */
  .win-bar-wrap { display:flex; align-items:center; gap:8px; min-width:140px; }
  .win-bar { position:relative; flex:1; height:6px; background:var(--border); border-radius:4px; overflow:hidden; }
  .win-bar .fill { position:absolute; left:0; top:0; bottom:0; border-radius:4px; }
  .win-days { font-size:11px; font-weight:600; min-width:42px; text-align:right; color:var(--muted); font-variant-numeric:tabular-nums; }

  /* ---- EMPTY STATE ---- */
  .empty-state { padding:52px 24px; text-align:center; color:var(--muted); font-size:13px; }

  /* ---- FOOTER ---- */
  footer { text-align:center; font-size:11px; color:var(--muted); margin-top:40px; opacity:.7; }

  @media(max-width:760px) {
    .header { grid-template-columns:1fr auto; }
    .header-center { display:none; }
    .market-row { grid-template-columns:100px 1fr 80px; }
  }
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <h1>Accelerate Leaderboard <span>/ Sales Performance</span></h1>
    <div class="header-date" id="header-date">Data through __FILE_DATE_LABEL__</div>
  </div>
  <div class="header-center">
    <img src="KMA-wht.svg" class="kma-logo kma-logo-dark" alt="KM Academy">
    <img src="KMA-drk.svg" class="kma-logo kma-logo-light" alt="KM Academy">
  </div>
  <div class="header-right">
    <button class="btn-theme" id="btn-theme" onclick="toggleTheme()">&#9728; Light</button>
  </div>
</div>

<div class="wrap">

  <div class="data-note">
    <span style="font-size:15px;flex-shrink:0;">&#9432;</span>
    <span>To appear on this board, a rep must close a deal within their first 45 days of the Accelerate program, complete the curriculum, and personally move that deal through both Sales Qualified and Engage. Deals handed off to a manager or teammate do not count.</span>
  </div>

  <div class="data-note" id="dataNote"></div>

  <div class="stats" id="statStrip"></div>

  <!-- MARKET CHART -->
  <div class="section">
    <div class="section-head">
      <div>
        <h2>Closed-Won Revenue by Market <span class="count" id="marketCount">0</span></h2>
        <p>Closed Won revenue for the Accelerate cohort from __COHORT_LABEL_LONG__ onward, rolled up by market. Includes all cohort deals, not just leaderboard-qualifying wins.</p>
      </div>
    </div>
    <div id="marketChart" class="market-chart"></div>
  </div>

  <!-- MAIN LEADERBOARD -->
  <div class="section">
    <div class="section-head">
      <div>
        <h2>Closed-Won Leaderboard <span class="count" id="leaderboardCount">0</span></h2>
        <p>Reps who closed a deal within their first __WINDOW_DAYS__ days of the Accelerate program, completed the curriculum, and personally moved the deal through Sales Qualified and Engage. Each qualifying win is its own row.</p>
      </div>
      <input class="search-box" id="leaderboardSearch" placeholder="Filter by rep or account&hellip;" oninput="renderLeaderboard()">
    </div>
    <div id="leaderboardBody"></div>
  </div>

  <!-- ON DECK -->
  <div class="section">
    <div class="section-head">
      <div>
        <h2>On Deck <span class="count" id="onDeckCount">0</span></h2>
        <p>Reps currently inside their __WINDOW_DAYS__-day window with deals that don't yet qualify for the leaderboard. Shows what each deal is still missing.</p>
      </div>
    </div>
    <div id="onDeckBody"></div>
  </div>

  <!-- WINDOW TRACKER -->
  <div class="section">
    <div class="section-head">
      <div>
        <h2>Program Window Tracker <span class="count" id="trackerCount">0</span></h2>
        <p>Every Accelerate cohort member and where they sit against the __WINDOW_DAYS__-day clock, from program start to today. Leaderboard eligibility requires both an open window and a completed curriculum. Reps past day __WINDOW_DAYS__ are grouped separately below.</p>
      </div>
      <input class="search-box" id="trackerSearch" placeholder="Filter by rep&hellip;" oninput="renderTracker()">
    </div>
    <div class="table-scroll">
      <table id="trackerTable">
        <thead><tr>
          <th data-key="name" data-type="string">Rep</th>
          <th data-key="jobTitle" data-type="string">Job Title</th>
          <th data-key="market" data-type="string">Market</th>
          <th data-key="assignDate" data-type="date">Program Start</th>
          <th data-key="daysSince" data-type="number">Days in Program</th>
          <th data-key="eligible" data-type="string">__WINDOW_DAYS__-Day Window</th>
          <th data-key="curriculumComplete" data-type="string">Curriculum</th>
        </tr></thead>
        <tbody id="trackerTbody"></tbody>
      </table>
    </div>
    <div id="expiredTrackerWrap" style="margin-top:8px;">
      <button id="expiredToggle" onclick="toggleExpired()" style="background:none;border:1px solid var(--border);color:var(--muted);font-size:12px;padding:6px 14px;border-radius:6px;cursor:pointer;display:flex;align-items:center;gap:6px;">
        <span id="expiredToggleIcon">&#9660;</span> Show expired reps (<span id="expiredCount">0</span> past day __WINDOW_DAYS__)
      </button>
      <div id="expiredTrackerBody" style="display:none;margin-top:8px;">
        <div class="table-scroll">
          <table id="expiredTable">
            <thead><tr>
              <th>Rep</th><th>Job Title</th><th>Market</th><th>Program Start</th><th>Days in Program</th><th>Status</th><th>Curriculum</th>
            </tr></thead>
            <tbody id="expiredTbody"></tbody>
          </table>
        </div>
      </div>
    </div>
  </div>

  <!-- COHORT HISTORY -->
  <div class="section">
    <div class="section-head">
      <div>
        <h2>Cohort Closed-Won History <span class="count" id="historyCount">0</span> <span style="font-size:11px;font-weight:600;color:var(--muted);text-transform:none;letter-spacing:0;">since __COHORT_LABEL_SHORT__</span></h2>
        <p>Every Closed Won deal booked by an Accelerate-cohort rep with a close date on or after __COHORT_LABEL_LONG__, ordered from earliest hire date to close date, with who moved each opportunity into Sales Qualified and who then advanced it into Engage.</p>
      </div>
      <input class="search-box" id="historySearch" placeholder="Filter by rep or account&hellip;" oninput="renderHistory()">
    </div>
    <div class="table-scroll">
      <table id="historyTable">
        <thead><tr>
          <th data-key="name" data-type="string">Rep</th>
          <th data-key="accountName" data-type="string">Account</th>
          <th data-key="accountDesignation" data-type="string">Designation</th>
          <th data-key="revenueType" data-type="string">Revenue Type</th>
          <th data-key="assignDate" data-type="date">Program Start</th>
          <th data-key="closeDate" data-type="date">Close Date</th>
          <th data-key="assignToCloseDays" data-type="number">Program Day</th>
          <th data-key="salesQualifiedBy" data-type="string">Sales Qualified</th>
          <th data-key="engageBy" data-type="string">Engage</th>
          <th data-key="amount" data-type="number">Amount</th>
        </tr></thead>
        <tbody id="historyTbody"></tbody>
      </table>
    </div>
  </div>

  <!-- VERIFICATION -->
  <div class="section">
    <div class="section-head">
      <div>
        <h2>Salesforce Activity Verification <span class="count" id="verifyCount">0</span></h2>
        <p>Confirms whether each Accelerate hire has personally created an opportunity record in Salesforce &mdash; the strongest signal of system activity, since &ldquo;Owner&rdquo; alone can just mean a deal was assigned to them.</p>
      </div>
      <input class="search-box" id="verifySearch" placeholder="Filter by rep&hellip;" oninput="renderVerify()">
    </div>
    <div class="table-scroll">
      <table id="verifyTable">
        <thead><tr>
          <th data-key="name" data-type="string">Rep</th>
          <th data-key="jobTitle" data-type="string">Job Title</th>
          <th data-key="market" data-type="string">Market</th>
          <th data-key="hireDate" data-type="date">Hire Date</th>
          <th data-key="isCreatedBy" data-type="bool">Created a Record</th>
          <th data-key="createdCount" data-type="number">Opps Created</th>
          <th data-key="ownedCount" data-type="number">Opps Owned</th>
          <th data-key="isLastModifiedBy" data-type="bool">Recent Activity</th>
        </tr></thead>
        <tbody id="verifyTbody"></tbody>
      </table>
    </div>
  </div>

  <footer>Source: Accelerate Curriculum Report &middot; Opportunity Export &middot; generated client-side, recalculates daily</footer>
</div>

<script>
const DEALS        = __DEALS_JSON__;
const HIRES        = __HIRES_JSON__;
const VERIFICATION = __VERIF_JSON__;
const WINDOW_DAYS  = __WINDOW_DAYS__;
const SOURCE_AS_OF = __SOURCE_DATE_JS__;

const TODAY = new Date();
TODAY.setHours(0,0,0,0);

function parseDate(s){ if(!s) return null; const [y,m,d]=s.split('-').map(Number); return new Date(y,m-1,d); }
function daysBetween(a,b){ return Math.round((b-a)/86400000); }
function fmtDate(s){ if(!s) return '&#8212;'; return parseDate(s).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'}); }
function fmtMoney(n){ return '$'+n.toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}); }

document.getElementById('dataNote').innerHTML =
  '<span style="font-size:15px;flex-shrink:0;">&#9432;</span>' +
  '<span><strong>New hire and curriculum data reflects the source reports pulled as of ' +
  SOURCE_AS_OF.toLocaleDateString('en-US',{month:'long',day:'numeric',year:'numeric'}) +
  '. Closed Won and opportunity data is scoped to deals closed on or after __COHORT_LABEL_LONG__.</strong></span>';

// ---- Augment hires ----
const hireMap = {};
HIRES.forEach(h => {
  h.daysSince = h.assignDate ? daysBetween(parseDate(h.assignDate), TODAY) : 999;
  h.eligible  = h.daysSince >= 0 && h.daysSince <= WINDOW_DAYS;
  h.curriculumOk = h.curriculumComplete === 'Yes';
  h.leaderboardEligible = h.eligible && h.curriculumOk;
  hireMap[h.name] = h;
});

const leaderboardRows = DEALS.filter(d =>
  d.assignToCloseDays !== null &&
  d.assignToCloseDays >= 0 &&
  d.assignToCloseDays <= WINDOW_DAYS &&
  d.curriculumComplete === 'Yes' &&
  d.salesQualifiedBy === d.name &&
  d.engageBy === d.name
);

// ---- Stat strip ----
const totalAmount = leaderboardRows.reduce((s,d) => s+d.amount, 0);
const onBoardReps = new Set(leaderboardRows.map(d => d.name)).size;
const stats = [
  {label:'Reps In '+WINDOW_DAYS+'-Day Window',     value:HIRES.filter(h=>h.eligible).length,            sub:'of '+HIRES.length+' cohort members',              color:'var(--teal)'},
  {label:'Curriculum Complete + In Window',          value:HIRES.filter(h=>h.leaderboardEligible).length, sub:'eligible to appear on the board',                 color:'var(--accent3)'},
  {label:'Reps On The Board',                        value:onBoardReps,                                    sub:'eligible with a qualifying Closed Won deal',      color:'var(--green)'},
  {label:'Qualifying Deals',                         value:leaderboardRows.length,                         sub:'in-window, curriculum complete, self-progressed', color:'var(--accent)'},
  {label:'Leaderboard Revenue',                      value:fmtMoney(totalAmount),                          sub:'total across qualifying deals',                   color:'var(--accent2)'},
  {label:'Salesforce-Verified',                      value:VERIFICATION.filter(v=>v.isCreatedBy).length+' / '+VERIFICATION.length, sub:'cohort members with a created opportunity', color:'var(--teal)'},
];
document.getElementById('statStrip').innerHTML = stats.map(s =>
  '<div class="stat">' +
    '<div class="stat-label">'+s.label+'</div>' +
    '<div class="stat-value" style="color:'+s.color+'">'+s.value+'</div>' +
    '<div class="stat-sub">'+s.sub+'</div>' +
  '</div>'
).join('');

// ---- Market chart ----
(function(){
  const colors = [
    'var(--accent)', 'var(--teal)', 'var(--accent3)',
    'var(--green)',  'var(--accent2)', 'var(--red)'
  ];
  const byMarket = {};
  DEALS.forEach(d => {
    const m = d.market.replace(' Sls','');
    if(!byMarket[m]) byMarket[m] = {amount:0, count:0, reps:new Set()};
    byMarket[m].amount += d.amount;
    byMarket[m].count  += 1;
    byMarket[m].reps.add(d.name);
  });
  const rows = Object.entries(byMarket)
    .map(([market,v]) => ({market, amount:v.amount, count:v.count, repCount:v.reps.size}))
    .sort((a,b) => b.amount - a.amount);
  const max = Math.max(...rows.map(r=>r.amount), 1);
  document.getElementById('marketCount').textContent = rows.length;
  document.getElementById('marketChart').innerHTML = rows.length ? rows.map((r,i) =>
    '<div class="market-row">' +
      '<div><div class="market-label">'+r.market+'</div><div class="market-sub">'+r.count+' deal'+(r.count===1?'':'s')+' &middot; '+r.repCount+' rep'+(r.repCount===1?'':'s')+'</div></div>' +
      '<div class="market-bar-track"><div class="market-bar-fill" style="width:'+(r.amount/max*100).toFixed(1)+'%;background:'+colors[i%colors.length]+';opacity:.7;"></div></div>' +
      '<div class="market-amount">'+fmtMoney(r.amount)+'</div>' +
    '</div>'
  ).join('') : '<div class="empty-state">No deals yet</div>';
})();

// ---- Leaderboard ----
function renderLeaderboard(){
  const el = document.getElementById('leaderboardBody');
  const q  = (document.getElementById('leaderboardSearch').value||'').toLowerCase();
  const rows = leaderboardRows
    .filter(d => !q || d.name.toLowerCase().includes(q) || d.accountName.toLowerCase().includes(q))
    .sort((a,b) => b.amount - a.amount);
  document.getElementById('leaderboardCount').textContent = rows.length;
  if(!rows.length){
    el.innerHTML = '<div class="empty-state">' +
      '<div style="font-size:22px;margin-bottom:10px;">&#8212;</div>' +
      '<div style="font-size:14px;font-weight:600;margin-bottom:8px;">No qualifying wins yet</div>' +
      '<div style="font-size:13px;max-width:560px;margin:0 auto;line-height:1.6;">No rep has yet closed a deal within their first '+WINDOW_DAYS+' days of the Accelerate program while also being curriculum-complete and having personally moved the deal through Sales Qualified and Engage. Check the On Deck and Program Window Tracker sections below for current standings.</div>' +
    '</div>';
    return;
  }
  el.innerHTML = '<div class="table-scroll"><table>' +
    '<thead><tr><th>Rep</th><th>Job Title</th><th>Market</th><th>Designation</th><th>Revenue Type</th><th>Program Start</th><th>Close Date</th><th>Program Day</th><th>Curriculum</th><th>Amount</th></tr></thead>' +
    '<tbody>' + rows.map(d =>
      '<tr>' +
        '<td class="td-name">'+d.name+'</td>' +
        '<td class="td-muted">'+d.jobTitle+'</td>' +
        '<td class="td-muted">'+d.market.replace(' Sls','')+'</td>' +
        '<td><span class="badge '+(d.accountDesignation.toLowerCase()==='growth'?'badge-growth':'badge-retention')+'">'+(d.accountDesignation||'&#8212;')+'</span></td>' +
        '<td class="td-muted">'+d.revenueType+'</td>' +
        '<td class="td-muted">'+fmtDate(d.assignDate)+'</td>' +
        '<td class="td-muted">'+fmtDate(d.closeDate)+'</td>' +
        '<td class="td-num">'+(d.assignToCloseDays!==null?d.assignToCloseDays+'d':'&#8212;')+'</td>' +
        '<td><span class="badge '+(d.curriculumComplete==='Yes'?'badge-yes':'badge-no')+'">'+d.curriculumComplete+'</span></td>' +
        '<td class="td-amount">'+fmtMoney(d.amount)+'</td>' +
      '</tr>'
    ).join('') + '</tbody></table></div>';
}
renderLeaderboard();

// ---- On Deck ----
function renderOnDeck(){
  const lbSet = new Set(leaderboardRows.map(d => d.oppId));
  const inWindowNames = new Set(HIRES.filter(h => h.eligible).map(h => h.name));
  const onDeckDeals = DEALS.filter(d => inWindowNames.has(d.name) && !lbSet.has(d.oppId));
  const byRep = {};
  onDeckDeals.forEach(d => {
    if(!byRep[d.name]) byRep[d.name] = [];
    byRep[d.name].push(d);
  });
  const readyReps = HIRES.filter(h => h.eligible && h.curriculumOk && !onDeckDeals.some(d => d.name === h.name) && !leaderboardRows.some(d => d.name === h.name));
  const totalCount = Object.keys(byRep).length + readyReps.length;
  document.getElementById('onDeckCount').textContent = totalCount;
  if(!totalCount){
    document.getElementById('onDeckBody').innerHTML = '<div class="empty-state" style="padding:32px 0;"><div style="font-size:14px;font-weight:600;margin-bottom:6px;">Nothing on deck</div><div style="font-size:13px;">All in-window reps are either on the leaderboard or have no cohort deals yet.</div></div>';
    return;
  }
  let html = '<div class="table-scroll"><table><thead><tr><th>Rep</th><th>Account</th><th>Close Date</th><th>Program Day</th><th>What\'s Missing</th><th>Sales Qualified By</th><th>Engage By</th><th>Amount</th></tr></thead><tbody>';
  Object.entries(byRep).forEach(function([name, deals]){
    deals.sort((a,b) => b.amount - a.amount);
    deals.forEach(function(d, i){
      const missing = [];
      if(d.curriculumComplete !== 'Yes') missing.push('<span class="badge badge-no">Curriculum not done</span>');
      if(d.assignToCloseDays === null || d.assignToCloseDays > WINDOW_DAYS) missing.push('<span class="badge badge-out">Outside 45-day window</span>');
      if(d.salesQualifiedBy && d.salesQualifiedBy !== d.name) missing.push('<span class="badge" style="background:rgba(255,160,0,.18);color:var(--accent3);">Sales Qualified handed off</span>');
      if(!d.salesQualifiedBy) missing.push('<span class="badge" style="background:rgba(255,160,0,.18);color:var(--accent3);">Sales Qualified not recorded</span>');
      if(d.engageBy && d.engageBy !== d.name) missing.push('<span class="badge" style="background:rgba(255,100,80,.18);color:var(--red);">Engage handed off</span>');
      if(!d.engageBy) missing.push('<span class="badge" style="background:rgba(255,100,80,.18);color:var(--red);">Engage not recorded</span>');
      html +=
        '<tr>' +
        (i===0 ? '<td class="td-name" rowspan="'+deals.length+'">'+name+'</td>' : '') +
        '<td>'+d.accountName+'</td>' +
        '<td class="td-muted">'+fmtDate(d.closeDate)+'</td>' +
        '<td class="td-num">'+(d.assignToCloseDays!==null?d.assignToCloseDays+'d':'&#8212;')+'</td>' +
        '<td>'+missing.join(' ')+'</td>' +
        '<td class="'+(d.salesQualifiedBy===d.name?'td-self':'td-muted')+'">'+(d.salesQualifiedBy||'<span style="opacity:.5">Not recorded</span>')+'</td>' +
        '<td class="'+(d.engageBy===d.name?'td-self':'td-muted')+'">'+(d.engageBy||'<span style="opacity:.5">Not recorded</span>')+'</td>' +
        '<td class="td-amount">'+fmtMoney(d.amount)+'</td>' +
        '</tr>';
    });
  });
  readyReps.forEach(function(h){
    html +=
      '<tr>' +
      '<td class="td-name">'+h.name+'</td>' +
      '<td colspan="7" class="td-muted" style="font-style:italic;">Curriculum complete, in window &mdash; no cohort deals yet</td>' +
      '</tr>';
  });
  html += '</tbody></table></div>';
  document.getElementById('onDeckBody').innerHTML = html;
}
renderOnDeck();

// ---- Window Tracker ----
let trackerSort = {key:'daysSince', dir:1};
function _trackerRow(h){
  const pct   = Math.min(100, Math.max(0, h.daysSince/WINDOW_DAYS*100));
  const color = !h.eligible ? 'var(--border)' : pct<60 ? 'var(--teal)' : pct<90 ? 'var(--accent3)' : 'var(--red)';
  return '<tr>' +
    '<td class="td-name">'+h.name+'</td>' +
    '<td class="td-muted">'+h.jobTitle+'</td>' +
    '<td class="td-muted">'+h.market.replace(' Sls','')+'</td>' +
    '<td class="td-muted">'+fmtDate(h.assignDate)+'</td>' +
    '<td><div class="win-bar-wrap">' +
      '<div class="win-bar"><div class="fill" style="width:'+pct.toFixed(1)+'%;background:'+color+'"></div></div>' +
      '<span class="win-days">'+h.daysSince+'d</span>' +
    '</div></td>' +
    '<td>'+(h.eligible
      ? '<span class="badge badge-in">&#9679; In window</span>'
      : '<span class="badge badge-out">Closed &middot; day '+h.daysSince+'</span>')+'</td>' +
    '<td><span class="badge '+(h.curriculumComplete==='Yes'?'badge-yes':'badge-no')+'">'+h.curriculumComplete+'</span></td>' +
  '</tr>';
}
function renderTracker(){
  const q = (document.getElementById('trackerSearch').value||'').toLowerCase();
  const all = [...HIRES]
    .filter(h => !q || h.name.toLowerCase().includes(q))
    .sort((a,b) => {
      let av=a[trackerSort.key], bv=b[trackerSort.key];
      if(trackerSort.key==='assignDate'){ av=parseDate(av)||new Date(0); bv=parseDate(bv)||new Date(0); }
      if(av<bv) return -1*trackerSort.dir;
      if(av>bv) return  1*trackerSort.dir;
      return 0;
    });
  const active  = all.filter(h => h.eligible);
  const expired = all.filter(h => !h.eligible);
  document.getElementById('trackerCount').textContent = active.length;
  document.getElementById('trackerTbody').innerHTML = active.map(_trackerRow).join('') ||
    '<tr><td colspan="7" class="empty-state" style="padding:24px;text-align:center;">No reps currently in window</td></tr>';
  document.getElementById('expiredCount').textContent = expired.length;
  document.getElementById('expiredTbody').innerHTML = expired.map(_trackerRow).join('');
  document.getElementById('expiredTrackerWrap').style.display = expired.length ? '' : 'none';
  document.querySelectorAll('#trackerTable thead th').forEach(th => {
    th.classList.remove('sorted-asc','sorted-desc');
    if(th.dataset.key===trackerSort.key) th.classList.add(trackerSort.dir===1?'sorted-asc':'sorted-desc');
  });
}
let expiredOpen = false;
function toggleExpired(){
  expiredOpen = !expiredOpen;
  document.getElementById('expiredTrackerBody').style.display = expiredOpen ? '' : 'none';
  document.getElementById('expiredToggleIcon').textContent = expiredOpen ? '▲' : '▼';
}
document.querySelectorAll('#trackerTable thead th').forEach(th => {
  th.addEventListener('click', () => {
    const key = th.dataset.key;
    trackerSort.dir = trackerSort.key===key ? -trackerSort.dir : 1;
    trackerSort.key = key;
    renderTracker();
  });
});
renderTracker();

// ---- History ----
let histSort = {key:'assignDate', dir:1};
function renderHistory(){
  const q = (document.getElementById('historySearch').value||'').toLowerCase();
  const rows = [...DEALS]
    .filter(d => !q || d.name.toLowerCase().includes(q) || d.accountName.toLowerCase().includes(q))
    .sort((a,b) => {
      let av=a[histSort.key], bv=b[histSort.key];
      if(histSort.key==='assignDate'||histSort.key==='closeDate'){ av=av?parseDate(av):new Date(0); bv=bv?parseDate(bv):new Date(0); }
      if(av<bv) return -1*histSort.dir;
      if(av>bv) return  1*histSort.dir;
      return 0;
    });
  document.getElementById('historyCount').textContent = rows.length;
  document.getElementById('historyTbody').innerHTML = rows.map(d =>
    '<tr>' +
      '<td class="td-name">'+d.name+'</td>' +
      '<td>'+d.accountName+'</td>' +
      '<td><span class="badge '+(d.accountDesignation.toLowerCase()==='growth'?'badge-growth':'badge-retention')+'">'+(d.accountDesignation||'&#8212;')+'</span></td>' +
      '<td class="td-muted">'+d.revenueType+'</td>' +
      '<td class="td-muted">'+fmtDate(d.assignDate)+'</td>' +
      '<td class="td-muted">'+fmtDate(d.closeDate)+'</td>' +
      '<td class="td-num">'+(d.assignToCloseDays!==null?d.assignToCloseDays+'d':'&#8212;')+'</td>' +
      '<td class="'+(d.salesQualifiedBy===d.name?'td-self':'td-muted')+'">'+(d.salesQualifiedBy||'<span style="opacity:.5">Not recorded</span>')+'</td>' +
      '<td class="'+(d.engageBy===d.name?'td-self':'td-muted')+'">'+(d.engageBy||'<span style="opacity:.5">Not recorded</span>')+'</td>' +
      '<td class="td-amount">'+fmtMoney(d.amount)+'</td>' +
    '</tr>'
  ).join('');
  document.querySelectorAll('#historyTable thead th').forEach(th => {
    th.classList.remove('sorted-asc','sorted-desc');
    if(th.dataset.key===histSort.key) th.classList.add(histSort.dir===1?'sorted-asc':'sorted-desc');
  });
}
document.querySelectorAll('#historyTable thead th').forEach(th => {
  th.addEventListener('click', () => {
    const key = th.dataset.key;
    histSort.dir = histSort.key===key ? -histSort.dir : 1;
    histSort.key = key;
    renderHistory();
  });
});
renderHistory();
document.querySelector('#historyTable thead th[data-key="assignDate"]').classList.add('sorted-asc');

// ---- Verification ----
let verifSort = {key:'name', dir:1};
function renderVerify(){
  const q = (document.getElementById('verifySearch').value||'').toLowerCase();
  const rows = [...VERIFICATION]
    .filter(v => !q || v.name.toLowerCase().includes(q))
    .sort((a,b) => {
      let av=a[verifSort.key], bv=b[verifSort.key];
      if(verifSort.key==='hireDate'){ av=parseDate(av)||new Date(0); bv=parseDate(bv)||new Date(0); }
      if(typeof av==='boolean'){ av=av?1:0; bv=bv?1:0; }
      if(av<bv) return -1*verifSort.dir;
      if(av>bv) return  1*verifSort.dir;
      return 0;
    });
  document.getElementById('verifyCount').textContent = rows.length;
  document.getElementById('verifyTbody').innerHTML = rows.map(v =>
    '<tr>' +
      '<td class="td-name">'+v.name+'</td>' +
      '<td class="td-muted">'+v.jobTitle+'</td>' +
      '<td class="td-muted">'+v.market.replace(' Sls','')+'</td>' +
      '<td class="td-muted">'+fmtDate(v.hireDate)+'</td>' +
      '<td>'+(v.isCreatedBy
        ? '<span class="badge badge-yes">&#10003; Confirmed</span>'
        : '<span class="badge badge-no">No record found</span>')+'</td>' +
      '<td class="td-num">'+v.createdCount+'</td>' +
      '<td class="td-num">'+v.ownedCount+'</td>' +
      '<td>'+(v.isLastModifiedBy
        ? '<span class="badge badge-yes">Active</span>'
        : '<span class="badge" style="background:var(--surface2);color:var(--muted)">Dormant</span>')+'</td>' +
    '</tr>'
  ).join('');
  document.querySelectorAll('#verifyTable thead th').forEach(th => {
    th.classList.remove('sorted-asc','sorted-desc');
    if(th.dataset.key===verifSort.key) th.classList.add(verifSort.dir===1?'sorted-asc':'sorted-desc');
  });
}
document.querySelectorAll('#verifyTable thead th').forEach(th => {
  th.addEventListener('click', () => {
    const key = th.dataset.key;
    verifSort.dir = verifSort.key===key ? -verifSort.dir : 1;
    verifSort.key = key;
    renderVerify();
  });
});
renderVerify();

// ---- Theme ----
(function(){
  if(localStorage.getItem('pb-theme')==='light'){
    document.body.classList.add('light-mode');
    document.getElementById('btn-theme').innerHTML='&#9790; Dark';
  }
})();
function toggleTheme(){
  const light = document.body.classList.toggle('light-mode');
  document.getElementById('btn-theme').innerHTML = light ? '&#9790; Dark' : '&#9728; Light';
  localStorage.setItem('pb-theme', light ? 'light' : 'dark');
}

// ---- Triple-click on h1 returns to index.html ----
(function(){
  var n=0,t;
  var h=document.querySelector('.header h1');
  if(h) h.addEventListener('click',function(){
    n++;clearTimeout(t);
    if(n>=3){n=0;window.location.href='index.html';}
    else t=setTimeout(function(){n=0;},1500);
  });
})();
</script>
</body>
</html>"""


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    lms_path, cw_path, sh_path = _find_files()
    print(f'LMS:           {os.path.basename(lms_path)}')
    print(f'Closed Won:    {os.path.basename(cw_path)}')
    print(f'Stage History: {os.path.basename(sh_path)}')

    source_as_of = _file_dt(lms_path)

    hires, cohort_start = _parse_lms(lms_path)
    print(f'  {len(hires)} hires, cohort start {cohort_start}')

    cw_rows = _parse_html_xls(cw_path)
    sh_rows = _parse_html_xls(sh_path)
    print(f'  {len(cw_rows)} Closed Won rows, {len(sh_rows)} Stage History rows')

    stage_idx = _stage_index(sh_rows)
    print(f'  Stage index: {len(stage_idx)} unique opportunity IDs')

    deals, verification = _build_data(cw_rows, stage_idx, hires, cohort_start)
    print(f'  {len(deals)} cohort deals, {len(verification)} verification entries')

    _generate_html(hires, deals, verification, source_as_of, cohort_start)


if __name__ == '__main__':
    main()
