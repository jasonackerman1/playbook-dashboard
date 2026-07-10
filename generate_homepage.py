#!/usr/bin/env python3
"""
generate_homepage.py
Generates index.html — Learning Group landing page with 4 dashboard cards and live baked-in stats.
Run from playbook-dashboard/ after running the individual dashboard generators.
"""

import os, re, sys, warnings
from pathlib import Path
from datetime import datetime
import pandas as pd
warnings.filterwarnings('ignore')

SCRIPT_DIR = Path(__file__).parent
sys.path.insert(0, str(SCRIPT_DIR))

TLG = {
    "Jason Ackerman","Bianca Davis","James Parker","Resmie Biba",
    "Chris Curtis","Sara Thompson","Jeremy MacBean","Bradley Pierce",
    "Laura Sefcik","Samantha Maresca","Staci Musco","CJ Homer","Rich Moore","Dale Kinsey",
    "John Lechner","Resmie Nesimi","Samantha D'Angelo","Bianca DiPasquale","Doug Falk"
}

def fmt_month(yyyymm):
    try:
        return datetime.strptime(yyyymm, '%Y-%m').strftime('%B %Y')
    except Exception:
        return yyyymm

# ── Playbook Traffic stats ────────────────────────────────────────────────────
PLAYBOOK_MAP = {
    "dx_playbook": "DX Playbook",
    "dx_competencies_leadership_drivers": "DX Playbook",
    "healthcare_vertical_playbook": "Healthcare Playbook",
    "legal_vertical_playbook": "Legal Playbook",
    "salesforce_playbook": "Salesforce Playbook",
    "public_sector_playbook": "Public Sector Playbook",
    "accelerate_sales_playbook": "Accelerate",
    "iq501": "IQ501",
    "gc_ip_sales_playbook": "GC/IP Sales Playbook",
    "road_to_dx": "Road to DX",
}

def get_playbook(url):
    m = re.search(r'/playbooks/([^/]+)/', str(url))
    if m:
        key = m.group(1).lower()
        return PLAYBOOK_MAP.get(key, key.replace('_', ' ').title())
    return "Accelerate"

def playbook_stats():
    data_dir = SCRIPT_DIR / 'data'
    pattern = re.compile(r'^playbook-monthly-(\d{4}-\d{2})\.xlsx$')
    files = sorted(
        [(m.group(1), p) for p in data_dir.glob('*.xlsx') if (m := pattern.match(p.name))],
        key=lambda x: x[0]
    )
    if not files:
        return {'total_views': '—', 'latest_month': '—', 'top_playbook': '—'}

    frames = []
    for label, path in files:
        df = pd.read_excel(path)
        df['Playbook'] = df['Url'].apply(get_playbook)
        frames.append(df)

    combined = pd.concat(frames, ignore_index=True)
    total_views  = len(combined)
    latest_month = fmt_month(files[-1][0])
    top = combined['Playbook'].value_counts()
    top_playbook = top.index[0] if len(top) else '—'
    unique_reps  = combined['Uid'].nunique() if 'Uid' in combined.columns else '—'

    return {
        'total_views':  f"{total_views:,}",
        'latest_month': latest_month,
        'top_playbook': top_playbook,
        'unique_reps':  unique_reps,
    }

# ── HC Cert stats ─────────────────────────────────────────────────────────────
def hc_cert_stats():
    cert_dir = SCRIPT_DIR / 'cert-data'
    hc_cert_path  = cert_dir / 'Healthcare=Certification-Report-07.06.2026.xlsx'
    hc_learn_path = cert_dir / 'Healthcare-Certification-Foundations-Curricula-Report-07.06.2026.xlsx'

    if hc_cert_path.exists() and hc_learn_path.exists():
        from update_cert_dashboard import load_rows_healthcare_v2, TLG as CT
        rows        = load_rows_healthcare_v2(str(hc_cert_path), str(hc_learn_path))
        non_tlg     = [r for r in rows if f"{r['FirstName']} {r['LastName']}" not in CT]
        total       = len(non_tlg)
        certified   = sum(1 for r in non_tlg if r['Certified'] == 'Yes')
        in_progress = sum(1 for r in non_tlg if r['Certified'] != 'Yes' and r.get('overallDone', 0) > 0)
        not_started = total - certified - in_progress
        avg_pct     = round(sum(r.get('overallPct', 0) for r in non_tlg) / total) if total else 0
        rate        = round(certified / total * 100) if total else 0
        return {'total': total, 'certified': certified, 'rate': rate,
                'in_progress': in_progress, 'not_started': not_started, 'avg_pct': avg_pct}

    # Fallback: old single-file format
    from update_cert_dashboard import load_rows, extract_file_date, person_key, TLG as CT
    files = sorted(
        [f for f in os.listdir(cert_dir) if re.search(r'(cert-healthcare|FY\d+-Healthcare)', f, re.I)],
        key=extract_file_date
    )
    if not files:
        return None
    all_rows = []
    for fname in files:
        rows = load_rows(str(cert_dir / fname))
        fd = extract_file_date(fname)
        for r in rows:
            r['_fd'] = fd
        all_rows.extend(rows)
    seen = {}
    for r in sorted(all_rows, key=lambda r: r['_fd']):
        seen[person_key(r)] = r
    deduped   = [r for r in seen.values() if f"{r['FirstName']} {r['LastName']}" not in CT]
    total     = len(deduped)
    certified = sum(1 for r in deduped if r['Complete'] == 'Yes')
    rate      = round(certified / total * 100) if total else 0
    return {'total': total, 'certified': certified, 'rate': rate,
            'in_progress': 0, 'not_started': total - certified, 'avg_pct': rate}

# ── PS Cert stats ─────────────────────────────────────────────────────────────
# PS learning file: update filename below when Resmie provides the learning report.
# Once the file exists in cert-data/, in_progress will populate automatically.
PS_LEARN_FILE = 'Public Sector Foundations Learning Report.xlsx'

def ps_cert_stats():
    cert_dir = SCRIPT_DIR / 'cert-data'
    from update_cert_dashboard import load_rows_publicsector, extract_file_date, person_key
    files = sorted(
        [f for f in os.listdir(cert_dir)
         if f.endswith('.xlsx') and re.search(r'(cert-publicsector|FY\d+-PublicSector)', f, re.I)],
        key=extract_file_date
    )
    if not files:
        return None
    all_rows = []
    for fname in files:
        rows = load_rows_publicsector(str(cert_dir / fname))
        fd = extract_file_date(fname)
        for r in rows:
            r['_fd'] = fd
        all_rows.extend(rows)
    seen = {}
    for r in sorted(all_rows, key=lambda r: r['_fd']):
        seen[person_key(r)] = r
    deduped   = list(seen.values())
    total     = len(deduped)
    certified = sum(1 for r in deduped if r['PublicSector'] == 'Yes')
    rate      = round(certified / total * 100) if total else 0

    # In-progress breakdown requires a PS learning file (not yet available).
    # When PS_LEARN_FILE exists in cert-data/, add parsing logic here similar to HC v2.
    ps_learn_path = cert_dir / PS_LEARN_FILE
    if ps_learn_path.exists():
        # TODO: parse ps_learn_path to compute in_progress and not_started
        in_progress = 0   # replace with real count once parsing is implemented
        not_started = total - certified - in_progress
    else:
        in_progress = 0
        not_started = total - certified

    return {'total': total, 'certified': certified, 'rate': rate,
            'in_progress': in_progress, 'not_started': not_started}

# ── Onboarding stats ──────────────────────────────────────────────────────────
def onboarding_stats():
    from update_onboarding_dashboard import load_lms
    orig = os.getcwd()
    try:
        os.chdir(str(SCRIPT_DIR))
        records = load_lms()  # dict: email -> record
    finally:
        os.chdir(orig)
    people    = list(records.values())
    total     = len(people)
    completed = sum(1 for p in people if p.get('overallDone'))
    overdue   = sum(
        1 for p in people
        if not p.get('overallDone') and any(
            c.get('daysRem') is not None and c['daysRem'] < 0 and not c.get('complete')
            for c in p.get('curricula', {}).values()
        )
    )
    avg_pct = round(sum(p.get('overallPct', 0) for p in people) / total) if total else 0
    return {'total': total, 'completed': completed, 'overdue': overdue, 'avg_pct': avg_pct}

# ── Leaderboard stats ─────────────────────────────────────────────────────────
def leaderboard_stats():
    orig = os.getcwd()
    try:
        os.chdir(str(SCRIPT_DIR))
        from update_leaderboard_dashboard import (
            _find_files, _parse_lms, _parse_html_xls, _stage_index, _build_data, WINDOW_DAYS
        )
        lms_path, cw_path, sh_path = _find_files()
        hires, cohort_start = _parse_lms(lms_path)
        cw_rows = _parse_html_xls(cw_path)
        sh_rows = _parse_html_xls(sh_path)
        stage_idx = _stage_index(sh_rows)
        deals, _ = _build_data(cw_rows, stage_idx, hires, cohort_start)
        from datetime import date as _d
        today = _d.today()
        lb_rows = [
            d for d in deals
            if d.get('assignToCloseDays') is not None and
               0 <= d['assignToCloseDays'] <= WINDOW_DAYS and
               d['curriculumComplete'] == 'Yes' and
               d['salesQualifiedBy'] == d['name'] and
               d['engageBy'] == d['name']
        ]
        in_window = sum(1 for h in hires if h.get('assignDate') and
                        0 <= (_d.today() - _d.fromisoformat(h['assignDate'])).days <= WINDOW_DAYS)
        total_rev   = sum(d['amount'] for d in lb_rows)
        on_board    = len(set(d['name'] for d in lb_rows))
        return {
            'total':      len(hires),
            'in_window':  in_window,
            'on_board':   on_board,
            'total_rev':  f"${total_rev:,.0f}",
        }
    except Exception as e:
        print(f"    Leaderboard stats error: {e}")
        return None
    finally:
        os.chdir(orig)


# ── HTML generation ───────────────────────────────────────────────────────────
def generate_html(pb, hc, ps, ob, lb=None):
    today = datetime.now().strftime('%B %-d, %Y')

    pb_views = pb['total_views']   if pb else '—'
    pb_month = pb['latest_month']  if pb else '—'
    pb_top   = pb['top_playbook']  if pb else '—'
    pb_reps  = pb['unique_reps']   if pb else '—'

    hc_total       = hc['total']       if hc else '—'
    hc_certified   = hc['certified']   if hc else '—'
    hc_rate        = hc['rate']        if hc else '—'
    hc_in_progress = hc['in_progress'] if hc else '—'
    hc_not_started = hc['not_started'] if hc else '—'
    hc_avg_pct     = hc['avg_pct']     if hc else '—'

    ps_total       = ps['total']       if ps else '—'
    ps_completed   = ps['certified']   if ps else '—'
    ps_in_progress = ps['in_progress'] if ps else 0
    ps_not_started = ps['not_started'] if ps else '—'
    ps_rate        = ps['rate']        if ps else '—'
    ps_inprog_pill = (f'<span class="pill pill-blue">&#9679; {ps_in_progress} in progress</span>'
                      if ps_in_progress and ps_in_progress > 0 else '')

    ob_total     = ob['total']     if ob else '—'
    ob_completed = ob['completed'] if ob else '—'
    ob_overdue   = ob['overdue']   if ob else '—'
    ob_avg_pct   = ob['avg_pct']   if ob else '—'

    ob_on_track  = (ob['total'] - ob['completed'] - ob['overdue']) if ob else '—'

    lb_total     = lb['total']     if lb else '—'
    lb_in_window = lb['in_window'] if lb else '—'
    lb_on_board  = lb['on_board']  if lb else '—'
    lb_total_rev = lb['total_rev'] if lb else '—'

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Analytical Data Hub</title>
<style>
  :root{{
    --bg:#0f1117;--surface:#1a1d27;--surface2:#22263a;--border:#2e3347;
    --text:#e0e6f0;--muted:#8891aa;--accent:#4a7cf7;
    --green:#22c55e;--red:#ef4444;
    --font:'Inter',system-ui,sans-serif;
  }}
  body.light-mode{{
    --bg:#f4f7fb;--surface:#fff;--surface2:#f0f4ff;--border:#dde4f0;
    --text:#1a1e2e;--muted:#6b7a99;
  }}
  *{{box-sizing:border-box;margin:0;padding:0;}}
  body{{background:var(--bg);color:var(--text);font-family:var(--font);min-height:100vh;transition:background .2s,color .2s;}}

  .header{{padding:20px 32px;border-bottom:1px solid var(--border);background:var(--surface);display:grid;grid-template-columns:1fr auto 1fr;align-items:center;gap:12px;}}
  .header-left h1{{font-size:18px;font-weight:700;letter-spacing:.3px;}}
  .header-sub{{font-size:12px;color:var(--muted);margin-top:3px;}}
  .header-center{{display:flex;justify-content:center;}}
  .header-right{{display:flex;justify-content:flex-end;}}
  .btn-theme{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:5px 12px;font-size:12px;cursor:pointer;transition:all .15s;}}
  .btn-theme:hover{{border-color:var(--accent);color:var(--text);}}
  .kma-logo{{height:38px;width:auto;display:block;}}
  .kma-logo-light{{display:none;}}
  .light-mode .kma-logo-dark{{display:none;}}
  .light-mode .kma-logo-light{{display:block;}}

  .main{{padding:32px;max-width:1100px;margin:0 auto;}}
  .grid{{display:grid;grid-template-columns:repeat(2,1fr);gap:20px;}}
  @media(max-width:680px){{.grid{{grid-template-columns:1fr;}}}}

  .card{{background-color:var(--surface);background-image:linear-gradient(115deg,rgba(26,29,39,1.0) 0%,rgba(26,29,39,0.85) 40%,rgba(26,29,39,0.35) 100%),var(--card-bg,none);background-size:auto,cover;background-position:center;border:1px solid var(--border);border-radius:12px;padding:24px;display:flex;flex-direction:column;gap:16px;transition:border-color .15s;position:relative;overflow:hidden;min-height:240px;}}
  .card:hover{{border-color:var(--accent);}}
  .light-mode .card{{background-image:linear-gradient(115deg,rgba(255,255,255,1.0) 0%,rgba(255,255,255,0.85) 40%,rgba(255,255,255,0.32) 100%),var(--card-bg,none);}}
  .card-head{{display:flex;align-items:center;gap:12px;}}
  .card-icon{{font-size:22px;line-height:1;}}
  .card-title{{font-size:16px;font-weight:700;}}
  .card-desc{{font-size:13px;color:var(--muted);margin-top:2px;font-weight:500;}}

  .stat-main{{display:flex;align-items:baseline;gap:8px;}}
  .stat-num{{font-size:34px;font-weight:800;line-height:1;font-variant-numeric:tabular-nums;}}
  .stat-unit{{font-size:14px;color:var(--muted);font-weight:600;}}
  .stat-sub{{font-size:13px;color:var(--muted);margin-top:4px;font-weight:500;}}
  .stat-row{{display:flex;align-items:center;gap:10px;flex-wrap:wrap;}}

  .pill{{display:inline-flex;align-items:center;gap:4px;padding:3px 10px;border-radius:99px;font-size:12px;font-weight:600;}}
  .pill-green{{background:rgba(34,197,94,0.22);color:var(--green);}}
  .pill-red{{background:rgba(239,68,68,0.22);color:var(--red);}}
  .pill-blue{{background:rgba(74,124,247,0.22);color:var(--accent);}}
  .pill-muted{{background:rgba(136,145,170,0.22);color:var(--muted);}}

  .card-footer{{margin-top:auto;display:flex;justify-content:flex-end;}}
  .btn-open{{display:inline-flex;align-items:center;gap:6px;background:var(--accent);color:#fff;border:none;border-radius:8px;padding:10px 20px;font-size:13px;font-weight:600;cursor:pointer;text-decoration:none;transition:opacity .15s;}}
  .btn-open:hover{{opacity:.88;}}

  .divider{{width:100%;height:1px;background:var(--border);margin:4px 0;}}

  /* ── Visual effects ──────────────────────────────────────────── */
  #particles-bg{{position:fixed;inset:0;z-index:0;pointer-events:none;}}
  .light-mode #particles-bg{{opacity:.45;}}
  body>*:not(#particles-bg){{position:relative;z-index:1;}}
  .header{{position:relative;overflow:hidden;}}
  .header>*:not(.aurora-wrap){{position:relative;z-index:1;}}
  .aurora-wrap{{position:absolute;inset:0;overflow:hidden;pointer-events:none;}}
  .aurora-orb{{position:absolute;border-radius:50%;filter:blur(70px);animation:aurora-drift linear infinite alternate;}}
  .aurora-1{{width:400px;height:210px;background:rgba(74,124,247,.4);top:-120px;left:1%;animation-duration:11s;}}
  .aurora-2{{width:300px;height:175px;background:rgba(139,92,246,.32);top:-85px;left:40%;animation-duration:8.5s;animation-delay:-2.5s;}}
  .aurora-3{{width:270px;height:160px;background:rgba(20,184,166,.27);top:-100px;right:3%;animation-duration:14s;animation-delay:-5.5s;}}
  .light-mode .aurora-1{{background:rgba(74,124,247,.22);top:-80px;}}
  .light-mode .aurora-2{{background:rgba(139,92,246,.18);top:-60px;}}
  .light-mode .aurora-3{{background:rgba(20,184,166,.16);top:-70px;}}
  @keyframes aurora-drift{{0%{{transform:translate(0,0) scale(1);}}100%{{transform:translate(30px,22px) scale(1.22);}}}}
</style>
</head>
<body>
<canvas id="particles-bg"></canvas>

<div class="header">
  <div class="aurora-wrap">
    <div class="aurora-orb aurora-1"></div>
    <div class="aurora-orb aurora-2"></div>
    <div class="aurora-orb aurora-3"></div>
  </div>
  <div class="header-left">
    <h1>Analytical Data Hub</h1>
    <div class="header-sub">Powered by TLG &middot; {today}</div>
  </div>
  <div class="header-center">
    <img src="KMA-wht.svg" class="kma-logo kma-logo-dark" alt="KM Academy">
    <img src="KMA-drk.svg" class="kma-logo kma-logo-light" alt="KM Academy">
  </div>
  <div class="header-right">
    <button class="btn-theme" id="btn-theme" onclick="toggleTheme()">&#9728; Light</button>
  </div>
</div>

<div class="main">
  <div class="grid">

    <!-- Accelerate Onboarding -->
    <div class="card" style="--card-bg:url('https://jasonackerman1.github.io/accelerate_sales_playbook/img/accelerateHero.jpg');background-size:auto,160%;background-position:center,calc(50% - 20px) calc(50% + 20px);">
      <div class="card-head">
        <span class="card-icon">&#127919;</span>
        <div>
          <div class="card-title">Accelerate Onboarding</div>
          <div class="card-desc">45-day new rep program</div>
        </div>
      </div>
      <div>
        <div class="stat-main">
          <span class="stat-num" style="color:var(--accent)">{ob_total}</span>
          <span class="stat-unit">total learners</span>
        </div>
        <div class="stat-sub">{ob_avg_pct}% avg completion</div>
      </div>
      <div style="margin-top:auto;display:flex;align-items:flex-end;justify-content:space-between;gap:16px;">
        <div class="stat-row" style="flex:1;">
          <span class="pill pill-green">&#10003; {ob_completed} completed</span>
          <span class="pill pill-blue">&#9679; {ob_on_track} on track</span>
          <div style="width:100%;height:0;"></div>
          <span class="pill pill-red">&#9888; {ob_overdue} overdue</span>
        </div>
        <a href="onboarding.html" class="btn-open" style="flex-shrink:0;">Go to Dashboard &#8250;</a>
      </div>
    </div>

    <!-- Accelerate Leaderboard -->
    <div class="card" style="--card-bg:url('https://jasonackerman1.github.io/accelerate_sales_playbook/img/accelerateHero.jpg');background-size:auto,160%;background-position:center,calc(50% - 20px) calc(50% + 20px);">
      <div class="card-head">
        <span class="card-icon">&#127942;</span>
        <div>
          <div class="card-title">Accelerate Leaderboard</div>
          <div class="card-desc">First closed-won deals, 45-day window</div>
        </div>
      </div>
      <div>
        <div class="stat-main">
          <span class="stat-num" style="color:var(--accent)">{lb_total}</span>
          <span class="stat-unit">cohort members</span>
        </div>
        <div class="stat-sub">{lb_in_window} currently in 45-day window</div>
      </div>
      <div style="margin-top:auto;display:flex;align-items:flex-end;justify-content:space-between;gap:16px;">
        <div class="stat-row" style="flex:1;">
          <span class="pill pill-green">&#9733; {lb_on_board} on the board</span>
          <div style="width:100%;height:0;"></div>
          <span class="pill pill-blue">&#36; {lb_total_rev} qualifying revenue</span>
        </div>
        <a href="leaderboard.html" class="btn-open" style="flex-shrink:0;">Go to Dashboard &#8250;</a>
      </div>
    </div>

    <!-- Healthcare Certification -->
    <div class="card" style="--card-bg:url('https://cdn.jsdelivr.net/gh/BradleyAPierce/Healthcare_Playbook_Images/HealthcareHomeReduced.jpg')">
      <div class="card-head">
        <span class="card-icon">&#127973;</span>
        <div>
          <div class="card-title">Healthcare Certification</div>
          <div class="card-desc">LMS certification progress</div>
        </div>
      </div>
      <div>
        <div class="stat-main">
          <span class="stat-num" style="color:var(--accent)">{hc_total}</span>
          <span class="stat-unit">total learners</span>
        </div>
        <div class="stat-sub">{hc_avg_pct}% avg completion</div>
      </div>
      <div style="margin-top:auto;display:flex;align-items:flex-end;justify-content:space-between;gap:16px;">
        <div class="stat-row" style="flex:1;">
          <span class="pill pill-green">&#10003; {hc_certified} certified</span>
          <span class="pill pill-blue">&#9679; {hc_in_progress} in progress</span>
          <div style="width:100%;height:0;"></div>
          <span class="pill pill-muted">&#9675; {hc_not_started} not started</span>
        </div>
        <a href="cert-healthcare.html" class="btn-open" style="flex-shrink:0;">Go to Dashboard &#8250;</a>
      </div>
    </div>

    <!-- Public Sector Curriculum -->
    <div class="card" style="--card-bg:url('https://cdn.jsdelivr.net/gh/BradleyAPierce/Public_Sector_Playbook_Images/publicSectorHero3.jpg')">
      <div class="card-head">
        <span class="card-icon">&#127963;</span>
        <div>
          <div class="card-title">Public Sector Curriculum</div>
          <div class="card-desc">LMS curriculum progress</div>
        </div>
      </div>
      <div>
        <div class="stat-main">
          <span class="stat-num" style="color:var(--accent)">{ps_total}</span>
          <span class="stat-unit">total learners</span>
        </div>
        <div class="stat-sub">{ps_rate}% completion rate</div>
      </div>
      <div style="margin-top:auto;display:flex;align-items:flex-end;justify-content:space-between;gap:16px;">
        <div class="stat-row" style="flex:1;">
          <span class="pill pill-green">&#10003; {ps_completed} completed</span>
          {ps_inprog_pill}
          <div style="width:100%;height:0;"></div>
          <span class="pill pill-muted">&#9675; {ps_not_started} not started</span>
        </div>
        <a href="cert-publicsector.html" class="btn-open" style="flex-shrink:0;">Go to Dashboard &#8250;</a>
      </div>
    </div>

    <!-- Playbook Traffic -->
    <div class="card" style="--card-bg:url('https://cdn.jsdelivr.net/gh/BradleyAPierce/Legal_Images_Copy/Sales_Education.jpg')">
      <div class="card-head">
        <span class="card-icon">&#128202;</span>
        <div>
          <div class="card-title">Playbook Traffic</div>
          <div class="card-desc">Monthly page views by rep</div>
        </div>
      </div>
      <div>
        <div class="stat-main">
          <span class="stat-num" style="color:var(--accent)">{pb_views}</span>
          <span class="stat-unit">total views</span>
        </div>
        <div class="stat-sub">Latest data: {pb_month}</div>
      </div>
      <div style="margin-top:auto;display:flex;align-items:center;justify-content:space-between;gap:16px;">
        <div style="flex:1;display:flex;flex-direction:column;gap:8px;">
          <div class="stat-row">
            <span style="font-size:12px;color:var(--muted)">Top playbook:</span>
            <span class="pill pill-blue">{pb_top}</span>
          </div>
          <div class="stat-row">
            <span style="font-size:12px;color:var(--muted)">Unique reps:</span>
            <span class="pill pill-blue">{pb_reps}</span>
          </div>
        </div>
        <a href="playbook.html" class="btn-open" style="flex-shrink:0;">Go to Dashboard &#8250;</a>
      </div>
    </div>

  </div>
</div>

<script>
(function(){{
  if(localStorage.getItem('pb-theme')==='light'){{
    document.body.classList.add('light-mode');
    document.getElementById('btn-theme').innerHTML='&#9790; Dark';
  }}
}})();
function toggleTheme(){{
  var light=document.body.classList.toggle('light-mode');
  document.getElementById('btn-theme').innerHTML=light?'&#9790; Dark':'&#9728; Light';
  localStorage.setItem('pb-theme',light?'light':'dark');
}}

/* ── Particle network ───────────────────────────────────────────── */
(function(){{
  var cvs=document.getElementById('particles-bg'),ctx=cvs.getContext('2d');
  var pts=[],N=65,DIST=130;
  function resize(){{cvs.width=innerWidth;cvs.height=innerHeight;}}
  addEventListener('resize',resize);resize();
  for(var i=0;i<N;i++)pts.push({{
    x:Math.random()*cvs.width,y:Math.random()*cvs.height,
    vx:(Math.random()-.5)*.45,vy:(Math.random()-.5)*.45,
    r:Math.random()*1.6+.7
  }});
  function frame(){{
    ctx.clearRect(0,0,cvs.width,cvs.height);
    var dark=!document.body.classList.contains('light-mode');
    pts.forEach(function(p){{
      p.x+=p.vx;p.y+=p.vy;
      if(p.x<0||p.x>cvs.width)p.vx*=-1;
      if(p.y<0||p.y>cvs.height)p.vy*=-1;
      ctx.beginPath();ctx.arc(p.x,p.y,p.r,0,6.283);
      ctx.fillStyle=dark?'rgba(74,124,247,.55)':'rgba(74,124,247,.28)';
      ctx.fill();
    }});
    for(var i=0;i<N;i++)for(var j=i+1;j<N;j++){{
      var dx=pts[i].x-pts[j].x,dy=pts[i].y-pts[j].y,d=Math.sqrt(dx*dx+dy*dy);
      if(d<DIST){{
        ctx.beginPath();ctx.moveTo(pts[i].x,pts[i].y);ctx.lineTo(pts[j].x,pts[j].y);
        var a=(dark?.18:.07)*(1-d/DIST);
        ctx.strokeStyle='rgba(74,124,247,'+a+')';ctx.lineWidth=.7;ctx.stroke();
      }}
    }}
    requestAnimationFrame(frame);
  }}
  frame();
}})();

/* ── Stat counters ──────────────────────────────────────────────── */
document.addEventListener('DOMContentLoaded',function(){{
  document.querySelectorAll('.stat-num').forEach(function(el,i){{
    var orig=el.textContent.trim();
    if(orig==='—')return;
    var pre=orig[0]==='$'?'$':'';
    var suf=orig[orig.length-1]==='%'?'%':'';
    var num=parseFloat(orig.replace(/[$,%\s,]/g,''));
    if(isNaN(num)||num===0)return;
    var t0=null,DUR=1400;
    el.textContent=pre+'0'+suf;
    requestAnimationFrame(function tick(ts){{
      if(!t0)t0=ts;
      var p=Math.min((ts-t0)/DUR,1),e=1-Math.pow(1-p,3);
      el.textContent=pre+Math.round(e*num).toLocaleString()+suf;
      if(p<1)requestAnimationFrame(tick);else el.textContent=orig;
    }});
  }});
}});
</script>
</body>
</html>"""


def main():
    print("Generating homepage...")

    print("  Reading playbook traffic data...")
    pb = playbook_stats()
    print(f"    {pb['total_views']} views, latest: {pb['latest_month']}, top: {pb['top_playbook']}")

    print("  Reading Healthcare cert data...")
    hc = hc_cert_stats()
    if hc:
        print(f"    {hc['certified']} of {hc['total']} certified ({hc['rate']}%)")
    else:
        print("    No HC data found")

    print("  Reading Public Sector cert data...")
    ps = ps_cert_stats()
    if ps:
        print(f"    {ps['certified']} of {ps['total']} certified ({ps['rate']}%)")
    else:
        print("    No PS data found")

    print("  Reading onboarding data...")
    ob = onboarding_stats()
    if ob:
        print(f"    {ob['total']} learners, {ob['completed']} completed, {ob['overdue']} overdue")
    else:
        print("    No onboarding data found")

    print("  Reading leaderboard data...")
    lb = leaderboard_stats()
    if lb:
        print(f"    {lb['total']} cohort members, {lb['in_window']} in window, {lb['on_board']} on the board, {lb['total_rev']} qualifying revenue")
    else:
        print("    No leaderboard data found")

    html = generate_html(pb, hc, ps, ob, lb)
    out  = SCRIPT_DIR / 'index.html'
    out.write_text(html, encoding='utf-8')
    print(f"\nHomepage written to: {out}")


if __name__ == '__main__':
    main()
