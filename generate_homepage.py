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
    hc_cert_path  = cert_dir / 'Healthcare Certification Learning Report.xlsx'
    hc_learn_path = cert_dir / 'Healthcare Foundations for Direct Sales Learning Report.xlsx'

    if hc_cert_path.exists() and hc_learn_path.exists():
        from update_cert_dashboard import load_rows_healthcare_v2, TLG as CT
        rows      = load_rows_healthcare_v2(str(hc_cert_path), str(hc_learn_path))
        non_tlg   = [r for r in rows if f"{r['FirstName']} {r['LastName']}" not in CT]
        total     = len(non_tlg)
        certified = sum(1 for r in non_tlg if r['Certified'] == 'Yes')
        rate      = round(certified / total * 100) if total else 0
        return {'total': total, 'certified': certified, 'rate': rate}

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
    return {'total': total, 'certified': certified, 'rate': rate}

# ── PS Cert stats ─────────────────────────────────────────────────────────────
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
    return {'total': total, 'certified': certified, 'rate': rate}

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

# ── HTML generation ───────────────────────────────────────────────────────────
def generate_html(pb, hc, ps, ob):
    today = datetime.now().strftime('%B %-d, %Y')

    pb_views = pb['total_views']   if pb else '—'
    pb_month = pb['latest_month']  if pb else '—'
    pb_top   = pb['top_playbook']  if pb else '—'
    pb_reps  = pb['unique_reps']   if pb else '—'

    hc_total     = hc['total']     if hc else '—'
    hc_certified = hc['certified'] if hc else '—'
    hc_rate      = hc['rate']      if hc else '—'

    ps_total     = ps['total']     if ps else '—'
    ps_certified = ps['certified'] if ps else '—'
    ps_rate      = ps['rate']      if ps else '—'

    ob_total     = ob['total']     if ob else '—'
    ob_completed = ob['completed'] if ob else '—'
    ob_overdue   = ob['overdue']   if ob else '—'
    ob_avg_pct   = ob['avg_pct']   if ob else '—'

    hc_not_yet   = (hc['total'] - hc['certified']) if hc else '—'
    ps_not_yet   = (ps['total'] - ps['certified']) if ps else '—'
    ob_on_track  = (ob['total'] - ob['completed'] - ob['overdue']) if ob else '—'

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

  .card{{background-color:var(--surface);background-image:linear-gradient(115deg,rgba(26,29,39,1.0) 0%,rgba(26,29,39,0.85) 40%,rgba(26,29,39,0.35) 100%),var(--card-bg,none);background-size:auto,cover;background-position:center;border:1px solid var(--border);border-radius:12px;padding:24px;display:flex;flex-direction:column;gap:16px;transition:border-color .15s;position:relative;overflow:hidden;}}
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
  .pill-green{{background:rgba(34,197,94,0.80);color:#fff;}}
  .pill-red{{background:rgba(239,68,68,0.80);color:#fff;}}
  .pill-blue{{background:rgba(74,124,247,0.80);color:#fff;}}

  .card-footer{{margin-top:auto;}}
  .btn-open{{display:inline-flex;align-items:center;gap:6px;background:var(--accent);color:#fff;border:none;border-radius:8px;padding:10px 20px;font-size:13px;font-weight:600;cursor:pointer;text-decoration:none;transition:opacity .15s;}}
  .btn-open:hover{{opacity:.88;}}

  .divider{{width:100%;height:1px;background:var(--border);margin:4px 0;}}
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <h1>Analytical Data Hub</h1>
    <div class="header-sub">Powered by TLG &middot; {today}</div>
  </div>
  <div class="header-center">
    <img src="https://cdn.jsdelivr.net/gh/BradleyAPierce/RTDX_Images/KMA-wht.svg" class="kma-logo kma-logo-dark" alt="KM Academy">
    <img src="KMA-drk.svg" class="kma-logo kma-logo-light" alt="KM Academy">
  </div>
  <div class="header-right">
    <button class="btn-theme" id="btn-theme" onclick="toggleTheme()">&#9728; Light</button>
  </div>
</div>

<div class="main">
  <div class="grid">

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
      <div class="stat-row">
        <span style="font-size:12px;color:var(--muted)">Top:</span>
        <span class="pill pill-blue">{pb_top}</span>
        <span style="font-size:12px;color:var(--muted);margin-left:4px;">Reps:</span>
        <span class="pill pill-blue">{pb_reps}</span>
      </div>
      <div class="card-footer">
        <a href="playbook.html" class="btn-open">Open Dashboard &#8250;</a>
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
          <span class="stat-num" style="color:var(--green)">{hc_certified}</span>
          <span class="stat-unit">of {hc_total} certified</span>
        </div>
        <div class="stat-sub">{hc_rate}% completion rate</div>
      </div>
      <div class="stat-row">
        <span class="pill pill-red">&#9679; {hc_not_yet} not yet completed</span>
      </div>
      <div class="card-footer">
        <a href="cert-healthcare.html" class="btn-open">Open Dashboard &#8250;</a>
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
          <span class="stat-num" style="color:var(--green)">{ps_certified}</span>
          <span class="stat-unit">of {ps_total} comp</span>
        </div>
        <div class="stat-sub">{ps_rate}% completion rate</div>
      </div>
      <div class="stat-row">
        <span class="pill pill-red">&#9679; {ps_not_yet} not yet completed</span>
      </div>
      <div class="card-footer">
        <a href="cert-publicsector.html" class="btn-open">Open Dashboard &#8250;</a>
      </div>
    </div>

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
      <div class="stat-row">
        <span class="pill pill-green">&#10003; {ob_completed} completed</span>
        <span class="pill pill-blue">&#9679; {ob_on_track} on track</span>
        <span class="pill pill-red">&#9888; {ob_overdue} overdue</span>
      </div>
      <div class="card-footer">
        <a href="onboarding.html" class="btn-open">Open Dashboard &#8250;</a>
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

    html = generate_html(pb, hc, ps, ob)
    out  = SCRIPT_DIR / 'index.html'
    out.write_text(html, encoding='utf-8')
    print(f"\nHomepage written to: {out}")


if __name__ == '__main__':
    main()
