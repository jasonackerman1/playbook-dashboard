#!/usr/bin/env python3
"""
update_onboarding_dashboard.py
Generates onboarding.html from:
  - onboarding-data/*.xlsx  (LMS Accelerate item-level data, one row per course item per person)
  - data/*.xlsx             (playbook traffic, filtered for Accelerate URLs only)
"""

import openpyxl
import os
import re
import json
import warnings
warnings.filterwarnings('ignore')

# ── LMS column indices (0-based) ─────────────────────────────────────────────
COL_FIRST      = 2   # First Name
COL_LAST       = 3   # Last Name
COL_EMAIL      = 4   # Email Address
COL_JOBTITLE   = 5   # Job Title
COL_MARKET     = 7   # Market
COL_MGR_FIRST  = 9   # ManagerFirstName
COL_MGR_LAST   = 10  # ManagerLastName
COL_MGR_EMAIL  = 11  # SUPEMAILADDR
COL_MGR_TITLE  = 12  # Manager JobTitle
COL_HIRE_DATE  = 14  # Hire Date
COL_CURRIC_ID  = 17  # Curriculum ID
COL_CURRIC_TTL = 18  # Curriculum Title
COL_CURRIC_CMP = 19  # Curriculum Complete (Yes/No)
COL_ASSIGN_DT  = 20  # Curriculum Assignment Date
COL_DAYS_REM   = 21  # Days Remaining (LMS-computed per curriculum deadline)
COL_ITEM_TTL   = 26  # Item Title
COL_ITEM_TYPE  = 23  # Item Type (ONLINE/VILT)
COL_ITEM_DATE  = 27  # Item Completion Date
COL_ITEM_STS   = 29  # Item Completion Status Description

# ── Playbook traffic columns ──────────────────────────────────────────────────
PB_FIRST = 1
PB_LAST  = 2
PB_DATE  = 9
PB_URL   = 10

# ── Sub-curricula in display order ───────────────────────────────────────────
CURRICULA = [
    ('ACCELERATE_GS', 'Getting Started'),
    ('ACCELERATE_SW', 'Sales Workflow'),
    ('ACCELERATE_CP', 'Core Portfolio'),
    ('ACCELERATE_P',  'Prospecting'),
    ('ACCELERATE_SS', 'Sales Skills'),
    ('ACCELERATE_PM', 'Pipeline Mgmt'),
]

PROGRAM_DAYS = 45

TLG = {
    'jason ackerman', 'bianca davis', 'james parker', 'resmie biba',
    'chris curtis', 'sara thompson', 'jeremy macbean', 'bradley pierce',
    'laura sefcik', 'samantha maresca', 'staci musco', 'cj homer',
    'rich moore', 'dale kinsey',
}

PAGE_LABELS = {
    'coreportfolio':         'Core Portfolio',
    'salesworkflow':         'Sales Workflow',
    'overview':              'Overview',
    'welcome':               'Welcome',
    'resources':             'Resources',
    'understandingsalesforce': 'Understanding Salesforce',
    'managers':              'Managers',
    'index':                 'Home',
    'prospectingskills':     'Prospecting Skills',
    'callprep':              'Call Prep',
    'workingwithnumbers':    'Working With Numbers',
}


def pkey(first, last):
    return f"{str(first).strip().lower()} {str(last).strip().lower()}"


def extract_date(fname):
    m = re.search(r'(\d{4}-\d{2})', os.path.basename(fname))
    return m.group(1) if m else '0000-00'


def page_label(url):
    m = re.search(r'/([^/]+?)(?:\.html)?$', url.rstrip('/'))
    if not m or m.group(1) in ('accelerate_sales_playbook', ''):
        return 'Home'
    slug = m.group(1).lower()
    return PAGE_LABELS.get(slug, slug.replace('-', ' ').title())


def load_lms():
    folder = 'onboarding-data'
    files = sorted(
        [f for f in os.listdir(folder) if f.endswith('.xlsx')],
        key=extract_date
    )
    seen = {}  # email -> record (latest file wins)

    for fname in files:
        wb = openpyxl.load_workbook(f"{folder}/{fname}", read_only=True, data_only=True)
        ws = wb.active
        all_rows = list(ws.iter_rows(values_only=True))
        if len(all_rows) < 2:
            continue
        data = all_rows[1:]

        from collections import defaultdict
        by_email = defaultdict(list)
        for r in data:
            email = str(r[COL_EMAIL] or '').strip().lower()
            if email:
                by_email[email].append(r)

        for email, prows in by_email.items():
            first = str(prows[0][COL_FIRST] or '').strip()
            last  = str(prows[0][COL_LAST]  or '').strip()

            if pkey(first, last) in TLG:
                continue

            # Parent curriculum row (overall program)
            parent = next((r for r in prows if r[COL_CURRIC_ID] == 'ACCELERATE'), None)
            assign_date  = parent[COL_ASSIGN_DT]  if parent else None
            overall_done = (parent[COL_CURRIC_CMP] == 'Yes') if parent else False
            days_rem_lms = int(parent[COL_DAYS_REM]) if parent and parent[COL_DAYS_REM] is not None else None

            # Per sub-curriculum data
            curricula_data = {}
            for cid, cname in CURRICULA:
                crows = [r for r in prows if r[COL_CURRIC_ID] == cid]
                if not crows:
                    continue

                days_rem       = crows[0][COL_DAYS_REM]
                curric_done    = crows[0][COL_CURRIC_CMP] == 'Yes'
                item_rows      = [r for r in crows if r[COL_ITEM_TTL]]

                items = []
                for ir in item_rows:
                    title = str(ir[COL_ITEM_TTL] or '')
                    if 'coming soon' in title.lower():
                        continue
                    items.append({
                        'title': title,
                        'type':  str(ir[COL_ITEM_TYPE] or 'ONLINE'),
                        'done':  bool(ir[COL_ITEM_STS]),
                        'date':  ir[COL_ITEM_DATE].strftime('%Y-%m-%d') if ir[COL_ITEM_DATE] else None,
                    })

                total = len(items)
                done  = sum(1 for i in items if i['done'])
                pct   = round(done / total * 100) if total else 0

                curricula_data[cid] = {
                    'title':    cname,
                    'complete': curric_done,
                    'daysRem':  days_rem,
                    'pct':      pct,
                    'done':     done,
                    'total':    total,
                    'items':    items,
                }

            all_pcts    = [v['pct'] for v in curricula_data.values()]
            overall_pct = round(sum(all_pcts) / len(all_pcts)) if all_pcts else 0

            seen[email] = {
                'email':       email,
                'key':         pkey(first, last),
                'name':        f"{first} {last}",
                'first':       first,
                'last':        last,
                'jobTitle':    str(prows[0][COL_JOBTITLE]  or ''),
                'market':      str(prows[0][COL_MARKET]    or ''),
                'manager':     f"{prows[0][COL_MGR_FIRST] or ''} {prows[0][COL_MGR_LAST] or ''}".strip(),
                'mgrEmail':    str(prows[0][COL_MGR_EMAIL] or ''),
                'mgrTitle':    str(prows[0][COL_MGR_TITLE] or ''),
                'hireDate':    prows[0][COL_HIRE_DATE].strftime('%Y-%m-%d') if prows[0][COL_HIRE_DATE] else None,
                'assignDate':  assign_date.strftime('%Y-%m-%d') if assign_date else None,
                'daysRem':     days_rem_lms,
                'overallDone': overall_done,
                'overallPct':  overall_pct,
                'curricula':   curricula_data,
                'playbook':    [],
            }

    return seen


PB_EMAIL = 3  # playbook traffic email column

def attach_playbook(records):
    folder = 'data'
    if not os.path.exists(folder):
        return records
    files = sorted([f for f in os.listdir(folder) if f.endswith('.xlsx')], key=extract_date)

    for fname in files:
        wb = openpyxl.load_workbook(f"{folder}/{fname}", read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))[1:]

        for r in rows:
            url = str(r[PB_URL] or '')
            if 'accelerate' not in url.lower():
                continue
            pb_email = str(r[PB_EMAIL] or '').strip().lower()
            if pb_email and pb_email in records:
                dt = r[PB_DATE]
                records[pb_email]['playbook'].append({
                    'page': page_label(url),
                    'url':  url,
                    'date': dt.strftime('%Y-%m-%d') if dt else '',
                })

    return records


def generate_html(records):
    from datetime import date
    people      = sorted(records.values(), key=lambda p: p['name'].lower())
    people_json = json.dumps(people, default=str)
    today_str   = date.today().isoformat()
    file_date   = date.today().strftime('%B %d, %Y')
    total       = len(people)

    curric_ids   = json.dumps([c[0] for c in CURRICULA])
    curric_names = json.dumps({c[0]: c[1] for c in CURRICULA})

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Accelerate Onboarding</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root {{
    --bg:#0d1117; --surface:#161b22; --surface2:#1e2530; --border:#30363d;
    --accent:#3b82f6; --accent2:#6d28d9; --accent3:#d97706;
    --text:#e6edf3; --muted:#8b949e; --green:#22c55e; --red:#ef4444;
    --amber:#f59e0b; --teal:#14b8a6;
    --green-subtle:#22c55e22; --red-subtle:#ef444422; --amber-subtle:#f59e0b22;
    --font:'Segoe UI',system-ui,sans-serif;
  }}
  body.light-mode {{
    --bg:#f4f6fb; --surface:#ffffff; --surface2:#eef1f7; --border:#d0d7e8;
    --accent:#2563eb; --text:#1a1d27; --muted:#475569; --green:#059669;
    --red:#dc2626; --amber:#d97706; --teal:#0f766e;
    --green-subtle:#05966922; --red-subtle:#dc262622; --amber-subtle:#d9770622;
  }}
  body.light-mode select,body.light-mode input{{color-scheme:light;}}
  *{{box-sizing:border-box;margin:0;padding:0;}}
  body{{background:var(--bg);color:var(--text);font-family:var(--font);min-height:100vh;transition:background .2s,color .2s;}}

  /* ── Header ── */
  .header{{padding:20px 28px 16px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;}}
  .header-left{{display:flex;align-items:center;gap:16px;}}
  .header-right{{display:flex;align-items:center;gap:8px;}}
  .header h1{{font-size:18px;font-weight:700;letter-spacing:.3px;}}
  .header h1 span{{color:var(--muted);font-weight:400;}}
  .data-badge{{font-size:11px;color:var(--muted);background:var(--surface2);border:1px solid var(--border);border-radius:20px;padding:3px 10px;white-space:nowrap;}}

  /* ── Hamburger ── */
  .hamburger{{position:relative;}}
  .hamburger-btn{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:6px 10px;font-size:16px;cursor:pointer;transition:all .15s;line-height:1;}}
  .hamburger-btn:hover,.hamburger-btn.open{{border-color:var(--accent);color:var(--text);}}
  .hamburger-menu{{position:absolute;top:calc(100% + 6px);left:0;background:var(--surface);border:1px solid var(--border);border-radius:8px;min-width:220px;box-shadow:0 4px 24px rgba(0,0,0,.4);display:none;z-index:200;overflow:hidden;}}
  .hamburger-menu.open{{display:block;}}
  .hamburger-section-label{{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--muted);padding:10px 14px 4px;}}
  .hamburger-item{{display:flex;align-items:center;gap:8px;padding:10px 14px;font-size:13px;color:var(--text);text-decoration:none;transition:background .1s;}}
  .hamburger-item:hover{{background:var(--surface2);}}

  /* ── Buttons ── */
  .btn-theme{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:5px 12px;font-size:12px;cursor:pointer;transition:all .15s;}}
  .btn-theme:hover{{border-color:var(--accent);color:var(--text);}}
  .btn-export{{background:var(--accent);border:1px solid var(--accent);color:#fff;border-radius:6px;padding:5px 12px;font-size:12px;cursor:pointer;font-weight:600;}}
  .btn-export:hover{{opacity:.88;}}
  .export-drop{{position:relative;}}
  .export-menu{{position:absolute;top:calc(100% + 6px);right:0;background:var(--surface);border:1px solid var(--border);border-radius:8px;min-width:210px;box-shadow:0 4px 24px rgba(0,0,0,.28);display:none;z-index:200;overflow:hidden;}}
  .export-menu.open{{display:block;}}
  .export-item{{display:block;width:100%;text-align:left;padding:10px 14px;font-size:13px;color:var(--text);background:transparent;border:none;cursor:pointer;transition:background .1s;font-family:inherit;}}
  .export-item:hover{{background:var(--surface2);}}

  /* ── Info tooltip ── */
  .info-btn{{display:inline-flex;align-items:center;justify-content:center;width:15px;height:15px;border-radius:50%;background:var(--surface2);border:1px solid var(--border);color:var(--muted);font-size:9px;font-weight:700;cursor:pointer;margin-left:5px;vertical-align:middle;flex-shrink:0;line-height:1;transition:border-color .15s,color .15s;}}
  .info-btn:hover{{border-color:var(--accent);color:var(--accent);}}
  .info-popover{{position:fixed;z-index:9999;background:var(--surface2);border:1px solid var(--border);border-radius:8px;padding:12px 14px;font-size:12px;color:var(--text);line-height:1.6;max-width:260px;box-shadow:0 4px 24px rgba(0,0,0,.5);display:none;}}
  .info-popover.visible{{display:block;}}

  /* ── Filters ── */
  .filters{{padding:14px 28px;display:flex;gap:10px;flex-wrap:wrap;align-items:center;border-bottom:1px solid var(--border);background:var(--surface);}}
  .filter-label{{font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.6px;margin-right:4px;}}
  select,input[type=text]{{background:var(--surface2);border:1px solid var(--border);color:var(--text);border-radius:6px;padding:6px 10px;font-size:13px;cursor:pointer;outline:none;color-scheme:dark;}}
  select:focus,input:focus{{border-color:var(--accent);}}
  .btn-reset{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:6px 14px;font-size:12px;cursor:pointer;transition:border-color .15s,color .15s;}}
  .btn-reset:hover{{border-color:var(--accent);color:var(--text);}}
  .btn-tlg{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:6px 14px;font-size:12px;cursor:pointer;transition:all .15s;}}
  .btn-tlg.active{{background:var(--red-subtle);border-color:var(--red);color:var(--red);}}
  .result-count{{margin-left:auto;font-size:12px;color:var(--muted);}}

  /* ── Stats ── */
  .stats{{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;padding:20px 28px;}}
  .stat{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:16px 18px;}}
  .stat-label{{font-size:11px;text-transform:uppercase;letter-spacing:.7px;color:var(--muted);margin-bottom:6px;}}
  .stat-value{{font-size:28px;font-weight:700;line-height:1;}}
  .stat-value.green{{color:var(--green);}}
  .stat-value.red{{color:var(--red);}}
  .stat-value.amber{{color:var(--amber);}}
  .stat-value.blue{{color:var(--accent);}}
  .stat-sub{{font-size:11px;color:var(--muted);margin-top:4px;}}

  /* ── Section / Table ── */
  .section{{padding:0 28px 32px;}}
  .section-header{{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;flex-wrap:wrap;gap:8px;padding-top:20px;}}
  .section-title{{font-size:13px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;}}
  body.light-mode .section-title{{color:var(--text);}}
  .section-hint{{font-size:11px;color:var(--muted);}}
  .table-wrap{{overflow-x:auto;border:1px solid var(--border);border-radius:10px;background:var(--surface);}}
  table{{border-collapse:collapse;width:100%;min-width:900px;}}
  thead th{{font-size:10px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;padding:10px 12px;text-align:left;border-bottom:1px solid var(--border);background:var(--surface2);white-space:nowrap;}}
  thead th.curric-col{{text-align:center;min-width:96px;white-space:normal;line-height:1.3;font-size:10px;}}
  thead th.overall-col{{text-align:center;min-width:80px;}}
  thead th.sortable-th{{cursor:pointer;user-select:none;}}
  thead th.sortable-th:hover{{color:var(--text);}}
  thead th.sort-active{{color:var(--accent) !important;}}
  tbody tr{{border-bottom:1px solid var(--border);cursor:pointer;transition:background .1s;}}
  tbody tr:last-child{{border-bottom:none;}}
  tbody tr:hover td{{background:var(--surface2) !important;}}
  tbody tr.selected td{{background:#3b82f611 !important;}}
  td{{padding:8px 12px;vertical-align:middle;font-size:12px;}}
  td.name-cell{{font-weight:500;white-space:nowrap;display:flex;align-items:center;gap:8px;}}
  td.market-cell{{font-size:11px;color:var(--muted);white-space:nowrap;}}
  td.pct-cell{{text-align:center;padding:5px 4px;}}
  tfoot tr{{border-top:2px solid var(--border);background:var(--surface2);}}
  tfoot td{{padding:7px 12px;font-size:11px;font-weight:600;color:var(--muted);}}
  tfoot td.pct-cell{{text-align:center;}}

  /* ── Pills & Badges ── */
  .pct-pill{{display:inline-block;border-radius:5px;padding:3px 9px;font-size:11px;font-weight:600;min-width:44px;text-align:center;line-height:1.4;}}
  .status-badge{{display:inline-block;font-size:10px;font-weight:700;padding:2px 9px;border-radius:20px;white-space:nowrap;}}
  .sb-completed{{background:var(--green-subtle);color:var(--green);}}
  .sb-ontrack{{background:#3b82f611;color:var(--accent);}}
  .sb-overdue{{background:var(--red-subtle);color:var(--red);}}
  .days-badge{{display:inline-block;font-size:10px;font-weight:600;padding:2px 7px;border-radius:4px;}}

  /* ── Modal overlay ── */
  .modal-overlay{{position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:500;display:none;align-items:flex-start;justify-content:center;padding:40px 16px;overflow-y:auto;}}
  .modal-overlay.open{{display:flex;}}
  .modal-card{{background:var(--surface);border:1px solid var(--border);border-radius:12px;width:100%;max-width:760px;position:relative;}}
  .modal-close{{position:absolute;top:14px;right:16px;background:transparent;border:none;color:var(--muted);font-size:18px;cursor:pointer;line-height:1;padding:4px 8px;border-radius:4px;}}
  .modal-close:hover{{color:var(--text);background:var(--surface2);}}
  .modal-header{{padding:20px 24px 16px;border-bottom:1px solid var(--border);}}
  .modal-name{{font-size:18px;font-weight:700;margin-bottom:4px;}}
  .modal-meta{{font-size:12px;color:var(--muted);display:flex;flex-wrap:wrap;gap:12px;}}
  .modal-body{{padding:20px 24px;}}

  /* ── 45-day progress ── */
  .progress-section{{margin-bottom:20px;}}
  .progress-label{{font-size:12px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px;display:flex;align-items:center;justify-content:space-between;}}
  .progress-bar-wrap{{background:var(--surface2);border-radius:6px;height:10px;overflow:hidden;}}
  .progress-bar-fill{{height:10px;border-radius:6px;transition:width .3s;}}
  .progress-sublabel{{font-size:11px;color:var(--muted);margin-top:5px;}}

  /* ── Curriculum sections ── */
  .curric-section{{margin-bottom:16px;border:1px solid var(--border);border-radius:8px;overflow:hidden;}}
  .curric-header{{display:flex;align-items:center;gap:10px;padding:10px 14px;background:var(--surface2);cursor:pointer;user-select:none;}}
  .curric-header:hover{{background:var(--border);}}
  .curric-title{{flex:1;font-size:13px;font-weight:600;}}
  .curric-chevron{{font-size:10px;color:var(--muted);transition:transform .2s;}}
  .curric-header.open .curric-chevron{{transform:rotate(180deg);}}
  .curric-items{{display:none;padding:8px 0;}}
  .curric-items.open{{display:block;}}
  .item-row{{display:flex;align-items:center;gap:10px;padding:7px 14px;border-bottom:1px solid var(--border);}}
  .item-row:last-child{{border-bottom:none;}}
  .item-check{{width:16px;height:16px;border-radius:3px;display:flex;align-items:center;justify-content:center;font-size:10px;flex-shrink:0;}}
  .item-check.done{{background:var(--green-subtle);color:var(--green);border:1px solid var(--green);}}
  .item-check.not{{background:var(--surface2);color:var(--muted);border:1px solid var(--border);}}
  .item-title{{flex:1;font-size:12px;}}
  .item-type{{font-size:9px;font-weight:700;padding:1px 6px;border-radius:3px;background:var(--surface2);color:var(--muted);text-transform:uppercase;}}
  .item-date{{font-size:10px;color:var(--muted);white-space:nowrap;}}

  /* ── Playbook section ── */
  .playbook-section{{margin-top:20px;}}
  .playbook-header{{font-size:12px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px;padding-bottom:8px;border-bottom:1px solid var(--border);}}
  .playbook-empty{{font-size:12px;color:var(--muted);padding:12px 0;}}
  .pb-visit-row{{display:flex;align-items:center;gap:10px;padding:6px 0;border-bottom:1px solid var(--border);}}
  .pb-visit-row:last-child{{border-bottom:none;}}
  .pb-page{{flex:1;font-size:12px;}}
  .pb-date{{font-size:11px;color:var(--muted);white-space:nowrap;}}

  /* ── Legend ── */
  .legend{{display:flex;gap:14px;flex-wrap:wrap;align-items:center;margin-bottom:10px;font-size:11px;color:var(--muted);}}
  .leg{{display:flex;align-items:center;gap:5px;}}
  .leg-dot{{width:10px;height:10px;border-radius:2px;}}

  /* ── Manager info ── */
  .mgr-block{{background:var(--surface2);border-radius:8px;padding:12px 14px;margin-bottom:16px;display:flex;flex-direction:column;gap:3px;}}
  .mgr-label{{font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);margin-bottom:2px;}}
  .mgr-name{{font-size:13px;font-weight:600;}}
  .mgr-detail{{font-size:11px;color:var(--muted);}}

  /* ── Charts ── */
  .charts{{display:grid;grid-template-columns:1fr 1fr;gap:16px;padding:0 28px 16px;}}
  @media(max-width:760px){{.charts{{grid-template-columns:1fr;}}}}
  .chart-card{{background:var(--surface);border:1px solid var(--border);border-radius:10px;padding:18px;}}
  .chart-title{{font-size:13px;font-weight:600;margin-bottom:14px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;}}
  body.light-mode .chart-title{{color:var(--text);}}
  .chart-wrap{{position:relative;height:220px;}}

  /* ── Print ── */
  @media print {{
    .header,.filters,.stats,.charts,.section,.print-hide,.modal-overlay{{display:none!important;}}
    #print-header,#print-roster-wrap{{display:block!important;}}
  }}
  #print-header,#print-roster-wrap{{display:none;}}
  #print-header{{padding:16px 0;margin-bottom:12px;border-bottom:2px solid #333;}}
  .ph-title{{font-size:18px;font-weight:700;}}
  .ph-sub{{font-size:12px;color:#555;margin-top:4px;}}
  #print-roster-wrap table{{width:100%;border-collapse:collapse;font-size:11px;}}
  #print-roster-wrap th{{background:#eee;padding:6px 8px;text-align:left;border:1px solid #ccc;font-size:10px;}}
  #print-roster-wrap td{{padding:5px 8px;border:1px solid #ddd;}}
</style>
</head>
<body>

<!-- Info popover -->
<div class="info-popover" id="info-popover"></div>

<!-- Header -->
<div class="header">
  <div class="header-left">
    <div class="hamburger" id="hamburger">
      <button class="hamburger-btn" id="hamburger-btn" onclick="toggleHamburger()">&#9776;</button>
      <div class="hamburger-menu" id="hamburger-menu">
        <div class="hamburger-section-label">Dashboards</div>
        <a class="hamburger-item" href="index.html">&#128202; Playbook Traffic Dashboard</a>
        <div class="hamburger-section-label">Certifications</div>
        <a class="hamburger-item" href="cert-healthcare.html">&#127973; Healthcare Certification</a>
        <a class="hamburger-item" href="cert-publicsector.html">&#127963; Public Sector Certification</a>
      </div>
    </div>
    <div>
      <h1>Accelerate Onboarding</h1>
    </div>
    <span class="data-badge" id="data-badge">{total} learners &middot; updated {file_date}</span>
  </div>
  <div class="header-right">
    <div class="export-drop print-hide" id="export-drop">
      <button class="btn-export" onclick="toggleExportDrop()">&#128438; Export &#9660;</button>
      <div class="export-menu" id="export-menu">
        <button class="export-item" onclick="runExport('full')">Full Report</button>
        <button class="export-item" onclick="runExport('overdue')">Overdue Only</button>
        <button class="export-item" onclick="runExport('manager-summary')">Manager Summary</button>
      </div>
    </div><span class="info-btn print-hide" onclick="showInfo(event,'export')">?</span>
    <button class="btn-theme" id="btn-theme" onclick="toggleTheme()">&#9728; Light</button>
  </div>
</div>

<!-- Filters -->
<div class="filters">
  <span class="filter-label">Market</span>
  <select id="f-market" onchange="applyFilters()"><option value="">All Markets</option></select>
  <span class="filter-label">Status</span>
  <select id="f-status" onchange="applyFilters()">
    <option value="">All</option>
    <option value="Completed">Completed</option>
    <option value="On Track">On Track</option>
    <option value="Overdue">Overdue</option>
  </select>
  <span class="filter-label">Sort</span>
  <select id="f-sort" onchange="applyFilters()">
    <option value="name">Name A→Z</option>
    <option value="pct-desc">Completion High→Low</option>
    <option value="pct-asc">Completion Low→High</option>
    <option value="days-asc">Most Urgent First</option>
  </select>
  <input type="text" id="f-search" placeholder="Search name or manager..." oninput="applyFilters()" style="width:200px;">
  <button class="btn-reset" onclick="resetFilters()">Reset</button>
  <button class="btn-tlg" id="btn-tlg" onclick="toggleTLG()">Hide TLG</button>
  <span class="result-count" id="result-count"></span>
</div>

<!-- Stat cards -->
<div class="stats">
  <div class="stat">
    <div class="stat-label">Total Enrolled <span class="info-btn" onclick="showInfo(event,'total-enrolled')">?</span></div>
    <div class="stat-value blue" id="s-total">&#8212;</div>
  </div>
  <div class="stat">
    <div class="stat-label">Overdue <span class="info-btn" onclick="showInfo(event,'overdue')">?</span></div>
    <div class="stat-value red" id="s-overdue">&#8212;</div>
    <div class="stat-sub" id="s-overdue-sub"></div>
  </div>
  <div class="stat">
    <div class="stat-label">On Track <span class="info-btn" onclick="showInfo(event,'ontrack')">?</span></div>
    <div class="stat-value" id="s-ontrack">&#8212;</div>
  </div>
  <div class="stat">
    <div class="stat-label">Completed <span class="info-btn" onclick="showInfo(event,'completed')">?</span></div>
    <div class="stat-value green" id="s-completed">&#8212;</div>
  </div>
  <div class="stat">
    <div class="stat-label">First Sale (Salesforce) <span class="info-btn" onclick="showInfo(event,'first-sale')">?</span></div>
    <div class="stat-value amber" id="s-sales">TBD</div>
    <div class="stat-sub">Integration pending</div>
  </div>
</div>

<!-- Charts -->
<div class="charts">
  <div class="chart-card">
    <div class="chart-title">Completion by Market <span class="info-btn" onclick="showInfo(event,'market-chart')">?</span></div>
    <div class="chart-wrap"><canvas id="marketChart"></canvas></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">Curriculum Progress <span class="info-btn" onclick="showInfo(event,'curric-chart')">?</span></div>
    <div class="chart-wrap"><canvas id="curricChart"></canvas></div>
  </div>
</div>

<!-- Heatmap section -->
<div class="section">
  <div class="section-header">
    <div>
      <div class="section-title">Learner Progress <span class="info-btn" onclick="showInfo(event,'heatmap')">?</span></div>
      <div class="section-hint">Click any row to see full detail &mdash; curriculum breakdown, course checklist &amp; playbook activity</div>
    </div>
    <div style="display:flex;align-items:center;gap:16px;flex-wrap:wrap;">
      <div class="legend">
        <span class="leg"><span class="leg-dot" style="background:#1baf7a"></span>75&ndash;100%</span>
        <span class="leg"><span class="leg-dot" style="background:#eda100"></span>40&ndash;74%</span>
        <span class="leg"><span class="leg-dot" style="background:#e34948"></span>1&ndash;39%</span>
        <span class="leg"><span class="leg-dot" style="background:#888780"></span>0%</span>
      </div>
      <input type="text" id="table-search" oninput="filterTableRows()" placeholder="Search name..." style="font-size:12px;padding:4px 10px;width:180px;background:var(--surface2);border:1px solid var(--border);color:var(--text);border-radius:6px;outline:none;">
    </div>
  </div>
  <div class="table-wrap">
    <table>
      <thead id="heatmap-head"></thead>
      <tbody id="heatmap-body"></tbody>
      <tfoot id="heatmap-foot"></tfoot>
    </table>
  </div>
</div>

<!-- User detail modal -->
<div class="modal-overlay" id="modal-overlay" onclick="closeModalBg(event)">
  <div class="modal-card" id="modal-card">
    <button class="modal-close" onclick="closeModal()">&#10005;</button>
    <div id="modal-content"></div>
  </div>
</div>

<!-- Print-only elements -->
<div id="print-header">
  <div class="ph-title">Accelerate Onboarding — <span id="ph-report-type">Full Report</span></div>
  <div class="ph-sub" id="ph-sub"></div>
</div>
<div id="print-roster-wrap"></div>

<script>
const PEOPLE = {people_json};
const CURRIC_IDS = {curric_ids};
const CURRIC_NAMES = {curric_names};
const PROGRAM_DAYS = {PROGRAM_DAYS};
const TODAY = new Date();
TODAY.setHours(0,0,0,0);

const TLG_SET = new Set([
  'jason ackerman','bianca davis','james parker','resmie biba',
  'chris curtis','sara thompson','jeremy macbean','bradley pierce',
  'laura sefcik','samantha maresca','staci musco','cj homer','rich moore','dale kinsey'
]);

let hideTLG = false;
let filtered = [];
let marketChartObj = null;
let curricChartObj = null;
let tableSort = {{col: 'overall', dir: 'asc'}};

/* ── Utilities ── */
function pct2color(p) {{
  if (p === 0) return {{bg:'#f1efe8', fg:'#888780'}};
  if (p < 40)  return {{bg:'#FCEBEB', fg:'#791F1F'}};
  if (p < 75)  return {{bg:'#FAEEDA', fg:'#633806'}};
  return {{bg:'#EAF3DE', fg:'#27500A'}};
}}

function computeStatus(p) {{
  if (p.overallDone) return 'Completed';
  if (p.daysRem === null || p.daysRem === undefined) return 'Unknown';
  return p.daysRem < 0 ? 'Overdue' : 'On Track';
}}

function daysLeft(p) {{
  if (p.daysRem === null || p.daysRem === undefined) return null;
  return p.daysRem;
}}

function daysElapsed(p) {{
  if (!p.assignDate) return 0;
  const start = new Date(p.assignDate);
  return Math.max(0, Math.round((TODAY - start) / 86400000));
}}

function fmtDate(d) {{
  if (!d) return '';
  const dt = new Date(d + 'T00:00:00');
  return dt.toLocaleDateString('en-US', {{month:'short', day:'numeric', year:'numeric'}});
}}

function cv() {{
  return getComputedStyle(document.body).getPropertyValue('--' + arguments[0]).trim();
}}

/* ── Theme ── */
function toggleTheme() {{
  const light = document.body.classList.toggle('light-mode');
  document.getElementById('btn-theme').innerHTML = light ? '&#9790; Dark' : '&#9728; Light';
  localStorage.setItem('pb-theme', light ? 'light' : 'dark');
  renderCharts();
}}
(function() {{
  const t = localStorage.getItem('pb-theme');
  if (t === 'light') {{
    document.body.classList.add('light-mode');
    document.getElementById('btn-theme').innerHTML = '&#9790; Dark';
  }}
}})();

/* ── Hamburger ── */
function toggleHamburger() {{
  document.getElementById('hamburger-btn').classList.toggle('open');
  document.getElementById('hamburger-menu').classList.toggle('open');
}}

/* ── Info popover ── */
const INFO = {{
  'total-enrolled': 'Total number of learners currently enrolled in the Accelerate Onboarding program, matching your active filters.',
  'overdue': 'Learners whose 45-day program window has expired and who have not yet completed all curricula. Based on their assignment date.',
  'ontrack': 'Learners still within their 45-day window who have not yet fully completed the program.',
  'completed': 'Learners who have completed all required curricula in the Accelerate Onboarding program.',
  'avg-completion': 'Average overall completion percentage across all learners shown, calculated from per-curriculum item completion.',
  'first-sale': 'Tracks whether each learner has logged their first sale in Salesforce. Salesforce integration is pending — this will populate automatically once connected.',
  'market-chart': 'Average overall completion % for each market, based on learners currently shown. Helpful for spotting which regions need support.',
  'curric-chart': 'Average completion % per sub-curriculum across all shown learners. Low bars indicate where learners are getting stuck.',
  'heatmap': 'One row per learner. Each column shows their completion % for that curriculum. Color: green = 75%+, amber = 40-74%, red = 1-39%, gray = not started. Click a row for full detail.',
  'export': 'Export is scoped to your active filters. Full Report: all learners with status and overall %. Overdue Only: contact list for follow-up with manager emails. Manager Summary: one row per manager with team count and avg completion. Example: filter to West Sls market, then export Overdue Only for a regional outreach list.',
}};
function showInfo(e, key) {{
  e.stopPropagation();
  const pop = document.getElementById('info-popover');
  if (pop.dataset.key === key && pop.classList.contains('visible')) {{
    pop.classList.remove('visible'); pop.dataset.key = ''; return;
  }}
  pop.dataset.key = key;
  pop.textContent = INFO[key] || '';
  pop.classList.add('visible');
  const r = e.target.getBoundingClientRect();
  pop.style.left = Math.min(r.left, window.innerWidth - 280) + 'px';
  pop.style.top  = (r.bottom + 6) + 'px';
}}

/* ── Export dropdown ── */
function toggleExportDrop() {{
  document.getElementById('export-menu').classList.toggle('open');
}}

/* ── Table search ── */
function filterTableRows() {{
  const q = (document.getElementById('table-search').value || '').toLowerCase();
  document.querySelectorAll('#heatmap-body tr').forEach(row => {{
    row.style.display = (!q || row.dataset.name.includes(q)) ? '' : 'none';
  }});
}}

/* ── Column sort ── */
function sortByCol(col) {{
  if (tableSort.col === col) {{
    tableSort.dir = tableSort.dir === 'asc' ? 'desc' : 'asc';
  }} else {{
    tableSort.col = col;
    tableSort.dir = 'asc';
  }}
  renderTable();
}}

/* ── TLG ── */
function toggleTLG() {{
  hideTLG = !hideTLG;
  document.getElementById('btn-tlg').classList.toggle('active', hideTLG);
  applyFilters();
}}

/* ── Filters ── */
function resetFilters() {{
  document.getElementById('f-market').value = '';
  document.getElementById('f-status').value = '';
  document.getElementById('f-sort').value = 'name';
  document.getElementById('f-search').value = '';
  document.getElementById('table-search').value = '';
  filterTableRows();
  applyFilters();
}}

function applyFilters() {{
  const mkt    = document.getElementById('f-market').value;
  const status = document.getElementById('f-status').value;
  const sort   = document.getElementById('f-sort').value;
  const search = document.getElementById('f-search').value.toLowerCase();

  filtered = PEOPLE.filter(p => {{
    if (hideTLG && TLG_SET.has(p.name.toLowerCase())) return false;
    if (mkt && p.market !== mkt) return false;
    const st = computeStatus(p);
    if (status && st !== status) return false;
    if (search && !p.name.toLowerCase().includes(search) && !p.manager.toLowerCase().includes(search)) return false;
    return true;
  }});

  if (sort === 'pct-desc') filtered.sort((a,b) => b.overallPct - a.overallPct);
  else if (sort === 'pct-asc') filtered.sort((a,b) => a.overallPct - b.overallPct);
  else if (sort === 'days-asc') filtered.sort((a,b) => {{
    const da = daysLeft(a) ?? 999, db = daysLeft(b) ?? 999;
    return da - db;
  }});
  else filtered.sort((a,b) => a.name.localeCompare(b.name));

  renderStats();
  renderTable();
  renderCharts();
  document.getElementById('result-count').textContent = filtered.length + ' shown';
}}

/* ── Stats ── */
function renderStats() {{
  const total = filtered.length;
  const overdue = filtered.filter(p => computeStatus(p) === 'Overdue').length;
  const ontrack = filtered.filter(p => computeStatus(p) === 'On Track').length;
  const completed = filtered.filter(p => computeStatus(p) === 'Completed').length;

  document.getElementById('s-total').textContent = total;
  document.getElementById('s-overdue').textContent = overdue;
  document.getElementById('s-overdue-sub').textContent = overdue ? overdue + ' past their 45-day window' : '';
  document.getElementById('s-ontrack').textContent = ontrack;
  document.getElementById('s-completed').textContent = completed;
}}

/* ── Heatmap table ── */
function renderTable() {{
  // Header
  const thead = document.getElementById('heatmap-head');
  function thS(label, col, extraCls) {{
    const active = tableSort.col === col;
    const arrow = active ? (tableSort.dir === 'asc' ? ' ↑' : ' ↓') : '';
    const cls = ['sortable-th', extraCls, active ? 'sort-active' : ''].filter(Boolean).join(' ');
    return '<th class="' + cls + '" data-col="' + col + '" onclick="sortByCol(this.dataset.col)">' + label + arrow + '</th>';
  }}
  let hRow = '<tr>' + thS('Learner','name') + thS('Market','market') + thS('Status','status') + thS('Days Left','days');
  hRow += thS('Overall','overall','overall-col');
  CURRIC_IDS.forEach(cid => {{
    hRow += thS(CURRIC_NAMES[cid], cid, 'curric-col');
  }});
  hRow += '</tr>';
  thead.innerHTML = hRow;

  // Body
  const tbody = document.getElementById('heatmap-body');
  if (!filtered.length) {{
    tbody.innerHTML = '<tr><td colspan="' + (5 + CURRIC_IDS.length) + '" style="text-align:center;color:var(--muted);padding:32px;">No learners match the current filters.</td></tr>';
    document.getElementById('heatmap-foot').innerHTML = '';
    return;
  }}

  let display = filtered.slice();
  if (tableSort.col) {{
    display.sort((a, b) => {{
      let va, vb;
      const col = tableSort.col;
      if (col === 'name')         {{ va = a.name; vb = b.name; }}
      else if (col === 'market')  {{ va = a.market; vb = b.market; }}
      else if (col === 'status')  {{ va = computeStatus(a); vb = computeStatus(b); }}
      else if (col === 'days')    {{ va = daysLeft(a) ?? 9999; vb = daysLeft(b) ?? 9999; }}
      else if (col === 'overall') {{ va = a.overallPct; vb = b.overallPct; }}
      else {{ va = a.curricula[col] ? a.curricula[col].pct : 0; vb = b.curricula[col] ? b.curricula[col].pct : 0; }}
      if (va < vb) return tableSort.dir === 'asc' ? -1 : 1;
      if (va > vb) return tableSort.dir === 'asc' ? 1 : -1;
      return 0;
    }});
  }}
  tbody.innerHTML = display.map(p => {{
    const status = computeStatus(p);
    const dl = daysLeft(p);
    const statusClass = status === 'Completed' ? 'sb-completed' : status === 'On Track' ? 'sb-ontrack' : 'sb-overdue';
    const daysStr = p.overallDone ? '&mdash;' :
      dl === null ? '&mdash;' :
      dl < 0 ? '<span style="color:var(--red);font-weight:700">' + Math.abs(dl) + ' past due</span>' :
      '<span style="color:var(--green)">' + dl + ' days</span>';

    let cells = '';
    CURRIC_IDS.forEach(cid => {{
      const c = p.curricula[cid];
      const pct = c ? c.pct : 0;
      const clr = pct2color(pct);
      cells += '<td class="pct-cell"><span class="pct-pill" style="background:' + clr.bg + ';color:' + clr.fg + '">' + pct + '%</span></td>';
    }});

    const oclr = pct2color(p.overallPct);

    return '<tr data-email="' + escHtml(p.email) + '" data-name="' + escHtml(p.name.toLowerCase()) + '" onclick="openModal(this.dataset.email)" title="Click to see full detail">' +
      '<td class="name-cell">' + escHtml(p.name) + '</td>' +
      '<td class="market-cell">' + escHtml(p.market) + '</td>' +
      '<td><span class="status-badge ' + statusClass + '">' + status + '</span></td>' +
      '<td style="font-size:11px;">' + daysStr + '</td>' +
      '<td class="pct-cell"><span class="pct-pill" style="background:' + oclr.bg + ';color:' + oclr.fg + ';font-weight:700">' + p.overallPct + '%</span></td>' +
      cells +
    '</tr>';
  }}).join('');

  // Footer (averages)
  const tfoot = document.getElementById('heatmap-foot');
  let fRow = '<tr><td colspan="4" style="font-weight:700;font-size:11px;">Averages (' + filtered.length + ' learners)</td>';
  const oAvg = filtered.length ? Math.round(filtered.reduce((s,p) => s+p.overallPct,0)/filtered.length) : 0;
  const oclr = pct2color(oAvg);
  fRow += '<td class="pct-cell"><span class="pct-pill" style="background:' + oclr.bg + ';color:' + oclr.fg + ';font-weight:700">' + oAvg + '%</span></td>';
  CURRIC_IDS.forEach(cid => {{
    const vals = filtered.map(p => p.curricula[cid] ? p.curricula[cid].pct : 0);
    const avg = vals.length ? Math.round(vals.reduce((s,v) => s+v,0) / vals.length) : 0;
    const clr = pct2color(avg);
    fRow += '<td class="pct-cell"><span class="pct-pill" style="background:' + clr.bg + ';color:' + clr.fg + '">' + avg + '%</span></td>';
  }});
  fRow += '</tr>';
  tfoot.innerHTML = fRow;
}}

function escHtml(s) {{
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}}

/* ── Charts ── */
function renderCharts() {{
  renderMarketChart();
  renderCurricChart();
}}

function renderMarketChart() {{
  const markets = [...new Set(filtered.map(p => p.market))].sort();
  const avgs = markets.map(m => {{
    const mp = filtered.filter(p => p.market === m);
    return Math.round(mp.reduce((s,p) => s+p.overallPct,0)/mp.length);
  }});
  const ctx = document.getElementById('marketChart').getContext('2d');
  if (marketChartObj) marketChartObj.destroy();
  marketChartObj = new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: markets,
      datasets: [{{
        label: 'Avg Completion %',
        data: avgs,
        backgroundColor: '#3b82f6aa',
        borderColor: '#3b82f6',
        borderWidth: 1,
        borderRadius: 4,
      }}]
    }},
    options: {{
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend: {{display:false}},
        tooltip: {{callbacks: {{label: ctx => ' ' + ctx.parsed.x + '%'}}}}
      }},
      scales: {{
        x: {{
          min:0, max:100,
          grid: {{color: 'rgba(255,255,255,0.06)'}},
          ticks: {{color: cv('muted'), font:{{size:11}}, callback: v => v+'%'}},
          border: {{color:'transparent'}},
        }},
        y: {{
          grid: {{display:false}},
          ticks: {{color: cv('muted'), font:{{size:11}}}},
          border: {{color:'transparent'}},
        }}
      }}
    }}
  }});
}}

function renderCurricChart() {{
  const labels = CURRIC_IDS.map(cid => CURRIC_NAMES[cid]);
  const avgs = CURRIC_IDS.map(cid => {{
    const vals = filtered.map(p => p.curricula[cid] ? p.curricula[cid].pct : 0);
    return vals.length ? Math.round(vals.reduce((s,v) => s+v,0)/vals.length) : 0;
  }});
  const ctx = document.getElementById('curricChart').getContext('2d');
  if (curricChartObj) curricChartObj.destroy();
  const colors = ['#3b82f6','#22c55e','#f59e0b','#a855f7','#ef4444','#14b8a6'];
  curricChartObj = new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels,
      datasets: [{{
        label: 'Avg %',
        data: avgs,
        backgroundColor: colors.map(c => c+'99'),
        borderColor: colors,
        borderWidth: 1,
        borderRadius: 4,
      }}]
    }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend: {{display:false}},
        tooltip: {{callbacks: {{label: ctx => ' ' + ctx.parsed.y + '%'}}}}
      }},
      scales: {{
        y: {{
          min:0, max:100,
          grid: {{color: 'rgba(255,255,255,0.06)'}},
          ticks: {{color: cv('muted'), font:{{size:11}}, callback: v => v+'%'}},
          border: {{color:'transparent'}},
        }},
        x: {{
          grid: {{display:false}},
          ticks: {{color: cv('muted'), font:{{size:10}}, maxRotation:30}},
          border: {{color:'transparent'}},
        }}
      }}
    }}
  }});
}}

/* ── User detail modal ── */
function openModal(email) {{
  const p = PEOPLE.find(x => x.email === email);
  if (!p) return;
  const status = computeStatus(p);
  const dl = daysLeft(p);
  const elapsed = daysElapsed(p);
  const pct45 = Math.min(100, Math.round(elapsed / PROGRAM_DAYS * 100));
  const barColor = p.overallDone ? 'var(--green)' : dl < 0 ? 'var(--red)' : 'var(--accent)';

  let daysLabel = '';
  if (p.overallDone) daysLabel = 'Program complete';
  else if (dl === null) daysLabel = 'No assignment date';
  else if (dl < 0) daysLabel = Math.abs(dl) + ' days past the 45-day window';
  else daysLabel = 'Day ' + elapsed + ' of ' + PROGRAM_DAYS + ' &mdash; ' + dl + ' days remaining';

  const statusBadge = '<span class="status-badge ' +
    (status==='Completed'?'sb-completed':status==='On Track'?'sb-ontrack':'sb-overdue') +
    '">' + status + '</span>';

  // Curriculum sections
  let curricHtml = '';
  CURRIC_IDS.forEach(cid => {{
    const c = p.curricula[cid];
    if (!c) return;
    const clr = pct2color(c.pct);
    const drBg = c.daysRem < 0 ? 'var(--red-subtle)' : 'var(--green-subtle)';
    const drColor = c.daysRem < 0 ? 'var(--red)' : 'var(--green)';
    const drLabel = c.daysRem === null ? '' :
      '<span class="days-badge" style="background:' + drBg + ';color:' + drColor + '">' +
      (c.daysRem < 0 ? Math.abs(c.daysRem) + 'd overdue' : c.daysRem + 'd left') + '</span>';
    const doneBadge = c.complete ?
      '<span class="status-badge sb-completed" style="font-size:10px">Done</span>' : '';

    const itemsHtml = c.items.map(it => {{
      const checkClass = it.done ? 'done' : 'not';
      const checkMark  = it.done ? '&#10003;' : '&nbsp;';
      const typeLabel  = it.type === 'VILT' ?
        '<span class="item-type" style="background:#a855f711;color:#a855f7;">Workshop</span>' :
        '<span class="item-type">Online</span>';
      const dateStr = it.done && it.date ?
        '<span class="item-date">' + fmtDate(it.date) + '</span>' : '';
      return '<div class="item-row">' +
        '<div class="item-check ' + checkClass + '">' + checkMark + '</div>' +
        '<div class="item-title">' + escHtml(it.title) + '</div>' +
        typeLabel + dateStr + '</div>';
    }}).join('');

    curricHtml += '<div class="curric-section">' +
      '<div class="curric-header" onclick="toggleCurric(this)">' +
        '<span class="curric-title">' + escHtml(c.title) + '</span>' +
        '<span class="pct-pill" style="background:' + clr.bg + ';color:' + clr.fg + ';font-size:11px">' + c.pct + '%</span>' +
        drLabel + doneBadge +
        '<span class="curric-chevron">&#9660;</span>' +
      '</div>' +
      '<div class="curric-items">' + itemsHtml + '</div>' +
    '</div>';
  }});

  // Playbook activity
  let pbHtml = '';
  if (!p.playbook || p.playbook.length === 0) {{
    pbHtml = '<div class="playbook-empty">No Accelerate playbook visits recorded for this learner.</div>';
  }} else {{
    pbHtml = p.playbook.map(v =>
      '<div class="pb-visit-row">' +
        '<span class="pb-page">&#128196; ' + escHtml(v.page) + '</span>' +
        '<span class="pb-date">' + fmtDate(v.date) + '</span>' +
      '</div>'
    ).join('');
  }}

  const content = `
    <div class="modal-header">
      <div style="display:flex;align-items:flex-start;gap:12px;flex-wrap:wrap;">
        <div style="flex:1">
          <div class="modal-name">${{escHtml(p.name)}}</div>
          <div class="modal-meta">
            <span>${{escHtml(p.jobTitle)}}</span>
            <span>&middot; ${{escHtml(p.market)}}</span>
            ${{p.hireDate ? '<span>&middot; Hired ' + fmtDate(p.hireDate) + '</span>' : ''}}
            ${{p.assignDate ? '<span>&middot; Enrolled ' + fmtDate(p.assignDate) + '</span>' : ''}}
          </div>
        </div>
        ${{statusBadge}}
      </div>
    </div>
    <div class="modal-body">
      <div class="progress-section">
        <div class="progress-label">
          <span>45-Day Program Window</span>
          <span style="font-size:11px;color:var(--muted)">${{daysLabel}}</span>
        </div>
        <div class="progress-bar-wrap">
          <div class="progress-bar-fill" style="width:${{pct45}}%;background:${{barColor}}"></div>
        </div>
        <div class="progress-sublabel">Overall curriculum completion: <strong>${{p.overallPct}}%</strong></div>
      </div>

      ${{p.manager ? '<div class="mgr-block"><div class="mgr-label">Manager</div><div class="mgr-name">' + escHtml(p.manager) + '</div><div class="mgr-detail">' + escHtml(p.mgrTitle) + ' &middot; <a href="mailto:' + escHtml(p.mgrEmail) + '" style="color:var(--accent)">' + escHtml(p.mgrEmail) + '</a></div></div>' : ''}}

      <div style="font-size:12px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:10px;">Curriculum Progress</div>
      ${{curricHtml}}

      <div class="playbook-section">
        <div class="playbook-header">Accelerate Playbook Activity (${{p.playbook ? p.playbook.length : 0}} visits)</div>
        ${{pbHtml}}
      </div>
    </div>
  `;

  document.getElementById('modal-content').innerHTML = content;
  document.getElementById('modal-overlay').classList.add('open');
}}

function toggleCurric(header) {{
  header.classList.toggle('open');
  header.nextElementSibling.classList.toggle('open');
}}

function closeModal() {{
  document.getElementById('modal-overlay').classList.remove('open');
}}

function closeModalBg(e) {{
  if (e.target === document.getElementById('modal-overlay')) closeModal();
}}

/* ── Export ── */
function runExport(type) {{
  document.getElementById('export-menu').classList.remove('open');
  const now = new Date().toLocaleDateString('en-US', {{month:'long', day:'numeric', year:'numeric'}});
  document.getElementById('ph-sub').textContent = filtered.length + ' learners · ' + now +
    (type !== 'full' ? ' · Filters active' : '');
  document.getElementById('ph-report-type').textContent =
    type === 'full' ? 'Full Report' :
    type === 'overdue' ? 'Overdue Learners' : 'Manager Summary';

  const wrap = document.getElementById('print-roster-wrap');
  if (type === 'full') {{
    const rows = filtered.map(p => {{
      const status = computeStatus(p);
      const dl = daysLeft(p);
      const daysStr = p.overallDone ? 'Complete' : dl === null ? '—' : dl < 0 ? (Math.abs(dl) + 'd overdue') : (dl + 'd left');
      return '<tr><td>' + escHtml(p.name) + '</td><td>' + escHtml(p.market) + '</td><td>' + status + '</td><td>' + daysStr + '</td><td>' + p.overallPct + '%</td><td>' + escHtml(p.manager) + '</td></tr>';
    }}).join('');
    wrap.innerHTML = '<table><thead><tr><th>Name</th><th>Market</th><th>Status</th><th>Days Left</th><th>Overall %</th><th>Manager</th></tr></thead><tbody>' + rows + '</tbody></table>';
  }} else if (type === 'overdue') {{
    const od = filtered.filter(p => computeStatus(p) === 'Overdue');
    const rows = od.map(p => '<tr><td>' + escHtml(p.name) + '</td><td>' + escHtml(p.email) + '</td><td>' + escHtml(p.market) + '</td><td>' + p.overallPct + '%</td><td>' + escHtml(p.manager) + '</td><td>' + escHtml(p.mgrEmail) + '</td></tr>').join('');
    wrap.innerHTML = '<table><thead><tr><th>Name</th><th>Email</th><th>Market</th><th>Completion %</th><th>Manager</th><th>Manager Email</th></tr></thead><tbody>' + rows + '</tbody></table>';
  }} else {{
    const byMgr = {{}};
    filtered.forEach(p => {{
      const m = p.manager || 'Unknown';
      if (!byMgr[m]) byMgr[m] = {{mgrEmail:p.mgrEmail, people:[]}};
      byMgr[m].people.push(p);
    }});
    const rows = Object.entries(byMgr).sort((a,b)=>a[0].localeCompare(b[0])).map(([mgr,data]) => {{
      const avg = Math.round(data.people.reduce((s,p)=>s+p.overallPct,0)/data.people.length);
      const od = data.people.filter(p=>computeStatus(p)==='Overdue').length;
      return '<tr><td>' + escHtml(mgr) + '</td><td>' + escHtml(data.mgrEmail) + '</td><td>' + data.people.length + '</td><td>' + od + '</td><td>' + avg + '%</td></tr>';
    }}).join('');
    wrap.innerHTML = '<table><thead><tr><th>Manager</th><th>Manager Email</th><th>Team Count</th><th>Overdue</th><th>Avg Completion %</th></tr></thead><tbody>' + rows + '</tbody></table>';
  }}
  window.print();
}}

/* ── Init ── */
function init() {{
  // Populate market dropdown
  const markets = [...new Set(PEOPLE.map(p => p.market))].sort();
  const sel = document.getElementById('f-market');
  markets.forEach(m => {{
    const o = document.createElement('option');
    o.value = m; o.textContent = m;
    sel.appendChild(o);
  }});

  // Close dropdowns on outside click
  document.addEventListener('click', e => {{
    if (!document.getElementById('export-drop').contains(e.target))
      document.getElementById('export-menu').classList.remove('open');
    if (!document.getElementById('hamburger').contains(e.target)) {{
      document.getElementById('hamburger-btn').classList.remove('open');
      document.getElementById('hamburger-menu').classList.remove('open');
    }}
    if (!document.getElementById('info-popover').contains(e.target) && !e.target.classList.contains('info-btn'))
      document.getElementById('info-popover').classList.remove('visible');
  }});

  applyFilters();
}}

init();
</script>
</body>
</html>"""


def main():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    records = load_lms()
    print(f"Loaded {len(records)} learners from LMS data")
    records = attach_playbook(records)
    pb_matches = sum(1 for p in records.values() if p['playbook'])
    print(f"Playbook activity matched for {pb_matches} learners")
    html = generate_html(records)
    with open('onboarding.html', 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"Generated onboarding.html")


if __name__ == '__main__':
    main()
