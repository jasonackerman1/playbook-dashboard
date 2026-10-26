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
COL_ITEM_REQ   = 30  # Item Required Date
COL_PARENT_CURRIC = 16  # Parent Curriculum ID (ACCELERATE or ACCELERATE_BCA)

# ── Registration status report (VILT/workshop sign-ups) ─────────────────────
REGISTRATION_FILE_GLOB = 'onboarding-data/RegistrationStatusWithOrgCSV*.xlsx'

# ── Certification Journey tab data (Boot Camp / Social Selling / Capstone / Certified) ──
CERT_JOURNEY_FILE_GLOB       = 'onboarding-data/Accelerate-Cert-Journey-*.xlsx'
CERT_CERTIFICATION_FILE_GLOB = 'onboarding-data/Accelerate-Certification-Report-*.xlsx'
_CERT_STAGE_TITLE_MAP = [
    ('bootcamp', 'bootCamp'),
    ('social selling', 'socialSelling'),
    ('capstone', 'capstone'),
]

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

PROGRAM_DAYS = 35

TLG = {
    'jason ackerman', 'bianca davis', 'james parker', 'resmie biba',
    'chris curtis', 'sara thompson', 'jeremy macbean', 'bradley pierce',
    'laura sefcik', 'samantha maresca', 'staci musco', 'cj homer',
    'rich moore', 'dale kinsey',
    'john lechner', 'resmie nesimi', "samantha d'angelo", 'bianca dipasquale', 'doug falk',
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
    'salesforceprospecting': 'Salesforce Prospecting',
    'callprep':              'Prep Essentials',
    'workingwithnumbers':    'Working With Numbers',
    'movingdeals':           'Moving Deals Forward',
    'pipelineownership':     'Pipeline Ownership',
}


def pkey(first, last):
    return f"{str(first).strip().lower()} {str(last).strip().lower()}"


def _date(val):
    """Return YYYY-MM-DD string from a datetime, YYYY-MM-DD string, LMS M/D/YYYY format, or Excel serial int."""
    if not val:
        return None
    if hasattr(val, 'strftime'):
        return val.strftime('%Y-%m-%d')
    if isinstance(val, (int, float)):
        # Excel date serial number (days since 1899-12-30)
        from datetime import date as _d, timedelta
        try:
            return (_d(1899, 12, 30) + timedelta(days=int(val))).strftime('%Y-%m-%d')
        except Exception:
            return None
    s = str(val).strip()
    if not s:
        return None
    # LMS exports dates as "M/D/YYYY timezone" e.g. "6/11/2026 US/Alaska"
    m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{4})', s)
    if m:
        return f'{int(m.group(3)):04d}-{int(m.group(1)):02d}-{int(m.group(2)):02d}'
    return s


def extract_date(fname):
    base = os.path.basename(fname)
    m = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', base)  # MM.DD.YYYY
    if m: return f'{m.group(3)}-{m.group(1)}-{m.group(2)}'
    m = re.search(r'(\d{4}-\d{2})', base)               # YYYY-MM fallback
    return m.group(1) if m else '0000-00'


def _file_date_label(fname):
    """Return a human-readable 'Data through' date from a filename."""
    import datetime as _dt
    _MONTHS = ['January','February','March','April','May','June',
               'July','August','September','October','November','December']
    base = os.path.basename(str(fname))
    m = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', base)
    if m:
        return f'{_MONTHS[int(m.group(1))-1]} {int(m.group(2))}, {m.group(3)}'
    m = re.search(r'(\d{2})\.(\d{2})(?!\.\d)', base)
    if m:
        yr = _dt.datetime.today().year
        return f'{_MONTHS[int(m.group(1))-1]} {int(m.group(2))}, {yr}'
    m = re.search(r'(\d{4})-(\d{2})', base)
    if m:
        return f'{_MONTHS[int(m.group(2))-1]} {m.group(1)}'
    dt = _dt.datetime.fromtimestamp(os.path.getmtime(str(fname)))
    return f'{_MONTHS[dt.month-1]} {dt.day}, {dt.year}'


def page_label(url):
    m = re.search(r'/([^/]+?)(?:\.html)?$', url.rstrip('/'))
    if not m or m.group(1) in ('accelerate_sales_playbook', ''):
        return 'Home'
    slug = m.group(1).lower()
    return PAGE_LABELS.get(slug, slug.replace('-', ' ').title())


def _find_col(header, *substrings, exclude=None):
    """Find a column index by case-insensitive substring match on the header row."""
    for i, h in enumerate(header):
        hl = str(h or '').strip().lower()
        if exclude and exclude in hl:
            continue
        if any(s in hl for s in substrings):
            return i
    return None


def attach_registration_status(records):
    """Flags incomplete VILT (workshop) curriculum items as Registered when the learner has
    an Active Enrollment row for that exact course title in the registration status report.
    Non-fatal if the report isn't present yet — items just stay unflagged (renders as
    'Not Registered' in the UI) until Resmie drops one in."""
    import glob
    files = sorted(glob.glob(REGISTRATION_FILE_GLOB))
    if not files:
        print("No registration status report found — Registered/Not Registered tags will show as Not Registered until one is added")
        return records
    fpath = files[-1]
    wb = openpyxl.load_workbook(fpath, data_only=True)
    ws = wb.active
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    email_col  = _find_col(header, 'email', exclude='supervisor')
    title_col  = _find_col(header, 'title', exclude='user')
    status_col = _find_col(header, 'status')
    if email_col is None or title_col is None or status_col is None:
        print(f"WARNING: registration report {os.path.basename(fpath)} missing an expected column (Email/Title/Status) — skipping")
        return records

    active = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        status = str(row[status_col] or '').strip().lower()
        if status != 'active enrollment':
            continue
        email = str(row[email_col] or '').strip().lower()
        title = str(row[title_col] or '').strip()
        if email and title:
            active.add((email, title))

    matched_people = 0
    matched_items = 0
    for p in records.values():
        person_matched = False
        for curric in p['curricula'].values():
            for item in curric['items']:
                if item['type'] == 'VILT' and not item['done'] and (p['email'], item['title']) in active:
                    item['registered'] = True
                    matched_items += 1
                    person_matched = True
        if person_matched:
            matched_people += 1
    print(f"Registration status ({os.path.basename(fpath)}): {matched_people} people, {matched_items} curriculum items flagged Registered")
    return records


def _cert_stage_for_title(title):
    t = str(title or '').strip().lower()
    for needle, key in _CERT_STAGE_TITLE_MAP:
        if needle in t:
            return key
    return None


def _first_date(*vals):
    for v in vals:
        if v is None:
            continue
        d = _date(v)
        if d:
            return d
    return None


def _apply_cert_registration_rows(ws, journey, warn_unknown):
    """Walk one registration-report worksheet (shared schema: Class ID / Start Date/Time /
    End Date/Time / Student Email / Title / Status / ENRL_DTE) and fold any Boot Camp / Social
    Selling / Capstone rows into `journey` in place. Rows for unrelated workshops (Salesforce
    Live Workshop, DCM, Prospecting Live Workshop, etc.) are silently ignored when warn_unknown
    is False — the general registration report legitimately carries many workshop types that
    have nothing to do with the Certification Journey. Returns the count of rows matched."""
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    email_col   = _find_col(header, 'student email') or _find_col(header, 'email', exclude='supervisor')
    title_col   = _find_col(header, 'title', exclude='user')
    status_col  = _find_col(header, 'status')
    classid_col = _find_col(header, 'class id')
    start_col   = _find_col(header, 'start date')
    end_col     = _find_col(header, 'end date')
    enrl_col    = _find_col(header, 'enrl')
    if email_col is None or title_col is None or status_col is None:
        return None

    unknown_titles = set()
    matched_rows = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        email = str(row[email_col] or '').strip().lower()
        title = row[title_col]
        if not email or not title:
            continue
        stage = _cert_stage_for_title(title)
        if not stage:
            unknown_titles.add(str(title))
            continue
        status = str(row[status_col] or '').strip().lower()
        # Cancelled/Waitlisted rows exist in past registration files (though not in today's) —
        # never treat either as "registered"; only Active Enrollment or a completed session count.
        if 'complete' not in status and 'active' not in status:
            continue
        rec = journey.setdefault(email, {})
        if 'complete' in status:
            rec[stage] = {
                'attended': True,
                'date': _first_date(row[end_col] if end_col is not None else None,
                                     row[start_col] if start_col is not None else None),
            }
        else:
            rec[stage] = {
                'registered':   True,
                'sessionDate':  _date(row[start_col]) if start_col is not None else None,
                'sessionEnd':   _date(row[end_col]) if end_col is not None else None,
                'classId':      str(row[classid_col]) if classid_col is not None and row[classid_col] is not None else None,
                'enrolledDate': _date(row[enrl_col]) if enrl_col is not None else None,
            }
        matched_rows += 1
    if warn_unknown:
        for t in sorted(unknown_titles):
            print(f"WARNING: unknown Certification Journey class title '{t}' — not mapped to Boot Camp/Social Selling/Capstone")
    return matched_rows


def load_cert_journey():
    """Real Boot Camp / Social Selling / Capstone / Certified data for the Certification
    Journey tab — replaces the frozen 2026-10-21 placeholder (onboarding-data/_cert_journey_placeholder.json)
    now that Resmie's real reports have started arriving. Three independent, optional sources —
    missing any one just leaves that part of the journey blank rather than failing:

      - Accelerate-Cert-Journey-*.xlsx: one row per class enrollment, dedicated to this journey
        (Boot Camp + Social Selling as of 2026-10-26 — no Capstone rows in this one yet).
      - RegistrationStatusWithOrgCSV*.xlsx: the broader registration report also used elsewhere
        for general VILT registration flagging — it carries many unrelated workshop types, but
        also real Boot Camp AND Capstone rows (Social Selling never appears in it). Confirmed
        2026-10-26 the two files agree wherever both cover the same person+stage (zero conflicts),
        so merging them is safe and is what actually surfaces real Capstone registrations.
      - Accelerate-Certification-Report-*.xlsx: one row per person who has completed the
        "Accelerate Certification" course — same "any row present = certified" rule already used
        for Layered Security's Certification Report.

    A row's own Status tells us whether it's a completed session ('...complete...' -> attended)
    or just booked ('Active Enrollment' -> registered, with the session's dates). The title match
    is substring-based, so Capstone attendance (or any new session type) picks up automatically
    the moment the LMS starts reporting it, no code change needed.

    Returns (journey, meta) where journey is {email_lower: {bootCamp/socialSelling/capstone/
    certification: {...}}} — the exact shape certSteps() in the generated JS expects.
    """
    import glob

    journey = {}
    meta = {}
    reg_sources = []  # (label, fpath) for the asof note

    cj_files = sorted(glob.glob(CERT_JOURNEY_FILE_GLOB), key=extract_date)
    if cj_files:
        fpath = cj_files[-1]
        wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
        n = _apply_cert_registration_rows(wb.active, journey, warn_unknown=True)
        wb.close()
        if n is None:
            print(f"WARNING: {os.path.basename(fpath)} missing an expected column (Email/Title/Status) — skipped")
        else:
            print(f"Cert Journey registrations ({os.path.basename(fpath)}): {n} class rows matched")
            reg_sources.append(fpath)

    reg_files = sorted(glob.glob(REGISTRATION_FILE_GLOB), key=extract_date)
    if reg_files:
        fpath = reg_files[-1]
        wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
        n = _apply_cert_registration_rows(wb.active, journey, warn_unknown=False)
        wb.close()
        if n:
            print(f"Cert Journey registrations ({os.path.basename(fpath)}): {n} additional class rows matched (Capstone/Boot Camp overlap)")
            reg_sources.append(fpath)

    if reg_sources:
        print(f"Cert Journey registrations: {len(journey)} people total across {len(reg_sources)} report(s)")
        meta['registrationSource']  = 'LMS registration reports'
        meta['registrationThrough'] = max(extract_date(f) for f in reg_sources)

    cert_files = sorted(glob.glob(CERT_CERTIFICATION_FILE_GLOB), key=extract_date)
    if cert_files:
        fpath = cert_files[-1]
        wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
        ws = wb.active
        header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        email_col = _find_col(header, 'email')
        title_col = _find_col(header, 'course title')
        date_col  = _find_col(header, 'completion date')
        if email_col is None or date_col is None:
            print(f"WARNING: {os.path.basename(fpath)} missing an expected column (Email/Completion Date) — Certified status skipped")
        else:
            certified_count = 0
            for row in ws.iter_rows(min_row=2, values_only=True):
                email = str(row[email_col] or '').strip().lower()
                if not email:
                    continue
                title = row[title_col] if title_col is not None else None
                if title is not None and 'certification' not in str(title).lower():
                    continue
                d = _date(row[date_col])
                if not d:
                    continue
                journey.setdefault(email, {})['certification'] = {'completed': True, 'date': d}
                certified_count += 1
            print(f"Cert Journey certifications ({os.path.basename(fpath)}): {certified_count} people certified")
        wb.close()
        meta['completionSource']  = 'LMS Accelerate Certification report'
        meta['completionThrough'] = extract_date(fpath)

    return journey, meta


def load_lms():
    folder = 'onboarding-data'
    all_files = sorted(
        [f for f in os.listdir(folder) if f.startswith('Accelerate-Curriculum-Report-') and f.endswith('.xlsx')],
        key=extract_date
    )
    files = all_files[-1:]  # only use the most recent file — people not in it have left the program
    seen = {}  # email -> record

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

            # Detect Canadian learner — parent curriculum is ACCELERATE_BCA
            is_canada = any(str(r[COL_PARENT_CURRIC] or '').endswith('_BCA') for r in prows)
            suffix    = '_BCA' if is_canada else ''
            parent_id = f'ACCELERATE{suffix}'

            # Parent curriculum row (overall program)
            parent = next((r for r in prows if r[COL_CURRIC_ID] == parent_id), None)
            assign_date  = parent[COL_ASSIGN_DT]  if parent else None
            overall_done = (parent[COL_CURRIC_CMP] == 'Yes') if parent else False
            days_rem_lms = None  # computed in JS from assignDate

            # Per sub-curriculum data — always stored under US key so JS heatmap works unchanged
            curricula_data = {}
            for cid, cname in CURRICULA:
                crows = [r for r in prows if r[COL_CURRIC_ID] == cid + suffix]
                if not crows:
                    continue

                curric_done    = crows[0][COL_CURRIC_CMP] == 'Yes'
                item_rows      = [r for r in crows if r[COL_ITEM_TTL]]

                items = []
                for ir in item_rows:
                    title = str(ir[COL_ITEM_TTL] or '')
                    is_coming_soon = 'coming soon' in title.lower()
                    items.append({
                        'title': title,
                        'type':  str(ir[COL_ITEM_TYPE] or 'ONLINE'),
                        'done':  bool(ir[COL_ITEM_STS]),
                        'date':  _date(ir[COL_ITEM_DATE]),
                        # Coming Soon courses aren't launched yet, so they never have a real
                        # due date — force req=None so they can't trigger Overdue once the
                        # LMS starts populating Item Required Date for other courses.
                        'req':   None if is_coming_soon else _date(ir[COL_ITEM_REQ]),
                    })

                total = len(items)
                done  = sum(1 for i in items if i['done'])
                pct   = round(done / total * 100) if total else 0

                curricula_data[cid] = {
                    'title':    cname,
                    'complete': curric_done,
                    'pct':      pct,
                    'done':     done,
                    'total':    total,
                    'items':    items,
                }

            all_done    = sum(v['done']  for v in curricula_data.values())
            all_total   = sum(v['total'] for v in curricula_data.values())
            overall_pct = round(all_done / all_total * 100) if all_total else 0

            seen[email] = {
                'email':       email,
                'key':         pkey(first, last),
                'name':        f"{first} {last}",
                'first':       first,
                'last':        last,
                'jobTitle':    str(prows[0][COL_JOBTITLE]  or ''),
                'market':      str(prows[0][COL_MARKET]    or ''),
                'isCanada':    is_canada,
                'manager':     f"{prows[0][COL_MGR_FIRST] or ''} {prows[0][COL_MGR_LAST] or ''}".strip(),
                'mgrEmail':    str(prows[0][COL_MGR_EMAIL] or ''),
                'mgrTitle':    str(prows[0][COL_MGR_TITLE] or ''),
                'hireDate':    _date(prows[0][COL_HIRE_DATE]),
                'assignDate':  _date(assign_date),
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


def _parse_html_xls(path):
    from html.parser import HTMLParser
    class _TP(HTMLParser):
        def __init__(self):
            super().__init__()
            self.rows, self._row, self._cell, self._in = [], [], '', False
        def handle_starttag(self, tag, attrs):
            if tag in ('td', 'th'):
                self._in = True; self._cell = ''
            elif tag == 'tr':
                self._row = []
        def handle_endtag(self, tag):
            if tag in ('td', 'th'):
                self._row.append(self._cell.strip()); self._in = False
            elif tag == 'tr':
                if self._row: self.rows.append(self._row)
        def handle_data(self, data):
            if self._in: self._cell += data
    with open(path, encoding='utf-8', errors='ignore') as f:
        content = f.read()
    p = _TP()
    p.feed(content)
    if not p.rows:
        return []
    headers = p.rows[0]
    return [dict(zip(headers, row)) for row in p.rows[1:] if row]


def _load_salesforce(records):
    """Read Closed Won opps from onboarding-data/ xlsx (or leaderboard-data/ xls fallback), return first CW deal per cohort email."""
    import glob
    xlsx_files = sorted(glob.glob(os.path.join('onboarding-data', 'New-Opportunities-Report-*.xlsx')))
    if xlsx_files:
        import openpyxl
        wb = openpyxl.load_workbook(xlsx_files[-1], read_only=True, data_only=True)
        ws = wb.active
        rows_iter = iter(ws.rows)
        headers = [str(c.value).strip() if c.value is not None else '' for c in next(rows_iter)]
        raw_rows = [dict(zip(headers, [c.value for c in row])) for row in rows_iter]
        wb.close()
    else:
        xls_files = glob.glob(os.path.join('leaderboard-data', '*.xls'))
        cw_path = None
        for p in xls_files:
            with open(p, encoding='utf-8', errors='ignore') as f:
                chunk = f.read(2000)
            if 'Opportunity Owner Email' in chunk:
                cw_path = p
                break
        if not cw_path:
            return {}
        raw_rows = _parse_html_xls(cw_path)

    from datetime import datetime as _dt

    sales = {}
    for row in raw_rows:
        if str(row.get('Stage', '') or '').strip() != 'Closed Won':
            continue
        email = str(row.get('Opportunity Owner Email', '') or '').strip().lower()
        if email not in records:
            continue
        try:
            amount = float(str(row.get('Amount', '0') or '0').replace(',', '').strip())
        except (ValueError, TypeError):
            amount = 0.0
        raw_date = row.get('Close Date', '')
        if hasattr(raw_date, 'strftime'):
            close_date = raw_date.strftime('%Y-%m-%d')
        else:
            raw_date = str(raw_date or '')
            m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{4})', raw_date)
            close_date = f'{int(m.group(3)):04d}-{int(m.group(1)):02d}-{int(m.group(2)):02d}' if m else raw_date[:10]
        account = str(row.get('Account Name', '') or '').strip().title()
        # Keep earliest closed deal per rep (true "first sale")
        if email not in sales or close_date < sales[email]['closeDate']:
            # Days to Close = hire date -> earliest Closed Won deal's close date
            # (not Salesforce's own Age field, which measures opp creation -> close)
            hire_date = records[email].get('hireDate')
            days_to_close = None
            if hire_date and close_date:
                try:
                    delta = (_dt.strptime(close_date[:10], '%Y-%m-%d') -
                             _dt.strptime(hire_date, '%Y-%m-%d')).days
                    if delta >= 0:
                        days_to_close = delta
                except ValueError:
                    days_to_close = None
            sales[email] = {'amount': amount, 'accountName': account, 'closeDate': close_date, 'daysToClose': days_to_close}
    return sales


def generate_html(records, sales_map=None):
    from datetime import date
    people           = sorted(records.values(), key=lambda p: p['name'].lower())
    people_json      = json.dumps(people, default=str)
    sales_map        = sales_map or {}
    sales_map_json   = json.dumps(sales_map, ensure_ascii=False)
    cohort_total     = len(people)
    today_str        = date.today().isoformat()
    try:
        _lms_dir   = 'onboarding-data'
        _lms_files = [f for f in os.listdir(_lms_dir) if f.endswith('.xlsx')]
        _lms_files.sort()
        _lms_label = _file_date_label(os.path.join(_lms_dir, _lms_files[-1])) if _lms_files else date.today().strftime('%B %-d, %Y')
    except Exception:
        _lms_label = date.today().strftime('%B %-d, %Y')
    file_date         = _lms_label
    header_date_label = _lms_label
    total       = len(people)

    # ── Certification Journey (real data as of 2026-10-26) ──
    # Boot Camp / Social Selling / Capstone attendance+registration from Accelerate-Cert-Journey-*.xlsx,
    # Certified status+date from Accelerate-Certification-Report-*.xlsx — see load_cert_journey().
    # Falls back to the frozen 2026-10-21 placeholder only if neither real file is present at all
    # (e.g. a dev checkout without onboarding-data/).
    cert_journey, cert_meta = load_cert_journey()
    if not cert_journey and not cert_meta:
        try:
            with open(os.path.join('onboarding-data', '_cert_journey_placeholder.json'), encoding='utf-8') as _cj_f:
                cert_journey = json.load(_cj_f)
        except Exception:
            cert_journey = {}
        cert_meta = {'completionSource': 'LMS Accelerate Journey report', 'completionThrough': '2026-10-21',
                     'registrationSource': 'LMS registration report', 'registrationThrough': '2026-10-22'}
    cert_journey_json = json.dumps(cert_journey, separators=(',', ':'))
    cert_meta_json    = json.dumps(cert_meta, separators=(',', ':'))

    curric_ids   = json.dumps([c[0] for c in CURRICULA])
    curric_names = json.dumps({c[0]: c[1] for c in CURRICULA})

    # Page-level course map: exact LMS title → playbook page slug (confirmed 2026-07-09)
    _course_page = {
        'The KM Sales Experience':                                                          'understandingsalesforce',
        'Salesforce Live Workshop':                                                         'understandingsalesforce',
        'Mastering your Daily Sales Workflow':                                              'salesworkflow',
        'Why Konica Minolta: Selling the Value of Our Technology, Services and Solutions':  'coreportfolio',
        'bizhub One i-Series: Functionality, Ease of Use, Security':                        'coreportfolio',
        'Layered Security: Introducing the Model':                                          'coreportfolio',
        'User Authentication and Access Control':                                           'coreportfolio',
        'Document Security':                                                                'coreportfolio',
        'Introduction to Production Print':                                                 'coreportfolio',
        'Introducing BlueIrisIQ: Our Story':                                               'coreportfolio',
        'Managed IT (MIT) New Hire Kickstart':                                              'coreportfolio',
        'Identifying Real Sales Opportunities and Using the Lease Upgrade Sheet Part 1':    'prospectingskills',
        'Identifying Real Sales Opportunities and Using the Lease Upgrade Sheet Part 2':    'prospectingskills',
        'How to Focus on the Work That Matters Most: Highest Payoff Activities':            'prospectingskills',
        'Introduction to Microsoft Copilot':                                                'prospectingskills',
        'Microsoft Copilot: The Art of Prompting':                                          'prospectingskills',
        'How to Research and Write Messages That Win Meetings: Foundations of a Targeted Message':    'prospectingskills',
        'How to Research and Write Messages That Win Meetings: Leveraging the Persona Worksheet':     'prospectingskills',
        'Design Better Prospecting Calls with Disarm, Purpose, Question':                   'prospectingskills',
        'Build Better Voicemails: Structure, Clarity, Impact':                              'prospectingskills',
        'Prospecting Live Workshop':                                                        'prospectingskills',
        'Prospecting Foundations: Turning Skills into Action':                              'salesforceprospecting',
        'From Prep to Performance: Running Calls with Confidence':                          'callprep',
        'Discovery That Delivers: A Structured Approach to Better Sales Conversations':     'callprep',
        'Presentation Best Practices':                                                      'callprep',
        "Sales Math: Numbers Don't Lie":                                                    'workingwithnumbers',
        'Konica Minolta Premier Finance Leasing Live Workshop':                             'workingwithnumbers',
        'Managing and Moving Your Deals Forward':                                           'movingdeals',
        'How to Build Accurate Forecasts and Strong Pipelines':                             'pipelineownership',
    }
    _page_order = [
        'index', 'overview', 'welcome', 'understandingsalesforce',
        'salesworkflow',
        'coreportfolio',
        'prospectingskills', 'salesforceprospecting',
        'callprep', 'workingwithnumbers',
        'movingdeals', 'pipelineownership',
        'resources', 'managers',
    ]
    course_page_json = json.dumps(_course_page, ensure_ascii=False)
    page_order_json  = json.dumps(_page_order)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Accelerate Onboarding</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
<style>
  :root {{
    --bg:#0d1117; --surface:#161b22; --surface2:#1e2530; --border:#30363d;
    --accent:#3b82f6; --accent2:#6d28d9; --accent3:#d97706;
    --text:#e6edf3; --muted:#8b949e; --green:#22c55e; --red:#ef4444;
    --amber:#eab308; --teal:#14b8a6;
    --green-subtle:#22c55e22; --red-subtle:#ef444422; --amber-subtle:#eab30822;
    --font:'Segoe UI',system-ui,sans-serif;
  }}
  body.light-mode {{
    --bg:#f4f6fb; --surface:#ffffff; --surface2:#eef1f7; --border:#d0d7e8;
    --accent:#2563eb; --text:#1a1d27; --muted:#475569; --green:#059669;
    --red:#dc2626; --amber:#92400e; --teal:#0f766e;
    --green-subtle:#05966922; --red-subtle:#dc262622; --amber-subtle:#92400e22;
  }}
  body.light-mode select,body.light-mode input{{color-scheme:light;}}
  *{{box-sizing:border-box;margin:0;padding:0;}}
  body{{background:var(--bg);color:var(--text);font-family:var(--font);min-height:100vh;transition:background .2s,color .2s;}}

  /* ── Header ── */
  .header{{padding:20px 28px 16px;border-bottom:1px solid var(--border);display:grid;grid-template-columns:1fr auto 1fr;align-items:center;gap:12px;}}
  .header-left{{display:flex;align-items:center;gap:16px;}}
  .header-center{{display:flex;justify-content:center;align-items:center;}}
  .header-right{{display:flex;justify-content:flex-end;align-items:center;gap:8px;}}
  .kma-logo{{height:38px;width:auto;display:block;}}
  .kma-logo-light{{display:none;}}
  .light-mode .kma-logo-dark{{display:none;}}
  .light-mode .kma-logo-light{{display:block;}}
  .header h1{{font-size:18px;font-weight:700;letter-spacing:.3px;}}
  .header h1 span{{color:var(--muted);font-weight:400;}}
  .header-date{{font-size:11px;color:var(--muted);margin-top:2px;}}
  .data-badge{{font-size:11px;color:var(--muted);background:var(--surface2);border:1px solid var(--border);border-radius:20px;padding:3px 10px;white-space:nowrap;}}

  /* ── Buttons ── */
  .btn-theme{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:5px 12px;font-size:12px;cursor:pointer;transition:all .15s;}}
  .btn-theme:hover{{border-color:var(--accent);color:var(--text);}}
  .btn-export{{background:var(--accent);border:1px solid var(--accent);color:#fff;border-radius:6px;padding:5px 12px;font-size:12px;cursor:pointer;font-weight:600;}}
  .btn-export:hover{{opacity:.88;}}
  .export-drop{{position:relative;}}
  .export-menu{{position:absolute;top:calc(100% + 6px);right:0;background:var(--surface);border:1px solid var(--border);border-radius:8px;min-width:210px;box-shadow:0 4px 24px rgba(0,0,0,.28);display:none;z-index:200;overflow:visible;}}
  .export-menu.open{{display:block;}}
  .export-item{{display:block;width:100%;text-align:left;padding:10px 14px;font-size:13px;color:var(--text);background:transparent;border:none;cursor:pointer;transition:background .1s;font-family:inherit;}}
  .export-item:hover{{background:var(--surface2);}}
  .export-parent{{position:relative;display:flex;justify-content:space-between;align-items:center;padding:10px 14px;font-size:13px;color:var(--text);cursor:default;transition:background .1s;}}
  .export-parent:hover{{background:var(--surface2);}}
  .export-chevron{{font-size:11px;color:var(--muted);margin-left:10px;}}
  .export-submenu{{position:absolute;right:100%;top:0;background:var(--surface);border:1px solid var(--border);border-radius:8px;min-width:90px;box-shadow:0 4px 24px rgba(0,0,0,0.28);display:none;z-index:201;overflow:hidden;margin-right:4px;}}
  .export-parent:hover .export-submenu{{display:block;}}

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
  .btn-group{{display:inline-flex;border:1px solid var(--border);border-radius:6px;overflow:hidden;}}
  .btn-group button{{background:transparent;border:none;border-right:1px solid var(--border);color:var(--muted);padding:6px 12px;font-size:12px;cursor:pointer;transition:background .15s,color .15s;white-space:nowrap;}}
  .btn-group button:last-child{{border-right:none;}}
  .btn-group button:hover{{background:var(--surface2);color:var(--text);}}
  .btn-group button.active{{background:var(--accent);color:#fff;font-weight:600;}}

  .sort-btn{{background:transparent;border:1px solid var(--border);color:var(--muted);border-radius:6px;padding:4px 10px;font-size:11px;cursor:pointer;transition:all .15s;white-space:nowrap;}}
  .sort-btn:hover{{border-color:var(--accent);color:var(--text);}}
  .sort-btn.active{{border-color:var(--accent);color:var(--accent);background:var(--accent)11;}}
  .mgr-group-hdr-row td{{padding:0;}}
  .mgr-group-hdr{{display:flex;align-items:center;justify-content:space-between;padding:8px 14px;cursor:pointer;background:var(--surface2);border-bottom:1px solid var(--border);user-select:none;}}
  .mgr-group-hdr:hover{{background:var(--border);}}
  .mgr-group-hdr.open .mgr-chevron{{transform:rotate(180deg);}}
  .mgr-chevron{{font-size:10px;color:var(--muted);transition:transform .15s;}}
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
  table{{border-collapse:collapse;width:100%;min-width:680px;}}
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
  td{{padding:9px 12px;vertical-align:middle;font-size:13px;}}
  td.name-cell{{font-weight:600;white-space:nowrap;display:flex;align-items:center;gap:8px;}}
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
  .modal-card{{background:var(--surface);border:1px solid var(--border);border-radius:12px;width:100%;max-width:860px;position:relative;}}
  .modal-close{{position:absolute;top:14px;right:16px;background:transparent;border:none;color:var(--muted);font-size:18px;cursor:pointer;line-height:1;padding:4px 8px;border-radius:4px;}}
  .modal-close:hover{{color:var(--text);background:var(--surface2);}}
  .modal-header{{padding:20px 24px 16px;border-bottom:1px solid var(--border);}}
  .modal-name{{font-size:18px;font-weight:700;margin-bottom:4px;}}
  .modal-meta{{font-size:12px;color:var(--muted);display:flex;flex-wrap:wrap;gap:12px;}}
  .modal-body{{padding:20px 24px;}}

  /* ── Report-card risk banner + metric grid ── */
  .rc-riskbar{{display:flex;align-items:center;gap:12px;padding:14px 18px;border-radius:10px;margin-bottom:18px;flex-wrap:wrap;}}
  .rc-riskbar.risk-red{{background:var(--red-subtle);}}
  .rc-riskbar.risk-amber{{background:var(--amber-subtle);}}
  .rc-riskbar.risk-green{{background:var(--green-subtle);}}
  .rc-riskbar .rc-risklabel{{font-size:14px;font-weight:800;}}
  .rc-riskbar.risk-red .rc-risklabel{{color:var(--red);}}
  .rc-riskbar.risk-amber .rc-risklabel{{color:var(--amber);}}
  .rc-riskbar.risk-green .rc-risklabel{{color:var(--green);}}
  .rc-riskbar .rc-riskaction{{font-size:12.5px;color:var(--text);flex:1;min-width:200px;}}
  .metric-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:10px;margin-bottom:20px;}}
  .metric-box{{background:var(--surface2);border-radius:8px;padding:10px 12px;}}
  .metric-box .mb-label{{font-size:9.5px;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);margin-bottom:4px;}}
  .metric-box .mb-value{{font-size:17px;font-weight:800;line-height:1.15;}}
  .metric-box .mb-sub{{font-size:10.5px;color:var(--muted);margin-top:2px;}}
  .sched-table{{width:100%;border-collapse:collapse;margin-bottom:20px;}}
  .sched-table th{{text-align:left;padding:5px 10px;font-size:9.5px;text-transform:uppercase;letter-spacing:.4px;color:var(--muted);border-bottom:1px solid var(--border);}}
  .sched-table th.c{{text-align:center;}}
  .sched-table td{{padding:7px 10px;font-size:12.5px;border-bottom:1px solid var(--border);}}
  .sched-table td.c{{text-align:center;}}
  .sched-table tr:last-child td{{border-bottom:none;}}
  .sched-table .sched-actual{{font-weight:700;}}

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
  @media(max-width:900px){{.charts{{grid-template-columns:1fr;}}}}
  @media(max-width:480px){{.chart-wrap{{height:160px;}}}}
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

  /* ── Tabs ── */
  .tabs{{display:flex;gap:2px;padding:0 28px;border-bottom:1px solid var(--border);background:var(--surface);}}
  .tab-btn{{padding:11px 18px;font-size:13px;font-weight:600;color:var(--muted);background:transparent;border:none;border-bottom:2px solid transparent;cursor:pointer;font-family:inherit;transition:color .15s,border-color .15s;}}
  .tab-btn:hover{{color:var(--text);}}
  .tab-btn.active{{color:var(--text);border-bottom-color:var(--accent);}}

  /* ── Certification Journey tab ── */
  .cj-chip{{display:inline-block;font-size:10.5px;font-weight:700;padding:3px 9px;border-radius:5px;white-space:nowrap;line-height:1.4;}}
  .cj-chip.done{{background:#15803d;color:#fff;}}
  .cj-chip.warn{{background:#b45309;color:#fff;}}
  .cj-chip.none{{background:var(--surface2);color:var(--muted);font-weight:600;}}
  .cj-chip.reg{{background:#1d4ed8;color:#fff;}}
  .cj-sub{{display:block;font-size:10px;color:var(--muted);margin-top:2px;font-weight:400;}}
  td.cj-c, th.cj-c{{text-align:center;}}
  .cj-stepper{{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:20px;}}
  .cj-step{{background:var(--surface2);border:1px solid var(--border);border-radius:10px;padding:12px;text-align:center;}}
  .cj-step.done{{border-color:#15803d;background:var(--green-subtle);}}
  .cj-step.active{{border-color:#1d4ed8;}}
  .cj-step.warn{{border-color:#b45309;background:var(--amber-subtle);}}
  .cj-step.reg{{border-color:#1d4ed8;}}
  .cj-step.reg .cj-num{{background:#1d4ed8;}}
  .cj-num{{width:26px;height:26px;border-radius:50%;background:#6b7280;color:#fff;font-size:12px;font-weight:800;display:inline-flex;align-items:center;justify-content:center;margin-bottom:6px;}}
  .cj-step.done .cj-num{{background:#15803d;}}
  .cj-step.active .cj-num{{background:#1d4ed8;}}
  .cj-step.warn .cj-num{{background:#b45309;}}
  .cj-step-label{{font-size:12px;font-weight:700;}}
  .cj-step-status{{font-size:11px;color:var(--muted);margin-top:2px;}}
</style>
</head>
<body>

<!-- Info popover -->
<div class="info-popover" id="info-popover"></div>

<!-- Header -->
<div class="header">
  <div class="header-left">
    <div>
      <h1>Accelerate Onboarding</h1>
      <div class="header-date">Data through {header_date_label}</div>
    </div>
    <span class="data-badge" id="data-badge">{total} learners &middot; updated {file_date}</span>
  </div>
  <div class="header-center">
    <img src="https://jasonackerman1.github.io/playbook-dashboard/KMA-wht.svg" class="kma-logo kma-logo-dark" alt="KM Academy">
    <img src="https://jasonackerman1.github.io/playbook-dashboard/KMA-drk.svg" class="kma-logo kma-logo-light" alt="KM Academy">
  </div>
  <div class="header-right">
    <div class="export-drop print-hide" id="export-drop">
      <button class="btn-export" onclick="toggleExportDrop()">&#128438; Export &#9660;</button>
      <div class="export-menu" id="export-menu">
        <div class="export-parent">Full Report<span class="export-chevron">&#8249;</span><div class="export-submenu"><button class="export-item" onclick="runExport('full')">PDF</button><button class="export-item" onclick="runExportXLSX('full')">Excel</button></div></div>
        <div class="export-parent">Overdue Only<span class="export-chevron">&#8249;</span><div class="export-submenu"><button class="export-item" onclick="runExport('overdue')">PDF</button><button class="export-item" onclick="runExportXLSX('overdue')">Excel</button></div></div>
        <div class="export-parent">Manager Summary<span class="export-chevron">&#8249;</span><div class="export-submenu"><button class="export-item" onclick="runExport('manager-summary')">PDF</button><button class="export-item" onclick="runExportXLSX('manager-summary')">Excel</button></div></div>
        <div class="export-parent">Certification Journey<span class="export-chevron">&#8249;</span><div class="export-submenu"><button class="export-item" onclick="runCertExport()">PDF</button><button class="export-item" onclick="runCertExportXLSX()">Excel</button></div></div>
      </div>
    </div><span class="info-btn print-hide" onclick="showInfo(event,'export')">?</span>
    <button class="btn-theme" id="btn-theme" onclick="toggleTheme()">&#9728; Light</button>
  </div>
</div>

<!-- Tabs -->
<div class="tabs print-hide">
  <button class="tab-btn active" data-tab="onboarding" onclick="switchTab('onboarding')">Onboarding</button>
  <button class="tab-btn" data-tab="cert" onclick="switchTab('cert')">Certification Journey</button>
</div>

<!-- Filters -->
<div class="filters">
  <span class="filter-label">Location</span>
  <div class="btn-group" id="location-btns">
    <button class="active" onclick="setLocation('all',this)">All</button>
    <button onclick="setLocation('us',this)">&#127482;&#127480; US</button>
    <button onclick="setLocation('ca',this)">&#127464;&#127462; Canada</button>
  </div>
  <span class="filter-label">Market</span>
  <select id="f-market" onchange="applyFilters()"><option value="">All Markets</option></select>
  <span class="filter-label">Hire Status</span>
  <div class="btn-group" id="tenure-btns">
    <button class="active" onclick="setTenure('all',this)">All</button>
    <button onclick="setTenure('new',this)" title="Hired within the last 65 days">New Hires</button>
    <button onclick="setTenure('tenured',this)" title="Hired more than 65 days ago">Tenured</button>
  </div>
  <span id="f-status-wrap" style="display:contents"><span class="filter-label">Status</span>
  <select id="f-status" onchange="applyFilters()">
    <option value="">All</option>
    <option value="Completed">Completed</option>
    <option value="On Track">On Track</option>
    <option value="Overdue">Overdue</option>
  </select></span>
  <span class="filter-label">Group</span>
  <button class="btn-reset" id="test-group-toggle" onclick="toggleTestGroup(this)">Show Early Access Cohort</button>
  <button class="btn-reset" onclick="resetFilters()">Reset</button>
  <span class="info-btn" onclick="showInfo(event,'filters-info')" style="margin-left:4px;">?</span>
  <span class="result-count" id="result-count"></span>
</div>

<div id="tab-onboarding">
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
    <div class="stat-sub" id="s-ontrack-sub"></div>
  </div>
  <div class="stat">
    <div class="stat-label">Completed <span class="info-btn" onclick="showInfo(event,'completed')">?</span></div>
    <div class="stat-value green" id="s-completed">&#8212;</div>
    <div class="stat-sub" id="s-completed-sub"></div>
  </div>
  <div class="stat">
    <div class="stat-label">Days to First Sale <span class="info-btn" onclick="showInfo(event,'first-sale')">?</span></div>
    <div class="stat-value amber" id="s-sales">TBD</div>
    <div class="stat-sub">avg from program start to close</div>
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
      <div class="section-title">Progress Report <span class="info-btn" onclick="showInfo(event,'heatmap')">?</span></div>
      <div class="section-hint">Click any row to see full detail &mdash; curriculum breakdown, course checklist &amp; playbook activity</div>
    </div>
    <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
      <div style="display:flex;gap:6px;align-items:center;">
        <button class="sort-btn active" id="view-individual" onclick="setTableView('individual')">Individual</button>
        <button class="sort-btn" id="view-manager" onclick="setTableView('manager')">By Manager</button>
        <span class="info-btn" onclick="showInfo(event,'view-toggle')" style="margin-left:2px;">?</span>
      </div>
      <span class="filter-label">Sort</span>
      <select id="f-sort" onchange="applyFilters()">
        <option value="name">Name A→Z</option>
        <option value="pct-desc">Completion High→Low</option>
        <option value="pct-asc">Completion Low→High</option>
        <option value="days-asc">Most Urgent First</option>
      </select>
      <input type="text" id="table-search" oninput="filterTableRows()" placeholder="Search name..." style="font-size:12px;padding:4px 10px;width:180px;background:var(--surface2);border:1px solid var(--border);color:var(--text);border-radius:6px;outline:none;">
    </div>
  </div>
  <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px;padding:10px 16px;background:var(--surface);border:1px solid var(--border);border-radius:8px;font-size:11px;color:var(--muted);">
    <div>
      <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-bottom:5px;">Gap</div>
      <div style="display:flex;flex-wrap:wrap;gap:8px;">
        <span class="leg"><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#15803d;margin-right:4px;flex-shrink:0;"></span>On pace or ahead (0% or less)</span>
        <span class="leg"><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#b45309;margin-right:4px;flex-shrink:0;"></span>Check-in (1&ndash;40%)</span>
        <span class="leg"><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#b91c1c;margin-right:4px;flex-shrink:0;"></span>Coaching needed (41%+)</span>
      </div>
    </div>
    <div>
      <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-bottom:5px;">Curricula &#9632;</div>
      <div style="display:flex;flex-wrap:wrap;gap:8px;">
        <span class="leg"><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#15803d;margin-right:4px;flex-shrink:0;"></span>Done</span>
        <span class="leg"><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#1d4ed8;margin-right:4px;flex-shrink:0;"></span>In Progress</span>
        <span class="leg"><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#b91c1c;margin-right:4px;flex-shrink:0;"></span>Past Due</span>
        <span class="leg"><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#6b7280;margin-right:4px;flex-shrink:0;"></span>Not Started</span>
      </div>
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

</div><!-- /tab-onboarding -->

<!-- ═══════════ Certification Journey tab ═══════════ -->
<div id="tab-cert" style="display:none">

  <div class="stats">
    <div class="stat">
      <div class="stat-label">Total Enrolled <span class="info-btn" onclick="showInfo(event,'cert-total')">?</span></div>
      <div class="stat-value blue" id="cj-total">&#8212;</div>
    </div>
    <div class="stat">
      <div class="stat-label">Onboarding Complete <span class="info-btn" onclick="showInfo(event,'cert-onboarding')">?</span></div>
      <div class="stat-value" id="cj-onb">&#8212;</div>
      <div class="stat-sub" id="cj-onb-sub"></div>
    </div>
    <div class="stat">
      <div class="stat-label">Advanced Boot Camp <span class="info-btn" onclick="showInfo(event,'cert-bootcamp')">?</span></div>
      <div class="stat-value" id="cj-boot">&#8212;</div>
      <div class="stat-sub" id="cj-boot-sub"></div>
    </div>
    <div class="stat">
      <div class="stat-label">Social Selling <span class="info-btn" onclick="showInfo(event,'cert-social')">?</span></div>
      <div class="stat-value" id="cj-soc">&#8212;</div>
      <div class="stat-sub" id="cj-soc-sub"></div>
    </div>
    <div class="stat">
      <div class="stat-label">Capstone Workshop <span class="info-btn" onclick="showInfo(event,'cert-capstone')">?</span></div>
      <div class="stat-value" id="cj-cap">&#8212;</div>
      <div class="stat-sub" id="cj-cap-sub"></div>
    </div>
    <div class="stat">
      <div class="stat-label">Certified <span class="info-btn" onclick="showInfo(event,'cert-certified')">?</span></div>
      <div class="stat-value green" id="cj-cert">&#8212;</div>
      <div class="stat-sub" id="cj-cert-sub"></div>
    </div>
  </div>

  <div class="charts">
    <div class="chart-card">
      <div class="chart-title">Certified by Market <span class="info-btn" onclick="showInfo(event,'cert-market-certified')">?</span></div>
      <div class="chart-wrap"><canvas id="certMarketCertChart"></canvas></div>
    </div>
    <div class="chart-card">
      <div class="chart-title">Attendance by Market <span class="info-btn" onclick="showInfo(event,'cert-market-attendance')">?</span></div>
      <div class="chart-wrap"><canvas id="certMarketAttChart"></canvas></div>
    </div>
  </div>

  <div class="section">
    <div class="section-header">
      <div>
        <div class="section-title">Certification Journey Tracker <span class="info-btn" onclick="showInfo(event,'cert-table')">?</span></div>
        <div class="section-hint">Click any row to see each milestone in detail <span id="cert-asof"></span></div>
      </div>
      <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
        <div style="display:flex;gap:6px;align-items:center;">
          <button class="sort-btn active" id="cert-view-individual" onclick="setCertTableView('individual')">Individual</button>
          <button class="sort-btn" id="cert-view-manager" onclick="setCertTableView('manager')">By Manager</button>
          <span class="info-btn" onclick="showInfo(event,'cert-view-toggle')" style="margin-left:2px;">?</span>
        </div>
        <span class="filter-label">Milestone</span>
        <select id="cert-ms" onchange="renderCertTable()">
          <option value="">All</option>
          <optgroup label="Advanced Boot Camp">
            <option value="bootCamp:done">Attended</option>
            <option value="bootCamp:reg">Registered</option>
            <option value="bootCamp:none">Not registered</option>
          </optgroup>
          <optgroup label="Social Selling Workshop">
            <option value="social:done">Attended</option>
            <option value="social:reg">Registered</option>
            <option value="social:none">Not registered</option>
          </optgroup>
          <optgroup label="Capstone Workshop">
            <option value="capstone:done">Attended</option>
            <option value="capstone:reg">Registered</option>
            <option value="capstone:none">Not registered</option>
          </optgroup>
          <optgroup label="Certified">
            <option value="certification:done">Yes</option>
            <option value="certification:none">No</option>
          </optgroup>
        </select>
        <span class="filter-label">Sort</span>
        <select id="cert-sort" onchange="renderCertTable()">
          <option value="name">Name A&rarr;Z</option>
          <option value="ms-desc">Most Milestones</option>
          <option value="ms-asc">Fewest Milestones</option>
          <option value="onb-desc">Onboarding % High&rarr;Low</option>
        </select>
        <input type="text" id="cert-search" oninput="renderCertTable()" placeholder="Search name or manager..." style="font-size:12px;padding:4px 10px;width:200px;background:var(--surface2);border:1px solid var(--border);color:var(--text);border-radius:6px;outline:none;">
      </div>
    </div>
    <div class="legend">
      <span class="leg"><span class="leg-dot" style="background:#15803d"></span>Completed</span>
      <span class="leg"><span class="leg-dot" style="background:#1d4ed8"></span>Registered (booked into a session)</span>
      <span class="leg"><span class="leg-dot" style="background:#6b7280"></span>Not registered</span>
    </div>
    <div class="table-wrap">
      <table>
        <thead id="cert-head"></thead>
        <tbody id="cert-body"></tbody>
      </table>
    </div>
  </div>
</div><!-- /tab-certification -->

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
const SALES_MAP = {sales_map_json};
const COHORT_TOTAL = {cohort_total};
/* Certification Journey records. Keyed by email (any capitalisation).
   Anyone not listed has not completed that milestone. Shape:
   {{ "name@kmbs.konicaminolta.us": {{
       bootCamp:      {{ attended: true, date: "YYYY-MM-DD",           // completion report
                        registered: true, sessionDate: "YYYY-MM-DD", sessionEnd: "YYYY-MM-DD",   // registration report: the session booked
                        classId: "1376018", enrolledDate: "YYYY-MM-DD" }},
       socialSelling: {{ attended: true, date: "YYYY-MM-DD",
                        registered: true, sessionDate: "YYYY-MM-DD", sessionEnd: "YYYY-MM-DD" }},
       capstone:      {{ attended: true, date: "YYYY-MM-DD",           // Accelerate: Capstone Presentation Certification Workshop
                        registered: true, sessionDate: "YYYY-MM-DD", sessionEnd: "YYYY-MM-DD" }},
       certification: {{ completed: true, date: "YYYY-MM-DD" }}   // "Accelerate Certification" LMS course
   }} }}
   Being assigned a class in the LMS (enrolledDate) without a booked session does NOT count as
   Registered — only a real sessionDate from the registration report does. (Fixed 2026-10-23 per
   Resmie: the original build mistakenly treated LMS assignment as registration.)               */
const CERT_JOURNEY = {cert_journey_json};
/* Source of the completion records above. */
const CERT_META = {cert_meta_json};
const COURSE_PAGE = {course_page_json};
const PAGE_ORDER = {page_order_json};
const CURRIC_IDS = {curric_ids};
const CURRIC_NAMES = {curric_names};
/* The designed pacing schedule, straight from each curriculum's LMS title
   (e.g. "Sales Skills (Weeks 2 & 3)"). [startDay, endDay] in program days,
   1-indexed, inclusive -- used to say which curriculum someone SHOULD be
   focused on right now, independent of both deadlines and raw % complete. */
const CURRIC_WEEK_RANGE = {{
  "ACCELERATE_GS": [1, 7],
  "ACCELERATE_SW": [1, 7],
  "ACCELERATE_CP": [1, 7],
  "ACCELERATE_P":  [1, 14],
  "ACCELERATE_SS": [8, 21],
  "ACCELERATE_PM": [22, 35],
}};
const PROGRAM_DAYS = {PROGRAM_DAYS};
const PLAYBOOK_CURRIC = {{
  'welcome':                 'ACCELERATE_GS',
  'understandingsalesforce': 'ACCELERATE_GS',
  'salesworkflow':           'ACCELERATE_SW',
  'coreportfolio':           'ACCELERATE_CP',
  'prospectingskills':       'ACCELERATE_P',
  'salesforceprospecting':   'ACCELERATE_P',
  'callprep':                'ACCELERATE_SS',
  'workingwithnumbers':      'ACCELERATE_SS',
  'movingdeals':             'ACCELERATE_PM',
  'pipelineownership':       'ACCELERATE_PM',
}};
const TODAY = new Date();
TODAY.setHours(0,0,0,0);

const TLG_SET = new Set([
  'jason ackerman','bianca davis','james parker','resmie biba',
  'chris curtis','sara thompson','jeremy macbean','bradley pierce',
  'laura sefcik','samantha maresca','staci musco','cj homer','rich moore','dale kinsey',
  'john lechner','resmie nesimi',"samantha d'angelo",'bianca dipasquale','doug falk'
]);

let hideTLG = true;
let testGroupFilter = 'hide'; // 'all' | 'hide' | 'only' — hidden by default
let locationFilter = 'all';  // 'all' | 'us' | 'ca'
// Fixed roster of the original June 4 cohort — stable even when the LMS
// reassigns assignDate in future pulls (which happened July 2026).
const TEST_GROUP_NAMES = new Set([
  'Alford Wood', 'Anthony Petretta', 'Christopher Melton', 'Daniel Zepeda',
  'Dillon Basnight', 'Dre Ante Noel', 'Dylan McDonough', 'Emily Olstein',
  'Frances Rogers', 'Gregory Christie', 'Jared Poplawski', 'Kelsey Carrington',
  'Kevin Kojouri', 'Kyle Baker', 'Meagan Blythe', 'Michael Shields',
  'Paolo Castellon', 'Patrick Meagher', 'RICHARD ALLEN', 'Randahl Bradley',
  'Sophia Krawczak', 'Spencer Larson', 'Taylor Hill', 'Tracy Hendricks',
  'Ty Foster', 'Valencia Gilford', 'William Birkett'
]);

function toggleTestGroup(btn) {{
  testGroupFilter = (testGroupFilter === 'hide') ? 'only' : 'hide';
  btn.textContent = (testGroupFilter === 'only') ? 'Hide Early Access Cohort' : 'Show Early Access Cohort';
  applyFilters();
}}
function setLocation(val, btn) {{
  locationFilter = val;
  document.querySelectorAll('#location-btns button').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  applyFilters();
}}

function populateMarketDropdown() {{
  const sel = document.getElementById('f-market');
  const current = sel.value;

  // Markets available given the current cohort + location filters (ignore market filter itself)
  const visible = PEOPLE.filter(p => {{
    if (hideTLG && TLG_SET.has(p.name.toLowerCase())) return false;
    if (testGroupFilter === 'hide' && TEST_GROUP_NAMES.has(p.name)) return false;
    if (testGroupFilter === 'only' && !TEST_GROUP_NAMES.has(p.name)) return false;
    if (locationFilter === 'us' && p.isCanada) return false;
    if (locationFilter === 'ca' && !p.isCanada) return false;
    if (!passesTenure(p)) return false;
    return true;
  }});
  const markets = [...new Set(visible.map(p => p.market))].sort();

  sel.innerHTML = '<option value="">All Markets</option>';
  markets.forEach(m => {{
    const o = document.createElement('option');
    o.value = m; o.textContent = m;
    sel.appendChild(o);
  }});

  // Preserve selection if still valid, otherwise reset to All
  sel.value = markets.includes(current) ? current : '';
}}

let filtered = [];
let marketChartObj = null;
let curricChartObj = null;
let tableSort = {{col: 'overall', dir: 'asc'}};
let tableView = 'individual';

/* ── Utilities ── */
function pct2color(p) {{
  if (p === 0) return {{bg:'#f1efe8', fg:'#888780'}};
  if (p < 40)  return {{bg:'#FCEBEB', fg:'#791F1F'}};
  if (p < 75)  return {{bg:'#FAEEDA', fg:'#633806'}};
  return {{bg:'#EAF3DE', fg:'#27500A'}};
}}

/* ── Deadline engine ──
   Each item carries its own LMS-assigned due date (item.req, or null if the
   item has no fixed deadline). A curriculum's days-remaining is the soonest
   due date among its own INCOMPLETE items. */
function itemDaysLeft(item) {{
  if (!item.req) return null;
  return Math.round((new Date(item.req) - TODAY) / 86400000);
}}
function curricDaysLeft(c) {{
  if (!c || c.complete) return null;
  const vals = c.items.filter(i => !i.done && i.req).map(itemDaysLeft);
  return vals.length ? Math.min(...vals) : null;
}}
function curricTrackStatus(c) {{
  if (!c) return 'Unknown';
  if (c.complete) return 'Completed';
  const dl = curricDaysLeft(c);
  if (dl === null) return 'On Track';
  return dl <= 0 ? 'Overdue' : 'On Track';
}}

function curricPillStyle(c, p) {{
  if (!c) return 'background:#6b7280;color:#fff';
  if (c.complete) return 'background:#15803d;color:#fff';
  const dl = curricDaysLeft(c);
  if (dl !== null && dl <= 0) return 'background:#b91c1c;color:#fff';
  if (c.pct > 0)  return 'background:#1d4ed8;color:#fff';
  return 'background:#6b7280;color:#fff';
}}
function curricDotBg(c, p) {{
  if (!c) return '#6b7280';
  if (c.complete) return '#15803d';
  const dl = curricDaysLeft(c);
  if (dl !== null && dl <= 0) return '#b91c1c';
  if (c.pct > 0) return '#1d4ed8';
  return '#6b7280';
}}

function overallPillStyle(p) {{
  if (p.overallDone) return 'background:#15803d;color:#fff;font-weight:700';
  if (computeStatus(p) === 'Overdue') return 'background:#b91c1c;color:#fff;font-weight:700';
  if (p.overallPct > 0) return 'background:#1d4ed8;color:#fff;font-weight:700';
  return 'background:#6b7280;color:#fff;font-weight:700';
}}

function computeStatus(p) {{
  if (p.overallDone) return 'Completed';
  const ids = Object.keys(p.curricula);
  if (!ids.length) return 'Unknown';
  const anyOverdue = ids.some(cid => curricTrackStatus(p.curricula[cid]) === 'Overdue');
  return anyOverdue ? 'Overdue' : 'On Track';
}}

function overdueCount(p) {{
  return Object.values(p.curricula).filter(c => curricTrackStatus(c) === 'Overdue').length;
}}

function worstOverdueCurric(p) {{
  let worstCid = null, worstDl = null;
  Object.keys(p.curricula).forEach(cid => {{
    const dl = curricDaysLeft(p.curricula[cid]);
    if (dl === null) return;
    if (worstDl === null || dl < worstDl) {{ worstDl = dl; worstCid = cid; }}
  }});
  return worstCid ? {{cid: worstCid, dl: worstDl}} : null;
}}
function maxDaysOverdue(p) {{
  const w = worstOverdueCurric(p);
  if (!w || w.dl > 0) return 0;
  return -w.dl;
}}
function interventionInfo(p) {{
  const days = maxDaysOverdue(p);
  const w = worstOverdueCurric(p);
  const curricName = w ? CURRIC_NAMES[w.cid] : null;
  if (days >= 20) return {{tier: 2, days, label: 'Intervention 2', level: 'red', curricName}};
  if (days >= 10) return {{tier: 1, days, label: 'Intervention 1', level: 'amber', curricName}};
  return {{tier: 0, days, label: null, level: null, curricName: null}};
}}

function soonestDaysLeft(p) {{
  const vals = Object.values(p.curricula).map(curricDaysLeft).filter(v => v !== null);
  if (vals.length) return Math.min(...vals);
  return computedDaysLeft(p);
}}

function daysLeft(p) {{
  return soonestDaysLeft(p);
}}

/* Kept for the Expected%/Gap pacing metric only (not for Overdue/Days Left,
   which now come from each curriculum's own LMS due dates above). */
function daysElapsed(p) {{
  if (!p.assignDate) return 0;
  const start = new Date(p.assignDate);
  return Math.max(0, Math.round((TODAY - start) / 86400000));
}}
function computedDaysLeft(p) {{
  if (!p.assignDate) return null;
  return PROGRAM_DAYS - daysElapsed(p);
}}
function curricExpectedPct(cid, d) {{
  const range = CURRIC_WEEK_RANGE[cid] || [1, PROGRAM_DAYS];
  if (d >= range[1]) return 100;
  if (d < range[0]) return 0;
  const span = range[1] - range[0] + 1;
  const progressed = Math.min(span, d - range[0] + 1);
  return Math.round(progressed / span * 100);
}}
function curricDueDate(p, cid) {{
  if (!p.assignDate) return null;
  const range = CURRIC_WEEK_RANGE[cid] || [1, PROGRAM_DAYS];
  const dt = new Date(p.assignDate + 'T00:00:00');
  dt.setDate(dt.getDate() + range[1] - 1);
  return dt;
}}
function expectedPct(p) {{
  const d = daysElapsed(p);
  const ids = Object.keys(p.curricula);
  if (!ids.length) return 0;
  let totalItems = 0, expectedDone = 0;
  ids.forEach(cid => {{
    const c = p.curricula[cid];
    const total = c.total || 0;
    if (!total) return;
    totalItems += total;
    expectedDone += total * (curricExpectedPct(cid, d) / 100);
  }});
  if (!totalItems) return 0;
  return Math.min(100, Math.round(expectedDone / totalItems * 100));
}}
function gapPct(p) {{
  return expectedPct(p) - p.overallPct;
}}
/* Where they should be, curriculum-by-curriculum, per the designed weekly
   schedule -- independent of both LMS due dates (Days Left) and raw %
   complete (Biggest Gap). */
function curricScheduleStatus(cid, d) {{
  const range = CURRIC_WEEK_RANGE[cid];
  if (!range) return 'active';
  if (d < range[0]) return 'upcoming';
  if (d > range[1]) return 'due';
  return 'active';
}}
function expectedFocus(p) {{
  const d = daysElapsed(p);
  const active = [];
  const behindSchedule = [];
  CURRIC_IDS.forEach(cid => {{
    const c = p.curricula[cid];
    if (!c || c.complete) return;
    const st = curricScheduleStatus(cid, d);
    if (st === 'active') active.push(cid);
    else if (st === 'due') behindSchedule.push(cid);
  }});
  if (behindSchedule.length) return {{mode: 'behind', ids: behindSchedule}};
  if (active.length) return {{mode: 'active', ids: active}};
  return {{mode: 'ahead', ids: []}};
}}
function gapStyle(g) {{
  if (g > 40) return 'background:#b91c1c;color:#fff;font-weight:700';
  if (g > 0)  return 'background:#a16207;color:#fff;font-weight:700';
  return 'background:#15803d;color:#fff;font-weight:700';
}}
function pctPillStyle(pct) {{
  if (pct === 0)  return 'background:#b91c1c;color:#fff;font-weight:600';
  if (pct < 50)   return 'background:#a16207;color:#fff;font-weight:600';
  return 'background:#6b7280;color:#fff;font-weight:600';
}}
function urlSlug(url) {{
  const m = (url || '').match(/\/([^\/]+?)(?:\.html)?(?:\?.*)?$/);
  if (!m || !m[1] || m[1] === 'accelerate_sales_playbook' || m[1] === '') return 'index';
  return m[1].toLowerCase();
}}
function pbEngagement(p) {{
  const visits = (p.playbook || []).filter(v => !p.assignDate || v.date >= p.assignDate);
  if (!visits.length) {{
    const level = p.overallPct > 0 ? 'alert' : 'none';
    return {{totalVisits:0, uniquePages:0, visitedCurricula:{{}}, mismatches:[], level}};
  }}
  const totalVisits = visits.length;
  const slugs = visits.map(v => urlSlug(v.url));
  const uniquePages = new Set(slugs).size;
  const visitedCurricula = {{}};
  slugs.forEach(slug => {{
    const cid = PLAYBOOK_CURRIC[slug];
    if (cid) visitedCurricula[cid] = (visitedCurricula[cid] || 0) + 1;
  }});
  const mismatches = CURRIC_IDS.filter(cid => {{
    const c = p.curricula[cid];
    return c && c.done > 0 && !visitedCurricula[cid];
  }});
  const level = mismatches.length > 0 ? 'partial' : 'active';
  return {{totalVisits, uniquePages, visitedCurricula, mismatches, level}};
}}
function pbSummaryHtml(p) {{
  const eng = pbEngagement(p);
  let html = '<div style="margin-bottom:16px;padding-bottom:16px;border-bottom:1px solid var(--border);">';
  html += '<div style="font-size:12px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px;">Playbook Engagement</div>';
  if (eng.totalVisits === 0) {{
    const msg = p.overallPct > 0
      ? '<span style="color:var(--red)">&#9888; Completing courses with no playbook visits recorded</span>'
      : '<span style="color:var(--muted)">No playbook visits recorded yet</span>';
    html += '<div style="font-size:13px;margin-bottom:8px;">' + msg + '</div>';
  }} else {{
    html += '<div style="font-size:13px;color:var(--muted);margin-bottom:8px;">' + eng.totalVisits + ' visit' + (eng.totalVisits!==1?'s':'') + ' &middot; ' + eng.uniquePages + ' unique page' + (eng.uniquePages!==1?'s':'') + '</div>';
  }}
  html += '<div style="display:flex;flex-wrap:wrap;gap:5px;margin-bottom:8px;">';
  CURRIC_IDS.forEach(cid => {{
    const c = p.curricula[cid];
    const hasDone = c && c.done > 0;
    const visited = !!eng.visitedCurricula[cid];
    let bg, color, icon;
    if (visited)       {{ bg='var(--green-subtle)'; color='var(--green)';  icon='&#10003;'; }}
    else if (hasDone)  {{ bg='var(--red-subtle)';   color='var(--red)';    icon='&#10007;'; }}
    else               {{ bg='var(--surface2)';      color='var(--muted)';  icon='&middot;'; }}
    html += '<span style="display:inline-flex;align-items:center;gap:4px;padding:3px 9px;border-radius:99px;background:'+bg+';color:'+color+';font-size:11px;font-weight:600;">'+icon+' '+CURRIC_NAMES[cid]+'</span>';
  }});
  html += '</div>';
  if (eng.mismatches.length > 0) {{
    html += '<div style="background:var(--red-subtle);color:var(--red);border-radius:6px;padding:8px 12px;font-size:12px;">&#9888; Courses completed without visiting: ' + eng.mismatches.map(cid=>CURRIC_NAMES[cid]).join(', ') + '</div>';
  }}
  html += '</div>';
  return html;
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
  if (typeof renderCertCharts === 'function' && certTabActive) renderCertCharts();
}}
(function() {{
  const t = localStorage.getItem('pb-theme');
  if (t !== 'dark') {{
    document.body.classList.add('light-mode');
    document.getElementById('btn-theme').innerHTML = '&#9790; Dark';
  }}
}})();

/* ── Info popover ── */
const INFO = {{
  "total-enrolled": "The total number of people currently enrolled in the Accelerate Onboarding program. If filters are applied above, this reflects only the people currently shown.",
  "overdue": "People who have not finished at least one required course whose LMS-assigned due date has already passed. Each course carries its own due date directly from the LMS (Item Required Date) rather than a fixed number of days after hire — a course with no due date yet in the LMS can never count as overdue. Click any row to see exactly which courses are overdue and by how many days.",
  "ontrack": "People who have not missed any course deadlines yet. They may or may not be keeping pace — check the Gap column to see who is falling behind even if they are technically still on time. Click any row to see which courses still need attention.",
  "completed": "People who have finished every required course across all six curricula in the Accelerate program.",
  "first-sale": "Average number of days from hire date to each rep's first Closed Won deal. Only the learner's earliest qualifying deal is counted. Open pipeline is not included. Data pulled from Salesforce.",
  "filters-info": "Market: narrow the table to one market. Status: show only people who are Completed, On Track, or Overdue. Sort: change the default row order. Group: 'Show Early Access Cohort' brings the original June 4 launch cohort back into every stat, chart, and row (hidden by default). Location: show only US or only Canadian learners (Canadian learners are marked with a flag). All filters work together — for example, filter to a market and sort by Most Urgent to quickly find who needs attention there. Use Reset to clear everything.",
  "view-toggle": "Individual view shows one row per learner. By Manager groups learners under their manager so you can see how each team is tracking. The search box updates automatically — in Individual view it searches by learner name, in By Manager view it searches by manager name.",
  "market-chart": "Share of reps in each market broken down by status: green = completed, blue = on track, red = overdue. Each bar totals 100%. Hover a segment to see the percentage and underlying count. A long red segment is a signal that market may need extra support.",
  "curric-chart": "Share of reps in each curriculum broken down by status: green = completed, blue = on track, red = overdue. Each bar totals 100%. Hover a segment to see the percentage and underlying count. A long red segment is where reps are getting stuck.",
  "heatmap": "One row per learner. Click any row to open their full detail — every individual lesson, completion dates, curriculum breakdown, and playbook activity timeline.",
  "col-learner": "The learner's name. A small flag marks Canadian learners so they're easy to spot in a mixed roster. Click any row to open their full detail — every individual lesson, completion dates, curriculum breakdown, and playbook activity history.",
  "col-status": "Each person's current standing in the program. Completed = finished all required courses. Overdue = past the LMS deadline for at least one course. On Track = within all deadlines so far. Click a row to see full detail including which curricula are overdue or behind pace.",
  "col-actual": "Weighted overall completion percentage. Calculated as total lessons completed divided by total lessons assigned across all six curricula. Each lesson counts equally regardless of which curriculum it belongs to.",
  "col-expected": "How far along this person should be today based on each curriculum's own due-date window (e.g. Getting Started/Sales Workflow/Core Portfolio are due by day 7, Prospecting by day 14, Sales Skills weeks 8-21, Pipeline Mgmt weeks 22-35), weighted by how many items each curriculum contains. Compare to Actual % — if Actual is lower, they are behind pace.",
  "col-gap": "Expected % minus Actual %. A positive number means behind schedule by that amount. A negative (or zero) number means on pace or ahead. The color scales with severity: green when at/ahead, amber when a little behind, red when the gap is significant enough to warrant a check-in.",
  "col-curricula": "Six colored squares, one per curriculum in program order: Getting Started, Sales Workflow, Core Portfolio, Prospecting, Sales Skills, Pipeline Management. Green = complete. Blue = in progress. Red = past its deadline and not done. Gray = not started yet. Click the row to see individual lesson detail.",
  "col-days-to-close": "How many days it took this rep to close their first Salesforce deal, measured from their hire date to their earliest Closed Won deal. A dash means no deal has been recorded yet (or no hire date is on file).",
  "export": "Downloads a report based on whoever is currently visible — apply filters first, then export. Full Report: all visible learners with status and progress. Overdue Only: people past a course deadline with their manager's contact info for easy follow-up. Manager Summary: one row per manager with their team's headcount and progress. Example: filter to a market, then export Overdue Only to get a targeted outreach list.",
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
  if (tableView === 'manager') {{
    let showGroup = true;
    document.querySelectorAll('#heatmap-body tr').forEach(row => {{
      if (row.classList.contains('mgr-group-hdr-row')) {{
        showGroup = !q || (row.dataset.manager || '').includes(q);
        row.style.display = showGroup ? '' : 'none';
      }} else {{
        row.style.display = showGroup ? '' : 'none';
      }}
    }});
  }} else {{
    document.querySelectorAll('#heatmap-body tr').forEach(row => {{
      row.style.display = (!q || (row.dataset.name || '').includes(q)) ? '' : 'none';
    }});
  }}
}}

function setTableView(v) {{
  tableView = v;
  document.getElementById('view-individual').classList.toggle('active', v === 'individual');
  document.getElementById('view-manager').classList.toggle('active', v === 'manager');
  const searchEl = document.getElementById('table-search');
  searchEl.placeholder = v === 'manager' ? 'Search manager...' : 'Search name...';
  searchEl.value = '';
  renderTable();
}}

function toggleMgrGroup(el) {{
  el.classList.toggle('open');
  const isOpen = el.classList.contains('open');
  let row = el.closest('tr').nextElementSibling;
  while (row && !row.classList.contains('mgr-group-hdr-row')) {{
    row.style.display = isOpen ? '' : 'none';
    row = row.nextElementSibling;
  }}
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


/* ── Filters ── */
function resetFilters() {{
  document.getElementById('f-market').value = '';
  document.getElementById('f-status').value = '';
  document.getElementById('f-sort').value = 'name';
  document.getElementById('table-search').value = '';
  testGroupFilter = 'hide';
  document.getElementById('test-group-toggle').textContent = 'Show Early Access Cohort';
  locationFilter = 'all';
  document.querySelectorAll('#location-btns button').forEach((b,i) => b.classList.toggle('active', i===0));
  filterTableRows();
  if (typeof resetTenure === 'function') resetTenure();
  if (typeof resetCertControls === 'function') resetCertControls();
  applyFilters();
}}

function applyFilters() {{
  populateMarketDropdown();
  const mkt    = document.getElementById('f-market').value;
  const status = document.getElementById('f-status').value;
  const sort   = document.getElementById('f-sort').value;

  filtered = PEOPLE.filter(p => {{
    if (hideTLG && TLG_SET.has(p.name.toLowerCase())) return false;
    if (testGroupFilter === 'hide' && TEST_GROUP_NAMES.has(p.name)) return false;
    if (testGroupFilter === 'only' && !TEST_GROUP_NAMES.has(p.name)) return false;
    if (locationFilter === 'us' && p.isCanada) return false;
    if (locationFilter === 'ca' && !p.isCanada) return false;
    if (!passesTenure(p)) return false;
    if (mkt && p.market !== mkt) return false;
    const st = computeStatus(p);
    if (status && st !== status) return false;
    return true;
  }});

  if (sort === 'pct-desc') {{ filtered.sort((a,b) => b.overallPct - a.overallPct); tableSort = {{col:'overall', dir:'desc'}}; }}
  else if (sort === 'pct-asc') {{ filtered.sort((a,b) => a.overallPct - b.overallPct); tableSort = {{col:'overall', dir:'asc'}}; }}
  else if (sort === 'days-asc') {{
    const statusRank = {{ 'Overdue': 0, 'On Track': 1, 'Completed': 2, 'Unknown': 3 }};
    filtered.sort((a, b) => {{
      const ra = statusRank[computeStatus(a)] ?? 3, rb = statusRank[computeStatus(b)] ?? 3;
      if (ra !== rb) return ra - rb;
      const da = daysLeft(a) ?? 999, db = daysLeft(b) ?? 999;
      return da - db;
    }});
    tableSort = {{col:'days', dir:'asc'}};
  }}
  else {{ filtered.sort((a,b) => a.name.localeCompare(b.name)); tableSort = {{col:'name', dir:'asc'}}; }}

  certBase = PEOPLE.filter(p => passesBaseFilters(p, mkt));
  renderStats();
  renderTable();
  renderCharts();
  document.getElementById('result-count').textContent = filtered.length + ' shown';
  renderCert();
}}

/* ── Stats ── */
function renderStats() {{
  const total = filtered.length;
  const overdue = filtered.filter(p => computeStatus(p) === 'Overdue').length;
  const ontrack = filtered.filter(p => computeStatus(p) === 'On Track').length;
  const completed = filtered.filter(p => computeStatus(p) === 'Completed').length;

  const pct = (n) => total ? Math.round((n / total) * 100) + '%' : '—';

  document.getElementById('s-total').textContent = total;

  document.getElementById('s-overdue').textContent = pct(overdue);
  document.getElementById('s-overdue-sub').textContent = overdue
    ? overdue + ' of ' + total + ' \\u2014 past a curriculum deadline'
    : '0 of ' + total;

  document.getElementById('s-ontrack').textContent = pct(ontrack);
  document.getElementById('s-ontrack-sub').textContent = ontrack + ' of ' + total;

  document.getElementById('s-completed').textContent = pct(completed);
  document.getElementById('s-completed-sub').textContent = completed + ' of ' + total;
  const salesVals = filtered.map(p => SALES_MAP[p.email]).filter(v => v && v.daysToClose != null);
  const avgDays = salesVals.length ? Math.round(salesVals.reduce((s,v) => s + v.daysToClose, 0) / salesVals.length) : null;
  document.getElementById('s-sales').textContent = avgDays !== null ? avgDays + 'd' : '—';
}}

/* ── Heatmap table ── */
function renderTable() {{
  // Header
  const thead = document.getElementById('heatmap-head');
  function thS(label, col, extraCls, infoKey) {{
    const active = tableSort.col === col;
    const arrow = active ? (tableSort.dir === 'asc' ? ' ↑' : ' ↓') : '';
    const cls = ['sortable-th', extraCls, active ? 'sort-active' : ''].filter(Boolean).join(' ');
    const btn = infoKey ? '<span class="info-btn print-hide" onclick="event.stopPropagation();showInfo(event,\\'' + infoKey + '\\')">?</span>' : '';
    return '<th class="' + cls + '" data-col="' + col + '" onclick="sortByCol(this.dataset.col)">' + label + arrow + btn + '</th>';
  }}
  let hRow = '<tr>' + thS('Learner','name','','col-learner');
  hRow += thS('Actual %','overall','overall-col','col-actual');
  hRow += thS('Expected %','expected','overall-col','col-expected');
  hRow += thS('Gap','gap','overall-col','col-gap');
  hRow += thS('Progress','status','','col-status');
  hRow += '<th style="font-size:10px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;padding:10px 12px;text-align:center;background:var(--surface2);border-bottom:1px solid var(--border);white-space:nowrap;">Curricula<span class="info-btn print-hide" onclick="showInfo(event,\\'col-curricula\\')">?</span></th>';
  hRow += thS('Days to Close','daysToClose','overall-col','col-days-to-close');
  hRow += '</tr>';
  thead.innerHTML = hRow;

  // Body
  const tbody = document.getElementById('heatmap-body');
  if (!filtered.length) {{
    tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;color:var(--muted);padding:32px;">No learners match the current filters.</td></tr>';
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
      else if (col === 'overall')   {{ va = a.overallPct; vb = b.overallPct; }}
      else if (col === 'expected')  {{ va = expectedPct(a); vb = expectedPct(b); }}
      else if (col === 'gap')       {{ va = gapPct(a); vb = gapPct(b); }}
      else if (col === 'daysToClose') {{ va = SALES_MAP[a.email] ? (SALES_MAP[a.email].daysToClose ?? 999) : 999; vb = SALES_MAP[b.email] ? (SALES_MAP[b.email].daysToClose ?? 999) : 999; }}
      else {{ va = a.curricula[col] ? a.curricula[col].pct : 0; vb = b.curricula[col] ? b.curricula[col].pct : 0; }}
      if (va < vb) return tableSort.dir === 'asc' ? -1 : 1;
      if (va > vb) return tableSort.dir === 'asc' ? 1 : -1;
      return 0;
    }});
  }}
  function worstCurric(p) {{
    let worst = null, worstDl = null;
    CURRIC_IDS.forEach(cid => {{
      const c = p.curricula[cid];
      const dl = curricDaysLeft(c);
      if (dl === null) return;
      if (worstDl === null || dl < worstDl) {{ worstDl = dl; worst = cid; }}
    }});
    return worst ? {{cid: worst, name: CURRIC_NAMES[worst], dl: worstDl}} : null;
  }}


  function biggestGapCurric(p) {{
    let worst = null, worstPct = null;
    CURRIC_IDS.forEach(cid => {{
      const c = p.curricula[cid];
      if (!c || c.complete) return;
      if (worstPct === null || c.pct < worstPct) {{ worstPct = c.pct; worst = cid; }}
    }});
    return worst ? {{cid: worst, name: CURRIC_NAMES[worst], pct: worstPct}} : null;
  }}

  function personRow(p) {{
    const status = computeStatus(p);
    const statusClass = status === 'Completed' ? 'sb-completed' : status === 'On Track' ? 'sb-ontrack' : 'sb-overdue';
    const info = interventionInfo(p);
    let progressExtra = '';
    let progressTitle = '';
    if (status === 'Overdue' && info.days > 0) {{
      const dayColor = info.level === 'red' ? 'var(--red)' : 'var(--amber)';
      progressExtra = ' <span style="font-weight:800;color:' + dayColor + ';">&middot; ' + info.days + 'd</span>';
      progressTitle = ' title="' + escHtml(info.curricName + ': ' + info.days + 'd overdue') + '"';
    }}
    let dotsCell = '<td style="text-align:center;padding:5px 12px;"><div style="display:inline-flex;gap:3px;align-items:center;">';
    CURRIC_IDS.forEach(cid => {{
      const c = p.curricula[cid];
      const bg = curricDotBg(c, p);
      const cdl = curricDaysLeft(c);
      const statusTip = !c ? 'N/A' : c.complete ? 'Completed' : cdl === null ? 'On Track (no deadline yet)' :
        cdl <= 0 ? Math.abs(cdl) + 'd overdue' : cdl + 'd left';
      const lbl = CURRIC_NAMES[cid] + ': ' + (c ? c.pct + '%' : 'N/A') + ' · ' + statusTip;
      dotsCell += '<span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:' + bg + ';" title="' + lbl + '"></span>';
    }});
    dotsCell += '</div></td>';
    const eng = pbEngagement(p);
    const pSaleRow = SALES_MAP[p.email];
    const saleCell = pSaleRow
      ? '<td class="pct-cell" style="font-weight:700;font-size:11px;color:var(--accent);white-space:nowrap;">' + (pSaleRow.daysToClose != null ? pSaleRow.daysToClose + 'd' : '—') + '</td>'
      : '<td class="pct-cell" style="text-align:center;color:var(--muted);opacity:.4;">&#8212;</td>';
    return '<tr data-email="' + escHtml(p.email) + '" data-name="' + escHtml(p.name.toLowerCase()) + '" onclick="openModal(this.dataset.email)" title="Click to see full detail">' +
      '<td class="name-cell">' + escHtml(p.name) + (p.isCanada ? ' <span title="Canada" style="font-size:0.85em;">&#127464;&#127462;</span>' : '') + '</td>' +
      '<td class="pct-cell" style="font-weight:600;font-size:12px;">' + p.overallPct + '%</td>' +
      '<td class="pct-cell" style="font-weight:600;font-size:12px;color:var(--muted);">' + expectedPct(p) + '%</td>' +
      (function(){{ const g=gapPct(p); return '<td class="pct-cell"><span class="pct-pill" style="' + gapStyle(g) + '">' + g + '%</span></td>'; }})() +
      '<td><span class="status-badge ' + statusClass + '"' + progressTitle + '>' + status + progressExtra + '</span></td>' +
      dotsCell +
      saleCell +
    '</tr>';
  }}

  if (tableView === 'manager') {{
    const groups = {{}};
    filtered.forEach(p => {{
      const mgr = p.manager || '(No Manager)';
      if (!groups[mgr]) groups[mgr] = [];
      groups[mgr].push(p);
    }});
    const colCount = 7;
    let html = '';
    Object.keys(groups).sort().forEach(mgr => {{
      const team = groups[mgr];
      const avg = Math.round(team.reduce((s, p) => s + p.overallPct, 0) / team.length);
      const od = team.filter(p => computeStatus(p) === 'Overdue').length;
      const odStr = od > 0 ? ' &nbsp;·&nbsp; <span style="color:var(--red);font-weight:600">' + od + ' overdue</span>' : '';
      html += '<tr class="mgr-group-hdr-row" data-manager="' + escHtml(mgr.toLowerCase()) + '"><td colspan="' + colCount + '">' +
        '<div class="mgr-group-hdr open" onclick="toggleMgrGroup(this)">' +
        '<span style="font-weight:700;font-size:13px;">' + escHtml(mgr) + '</span>' +
        '<span style="font-size:11px;color:var(--muted);">' + team.length + ' learner' + (team.length !== 1 ? 's' : '') + ' &nbsp;·&nbsp; avg ' + avg + '%' + odStr + ' &nbsp;<span class="mgr-chevron">&#9660;</span></span>' +
        '</div></td></tr>';
      team.forEach(p => {{ html += personRow(p); }});
    }});
    tbody.innerHTML = html;
  }} else {{
    tbody.innerHTML = display.map(personRow).join('');
  }}

  document.getElementById('heatmap-foot').innerHTML = '';
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
  const marketStats = markets.map(m => {{
    const mp = filtered.filter(p => p.market === m);
    const total    = mp.length;
    const complete = mp.filter(p => p.overallDone).length;
    const overdue  = mp.filter(p => computeStatus(p) === 'Overdue').length;
    const onTrack  = total - complete - overdue;
    const pct = n => total ? Math.round((n/total)*100) : 0;
    return {{total, complete, onTrack, overdue, completePct: pct(complete), onTrackPct: pct(onTrack), overduePct: pct(overdue)}};
  }});
  // Grow the chart's height with the number of markets so every label gets room to render
  const wrap = document.getElementById('marketChart').closest('.chart-wrap');
  wrap.style.height = Math.max(220, markets.length * 32) + 'px';
  const ctx = document.getElementById('marketChart').getContext('2d');
  if (marketChartObj) marketChartObj.destroy();
  const statusColors = {{ complete: cv('green'), onTrack: cv('accent'), overdue: cv('red') }};
  marketChartObj = new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels: markets,
      datasets: [
        {{
          label: 'Completed',
          data: marketStats.map(s => s.completePct),
          backgroundColor: statusColors.complete + 'cc',
          borderColor: statusColors.complete,
          borderWidth: 1,
        }},
        {{
          label: 'On Track',
          data: marketStats.map(s => s.onTrackPct),
          backgroundColor: statusColors.onTrack + 'cc',
          borderColor: statusColors.onTrack,
          borderWidth: 1,
        }},
        {{
          label: 'Overdue',
          data: marketStats.map(s => s.overduePct),
          backgroundColor: statusColors.overdue + 'cc',
          borderColor: statusColors.overdue,
          borderWidth: 1,
        }},
      ]
    }},
    options: {{
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend: {{
          display: true,
          position: 'bottom',
          labels: {{color: cv('muted'), font:{{size:11}}, boxWidth:12, padding:12}}
        }},
        tooltip: {{
          callbacks: {{
            label: c => {{
              const s = marketStats[c.dataIndex];
              const raw = c.dataset.label === 'Completed' ? s.complete : c.dataset.label === 'On Track' ? s.onTrack : s.overdue;
              return '  ' + c.dataset.label + ': ' + c.parsed.x + '% (' + raw + ')';
            }},
            afterBody: items => {{
              const s = marketStats[items[0].dataIndex];
              return ['  Total in this market: ' + s.total];
            }}
          }}
        }}
      }},
      scales: {{
        x: {{
          stacked: true,
          min:0, max:100,
          grid: {{color: 'rgba(255,255,255,0.06)'}},
          ticks: {{color: cv('muted'), font:{{size:11}}, callback: v => v+'%'}},
          border: {{color:'transparent'}},
        }},
        y: {{
          stacked: true,
          grid: {{display:false}},
          ticks: {{color: cv('muted'), font:{{size:11}}, autoSkip:false}},
          border: {{color:'transparent'}},
        }}
      }}
    }}
  }});
}}

function renderCurricChart() {{
  const labels = CURRIC_IDS.map(cid => CURRIC_NAMES[cid]);
  const curricStats = CURRIC_IDS.map(cid => {{
    const people = filtered.map(p => p.curricula[cid]).filter(Boolean);
    let complete = 0, onTrack = 0, overdue = 0;
    people.forEach(c => {{
      const st = curricTrackStatus(c);
      if (st === 'Completed') complete++;
      else if (st === 'Overdue') overdue++;
      else onTrack++;
    }});
    return {{total: people.length, complete, onTrack, overdue}};
  }});
  const curricStatsPct = curricStats.map(s => {{
    const pct = n => s.total ? Math.round((n/s.total)*100) : 0;
    return {{...s, completePct: pct(s.complete), onTrackPct: pct(s.onTrack), overduePct: pct(s.overdue)}};
  }});
  const ctx = document.getElementById('curricChart').getContext('2d');
  if (curricChartObj) curricChartObj.destroy();
  const statusColors = {{ complete: cv('green'), onTrack: cv('accent'), overdue: cv('red') }};
  curricChartObj = new Chart(ctx, {{
    type: 'bar',
    data: {{
      labels,
      datasets: [
        {{
          label: 'Completed',
          data: curricStatsPct.map(s => s.completePct),
          backgroundColor: statusColors.complete + 'cc',
          borderColor: statusColors.complete,
          borderWidth: 1,
        }},
        {{
          label: 'On Track',
          data: curricStatsPct.map(s => s.onTrackPct),
          backgroundColor: statusColors.onTrack + 'cc',
          borderColor: statusColors.onTrack,
          borderWidth: 1,
        }},
        {{
          label: 'Overdue',
          data: curricStatsPct.map(s => s.overduePct),
          backgroundColor: statusColors.overdue + 'cc',
          borderColor: statusColors.overdue,
          borderWidth: 1,
        }},
      ]
    }},
    options: {{
      responsive: true,
      maintainAspectRatio: false,
      plugins: {{
        legend: {{
          display: true,
          position: 'bottom',
          labels: {{color: cv('muted'), font:{{size:11}}, boxWidth:12, padding:12}}
        }},
        tooltip: {{
          callbacks: {{
            label: c => {{
              const s = curricStatsPct[c.dataIndex];
              const raw = c.dataset.label === 'Completed' ? s.complete : c.dataset.label === 'On Track' ? s.onTrack : s.overdue;
              return '  ' + c.dataset.label + ': ' + c.parsed.y + '% (' + raw + ')';
            }},
            afterBody: items => {{
              const s = curricStatsPct[items[0].dataIndex];
              return ['  Total enrolled: ' + s.total];
            }}
          }}
        }}
      }},
      scales: {{
        y: {{
          stacked: true,
          min:0, max:100,
          grid: {{color: 'rgba(255,255,255,0.06)'}},
          ticks: {{color: cv('muted'), font:{{size:11}}, callback: v => v+'%'}},
          border: {{color:'transparent'}},
        }},
        x: {{
          stacked: true,
          grid: {{display:false}},
          ticks: {{color: cv('muted'), font:{{size:10}}, maxRotation:30}},
          border: {{color:'transparent'}},
        }}
      }}
    }}
  }});
}}

/* ── Report-card helpers ── */
function scheduleCheckHtml(p) {{
  const d = daysElapsed(p);
  let rows = '';
  CURRIC_IDS.forEach(cid => {{
    const c = p.curricula[cid];
    if (!c) return;
    const due = curricDueDate(p, cid);
    const dueLabel = due ? fmtDate(due.toISOString().slice(0,10)) : '—';
    const should = curricExpectedPct(cid, d);
    const actual = c.pct;
    const color = c.complete ? 'var(--green)' : actual >= should ? 'var(--green)' : (should - actual) > 25 ? 'var(--red)' : 'var(--amber)';
    rows += '<tr>' +
      '<td>' + escHtml(c.title) + '</td>' +
      '<td class="c" style="color:var(--muted);white-space:nowrap;">' + dueLabel + '</td>' +
      '<td class="c">' + should + '%</td>' +
      '<td class="c sched-actual" style="color:' + color + '">' + actual + '%</td>' +
    '</tr>';
  }});
  return '<div style="margin-bottom:6px;">' +
    '<div style="font-size:12px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px;">Progress by Curriculum</div>' +
    '<table class="sched-table"><thead><tr>' +
      '<th>Curriculum</th><th class="c">Due By</th><th class="c">Should Be At</th><th class="c">Actually At</th>' +
    '</tr></thead><tbody>' + rows + '</tbody></table>' +
  '</div>';
}}
function curriculaCompletedCount(p) {{
  const ids = Object.keys(p.curricula);
  const done = ids.filter(cid => p.curricula[cid].complete).length;
  return {{done, total: ids.length}};
}}
function overdueItemsCount(p) {{
  let n = 0;
  Object.values(p.curricula).forEach(c => {{
    c.items.forEach(it => {{
      if (!it.done && it.req) {{
        const dl = itemDaysLeft(it);
        if (dl !== null && dl <= 0) n++;
      }}
    }});
  }});
  return n;
}}
function cohortAvgPct() {{
  if (!PEOPLE.length) return 0;
  return Math.round(PEOPLE.reduce((s,p) => s + p.overallPct, 0) / PEOPLE.length);
}}
function projectedCompletionDate(p) {{
  if (p.overallDone) return null;
  if (!p.assignDate) return null;
  const elapsed = daysElapsed(p);
  if (elapsed <= 0 || p.overallPct <= 0) return null;
  const ratePerDay = p.overallPct / elapsed;
  if (ratePerDay <= 0) return null;
  const daysToFinish = Math.ceil((100 - p.overallPct) / ratePerDay);
  const proj = new Date(p.assignDate);
  proj.setDate(proj.getDate() + elapsed + daysToFinish);
  return proj;
}}
function riskInfo(p, status) {{
  const mgrName = p.manager || 'their manager';
  const firstName = p.first || p.name.split(' ')[0];
  if (status === 'Completed') {{
    return {{level:'green', label:'Program Complete', action:'No action needed — worth recognizing the completion with ' + mgrName + '.'}};
  }}
  const g = gapPct(p);
  const od = overdueCount(p);
  if (status === 'Overdue' && (g > 40 || od >= 2)) {{
    return {{level:'red', label:'High Risk — Coaching Needed', action:'Schedule a 1:1 with ' + firstName + ' and ' + mgrName + ' this week — ' + od + ' curricul' + (od===1?'um is':'a are') + ' past deadline.'}};
  }}
  if (status === 'Overdue') {{
    return {{level:'amber', label:'Overdue — Check In', action:'Send a check-in reminder; ' + od + ' curricul' + (od===1?'um':'a') + ' past its deadline.'}};
  }}
  return {{level:'green', label:'On Track', action:'No action needed — progressing at or ahead of the pacing schedule.'}};
}}

/* ── User detail modal ── */
function openModal(email) {{
  const p = PEOPLE.find(x => x.email === email);
  if (!p) return;
  const status = computeStatus(p);
  const dl = daysLeft(p);
  const elapsed = daysElapsed(p);
  const pct45 = Math.min(100, Math.round(elapsed / PROGRAM_DAYS * 100));
  const barColor = p.overallDone ? 'var(--green)' : dl <= 0 ? 'var(--red)' : 'var(--accent)';

  let daysLabel = '';
  if (p.overallDone) daysLabel = 'Program complete';
  else if (dl === null) daysLabel = 'No assignment date on file';
  else if (dl < 0) daysLabel = Math.abs(dl) + ' days past deadline';
  else if (dl === 0) daysLabel = 'Due today';
  else daysLabel = dl + ' days remaining';

  const statusBadge = '<span class="status-badge ' +
    (status==='Completed'?'sb-completed':status==='On Track'?'sb-ontrack':'sb-overdue') +
    '">' + status + '</span>';

  // Curriculum sections
  let curricHtml = '';
  CURRIC_IDS.forEach(cid => {{
    const c = p.curricula[cid];
    if (!c) return;
    const dl = curricDaysLeft(c);
    const drBg = dl !== null && dl <= 0 ? 'var(--red-subtle)' : 'var(--green-subtle)';
    const drColor = dl !== null && dl <= 0 ? 'var(--red)' : 'var(--green)';
    const drLabel = dl === null ? '' :
      '<span class="days-badge" style="background:' + drBg + ';color:' + drColor + '">' +
      (dl < 0 ? Math.abs(dl) + 'd overdue' : dl === 0 ? 'Due today' : dl + 'd left') + '</span>';
    const doneBadge = c.complete ?
      '<span class="status-badge sb-completed" style="font-size:10px">Done</span>' : '';

    let missingLabel = '';
    if (!c.complete) {{
      const missingCourses   = c.items.filter(i => !i.done && i.type !== 'VILT').length;
      const missingWorkshops = c.items.filter(i => !i.done && i.type === 'VILT').length;
      if (missingCourses > 0 || missingWorkshops > 0) {{
        const parts = [];
        if (missingCourses   > 0) parts.push(missingCourses   + ' ' + (missingCourses   === 1 ? 'course'   : 'courses'));
        if (missingWorkshops > 0) parts.push(missingWorkshops + ' ' + (missingWorkshops === 1 ? 'workshop' : 'workshops'));
        missingLabel = '<span style="font-size:11px;color:var(--muted);white-space:nowrap;">Missing: ' + parts.join(', ') + '</span>';
      }}
    }}

    const itemsHtml = c.items.map(it => {{
      const checkClass = it.done ? 'done' : 'not';
      const checkMark  = it.done ? '&#10003;' : '&nbsp;';
      const typeLabel  = it.type === 'VILT' ?
        '<span class="item-type" style="background:#a855f711;color:#a855f7;">Workshop</span>' :
        '<span class="item-type">Online</span>';
      const registeredLabel = (!it.done && it.type === 'VILT') ?
        (it.registered
          ? '<span class="item-type" style="background:var(--green-subtle);color:var(--green);">Registered</span>'
          : '<span class="item-type" style="background:var(--red-subtle);color:var(--red);">Not Registered</span>')
        : '';
      const dateStr = it.done && it.date ?
        '<span class="item-date">' + fmtDate(it.date) + '</span>' : '';
      return '<div class="item-row">' +
        '<div class="item-check ' + checkClass + '">' + checkMark + '</div>' +
        '<div class="item-title">' + escHtml(it.title) + '</div>' +
        typeLabel + registeredLabel + dateStr + '</div>';
    }}).join('');

    curricHtml += '<div class="curric-section">' +
      '<div class="curric-header" onclick="toggleCurric(this)">' +
        '<span class="curric-title">' + escHtml(c.title) + '</span>' +
        '<span class="pct-pill" style="' + curricPillStyle(c, null) + ';font-size:11px">' + c.pct + '%</span>' +
        drLabel + doneBadge + missingLabel +
        '<span class="curric-chevron">&#9660;</span>' +
      '</div>' +
      '<div class="curric-items">' + itemsHtml + '</div>' +
    '</div>';
  }});

  // Activity timeline — deduped playbook visits + LMS completions, sorted by date
  const tlEvents = [];
  const tlSeen = new Set();
  (p.playbook || []).filter(v => !p.assignDate || v.date >= p.assignDate).forEach(v => {{
    const k = (v.page || '') + '|' + (v.date || '');
    const pbSlug = urlSlug(v.url);
    if (!tlSeen.has(k)) {{ tlSeen.add(k); tlEvents.push({{t:'pb', date:v.date||'', page:v.page||'', slug:pbSlug, curricId:PLAYBOOK_CURRIC[pbSlug]||null}}); }}
  }});
  CURRIC_IDS.forEach(cid => {{
    const c = p.curricula[cid];
    if (!c) return;
    c.items.forEach(it => {{
      if (it.done && it.date && (!p.assignDate || it.date >= p.assignDate)) tlEvents.push({{t:'lms', date:it.date, title:it.title, curric:c.title, curricId:cid}});
    }});
  }});
  const _pgIdx = slug => {{ const i = PAGE_ORDER.indexOf(slug); return i >= 0 ? i : 999; }};
  tlEvents.sort((a,b) => {{
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    const aSlug = a.t === 'pb' ? a.slug : (COURSE_PAGE[a.title] || '');
    const bSlug = b.t === 'pb' ? b.slug : (COURSE_PAGE[b.title] || '');
    const ai = _pgIdx(aSlug), bi = _pgIdx(bSlug);
    if (ai !== bi) return ai - bi;
    return a.t === 'pb' && b.t !== 'pb' ? -1 : a.t !== 'pb' && b.t === 'pb' ? 1 : 0;
  }});

  // Build filtered playbook visits (on/after assignDate) and per-curriculum date lookup
  const filteredPlaybook = (p.playbook || []).filter(v => !p.assignDate || v.date >= p.assignDate);
  const pbVisitDates = {{}};
  filteredPlaybook.forEach(v => {{
    const cid = PLAYBOOK_CURRIC[urlSlug(v.url)];
    if (cid) {{ if (!pbVisitDates[cid]) pbVisitDates[cid] = new Set(); pbVisitDates[cid].add(v.date); }}
  }});
  function dayBefore(d) {{
    const dt = new Date(d + 'T00:00:00'); dt.setDate(dt.getDate() - 1);
    return dt.toISOString().slice(0, 10);
  }}
  // Path score: of all individual completions, what fraction had a playbook visit same day or day before
  let ledCount = 0, totalCount = 0;
  CURRIC_IDS.forEach(cid => {{
    const c = p.curricula[cid];
    if (!c) return;
    const visits = pbVisitDates[cid];
    c.items.filter(i => i.done && i.date && (!p.assignDate || i.date >= p.assignDate)).forEach(i => {{
      totalCount++;
      if (visits && (visits.has(i.date) || visits.has(dayBefore(i.date)))) ledCount++;
    }});
  }});
  let pathLabel = '', pathColor = '';
  if (totalCount > 0) {{
    const ratio = ledCount / totalCount;
    const pct = Math.round(ratio * 100) + '%';
    if (ratio >= 0.6)      {{ pathLabel = 'Playbook-led ' + pct;      pathColor = 'var(--green)'; }}
    else if (ratio >= 0.3) {{ pathLabel = 'Mixed ' + pct;              pathColor = 'var(--amber)'; }}
    else                   {{ pathLabel = 'Skipping playbook ' + pct;  pathColor = 'var(--red)'; }}
  }}

  // Sequence check: first completion dates should be non-decreasing across CURRIC_IDS order
  const curricSeq = CURRIC_IDS.map(cid => {{
    const c = p.curricula[cid];
    if (!c) return null;
    const first = c.items.filter(i => i.done && i.date).map(i => i.date).sort()[0];
    return first ? {{cid, first}} : null;
  }}).filter(Boolean);
  let seqLabel = '', seqColor = '';
  if (curricSeq.length >= 2) {{
    const inSeq = curricSeq.every((x, i) => i === 0 || x.first >= curricSeq[i-1].first);
    seqLabel = inSeq ? 'In sequence' : 'Out of sequence';
    seqColor = inSeq ? 'var(--green)' : 'var(--amber)';
  }}

  const tlHtml = tlEvents.length === 0
    ? '<div style="font-size:12px;color:var(--muted);padding:12px 14px;">No activity recorded yet.</div>'
    : tlEvents.map(ev => ev.t === 'pb'
        ? '<div class="item-row"><div style="width:10px;height:10px;border-radius:50%;background:var(--accent);flex-shrink:0;"></div><div style="flex:1;font-size:12px;"><span style="color:var(--accent);font-weight:500;">Playbook</span> &mdash; ' + escHtml(ev.page) + '</div><div class="item-date">' + fmtDate(ev.date) + '</div></div>'
        : '<div class="item-row"><div style="width:10px;height:10px;border-radius:50%;background:var(--green);flex-shrink:0;"></div><div style="flex:1;font-size:12px;"><span style="color:var(--green);font-weight:500;">Completed</span> &mdash; ' + escHtml(ev.title) + ' <span style="color:var(--muted);font-size:11px;">(' + escHtml(ev.curric) + ')</span></div><div class="item-date">' + fmtDate(ev.date) + '</div></div>'
      ).join('');

  const pSale = SALES_MAP[email];
  const saleHtml = pSale
    ? '<span>&middot; First Sale: <strong style="color:var(--accent)">' + (pSale.daysToClose != null ? pSale.daysToClose + 'd' : '') + '</strong> &middot; $' + Math.round(pSale.amount).toLocaleString('en-US') + ' &middot; ' + escHtml(pSale.accountName) + ' &middot; ' + fmtDate(pSale.closeDate) + '</span>'
    : '<span style="color:var(--muted);">&middot; First Sale: None yet</span>';
  const risk = riskInfo(p, status);
  const riskIcon = risk.level === 'red' ? '&#9888;&#65039; ' : risk.level === 'amber' ? '&#9679; ' : '&#10003; ';
  const curricCounts = curriculaCompletedCount(p);
  const overdueItems = overdueItemsCount(p);
  const cohortAvg = cohortAvgPct();
  const vsCohort = p.overallPct - cohortAvg;
  const projDate = projectedCompletionDate(p);
  const projLabel = p.overallDone ? 'Complete' : (projDate ? fmtDate(projDate.toISOString().slice(0,10)) : '—');
  const projSub = p.overallDone ? '' : (projDate ? 'at current pace' : 'not enough data yet');
  const riskBannerHtml =
    '<div class="rc-riskbar risk-' + risk.level + '">' +
      '<span class="rc-risklabel">' + riskIcon + risk.label + '</span>' +
      '<span class="rc-riskaction">' + risk.action + '</span>' +
    '</div>' +
    '<div class="metric-grid">' +
      '<div class="metric-box"><div class="mb-label">Overall Progress</div><div class="mb-value">' + p.overallPct + '%</div>' +
        '<div class="mb-sub">' + (vsCohort >= 0 ? vsCohort + '% above' : Math.abs(vsCohort) + '% below') + ' cohort avg (' + cohortAvg + '%)</div></div>' +
      '<div class="metric-box"><div class="mb-label">Pacing Target</div><div class="mb-value">' + expectedPct(p) + '%</div>' +
        '<div class="mb-sub">day ' + daysElapsed(p) + ' of ' + PROGRAM_DAYS + '</div></div>' +
      '<div class="metric-box"><div class="mb-label">Curricula Done</div><div class="mb-value">' + curricCounts.done + ' / ' + curricCounts.total + '</div>' +
        '<div class="mb-sub">' + overdueItems + ' item' + (overdueItems !== 1 ? 's' : '') + ' overdue</div></div>' +
      '<div class="metric-box"><div class="mb-label">Est. Completion</div><div class="mb-value">' + projLabel + '</div>' +
        '<div class="mb-sub">' + projSub + '</div></div>' +
    '</div>';
  const content = `
    <div class="modal-header" style="padding-right:52px;">
      <div style="position:absolute;top:17px;right:54px;">${{statusBadge}}</div>
      <div class="modal-name">${{escHtml(p.name)}}${{p.jobTitle ? ' <span style="font-size:14px;font-weight:400;color:var(--muted);">&middot; ' + escHtml(p.jobTitle) + '</span>' : ''}}</div>
      <div class="modal-meta" style="margin-top:6px;gap:8px;flex-wrap:wrap;">
        ${{p.email ? '<a href="mailto:' + escHtml(p.email) + '" style="color:var(--accent)">' + escHtml(p.email) + '</a>' : ''}}
        ${{p.market ? '<span>&middot; ' + escHtml(p.market) + '</span>' : ''}}
        ${{p.hireDate ? '<span>&middot; Hired ' + fmtDate(p.hireDate) + '</span>' : ''}}
        ${{p.assignDate ? '<span>&middot; Enrolled ' + fmtDate(p.assignDate) + '</span>' : ''}}
        ${{saleHtml}}
      </div>
      ${{p.manager ? '<div class="modal-meta" style="margin-top:3px;gap:6px;"><span style="color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.4px;font-weight:600;">Manager</span><span>' + escHtml(p.manager) + '</span><span>&middot; ' + escHtml(p.mgrTitle) + '</span><span>&middot; <a href="mailto:' + escHtml(p.mgrEmail) + '" style="color:var(--accent)">' + escHtml(p.mgrEmail) + '</a></span></div>' : ''}}
    </div>
    <div class="modal-body">
      ${{riskBannerHtml}}
      <div class="progress-section">
        <div class="progress-label">
          <span>35-Day LMS Window</span>
          <span style="font-size:11px;color:var(--muted)">${{daysLabel}}</span>
        </div>
        <div class="progress-bar-wrap">
          <div class="progress-bar-fill" style="width:${{pct45}}%;background:${{barColor}}"></div>
        </div>
        <div class="progress-sublabel" style="display:flex;gap:16px;flex-wrap:wrap;align-items:center;">
          <span>Actual: <strong>${{p.overallPct}}%</strong></span>
          <span style="color:var(--muted)">Expected: <strong style="color:var(--text)">${{expectedPct(p)}}%</strong></span>
          <span>Gap: <strong style="color:${{gapPct(p) > 40 ? 'var(--red)' : gapPct(p) > 0 ? 'var(--amber)' : 'var(--green)'}}">${{gapPct(p)}}%</strong></span>
        </div>
      </div>

      ${{scheduleCheckHtml(p)}}

      ${{pbSummaryHtml(p)}}

      <div style="font-size:12px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px;">Curriculum Progress</div>
      <div style="display:flex;gap:12px;flex-wrap:wrap;font-size:11px;color:var(--muted);margin-bottom:10px;align-items:center;">
        <span style="display:inline-flex;align-items:center;gap:4px;"><span style="width:8px;height:8px;background:#15803d;border-radius:2px;display:inline-block;"></span>Done</span>
        <span style="display:inline-flex;align-items:center;gap:4px;"><span style="width:8px;height:8px;background:#1d4ed8;border-radius:2px;display:inline-block;"></span>In Progress</span>
        <span style="display:inline-flex;align-items:center;gap:4px;"><span style="width:8px;height:8px;background:#b91c1c;border-radius:2px;display:inline-block;"></span>Past Due</span>
        <span style="display:inline-flex;align-items:center;gap:4px;"><span style="width:8px;height:8px;background:#6b7280;border-radius:2px;display:inline-block;"></span>Not Started</span>
      </div>
      ${{curricHtml}}

      <div class="curric-section" style="margin-top:20px;margin-bottom:0;">
        <div class="curric-header" onclick="toggleCurric(this)">
          <span class="curric-title">Activity Timeline</span>
          <span style="font-size:11px;color:var(--muted);">${{tlEvents.length}} event${{tlEvents.length !== 1 ? 's' : ''}}</span>
          ${{pathLabel ? '<span style="font-size:11px;font-weight:700;color:' + pathColor + ';">' + pathLabel + '</span>' : ''}}
          ${{seqLabel ? '<span style="font-size:11px;font-weight:700;color:' + seqColor + ';">' + seqLabel + '</span>' : ''}}
          <span class="curric-chevron">&#9660;</span>
        </div>
        <div class="curric-items">${{tlHtml}}</div>
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

/* ═══════════════ Accelerate Certification Journey tab ═══════════════ */
let certBase = [];
let certTabActive = false;
let certTableView = 'individual';   // 'individual' | 'manager' — same pattern as the Onboarding tab's tableView
let certMarketCertChartObj = null, certMarketAttChartObj = null;

/* The four requirements, in journey order. */
const CERT_STEPS = [
  {{key:'onboarding',   label:'Onboarding',     full:'Accelerate Onboarding curriculum (100%)'}},
  {{key:'bootCamp',     label:'Boot Camp',      full:'Accelerate Advanced Boot Camp'}},
  {{key:'social',       label:'Social Selling', full:'Social Selling Workshop: Building Your LinkedIn Brand & Driving Engagement'}},
  {{key:'capstone',      label:'Capstone Workshop', full:'Accelerate: Capstone Presentation Certification Workshop'}},
  {{key:'certification', label:'Certification', full:'Accelerate Certification course'}},
];

/* ── Tenure filter (shared by both tabs) ──
   New hire = hired within the last 65 days (hire date to today, inclusive); Tenured = more than 65 days. */
const NEW_HIRE_DAYS = 65;
let tenureFilter = 'all';   // 'all' | 'new' | 'tenured'
function daysSinceHire(p) {{
  if (!p.hireDate) return null;
  const d = new Date(p.hireDate + 'T00:00:00');
  if (isNaN(d)) return null;
  return Math.round((TODAY - d) / 86400000);   // round, not floor: stays correct across daylight-saving changes
}}
function isNewHire(p) {{ const n = daysSinceHire(p); return n !== null && n <= NEW_HIRE_DAYS; }}
function passesTenure(p) {{
  if (tenureFilter === 'all') return true;
  const n = daysSinceHire(p);
  if (n === null) return false;                 // no hire date on file: only shown under "All"
  return tenureFilter === 'new' ? n <= NEW_HIRE_DAYS : n > NEW_HIRE_DAYS;
}}
function setTenure(val, btn) {{
  tenureFilter = val;
  document.querySelectorAll('#tenure-btns button').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  applyFilters();
}}
function resetTenure() {{
  tenureFilter = 'all';
  document.querySelectorAll('#tenure-btns button').forEach((b, i) => b.classList.toggle('active', i === 0));
}}

/* Same filters as the Onboarding tab, minus Status (which is onboarding-specific). */
function passesBaseFilters(p, mkt) {{
  if (hideTLG && TLG_SET.has(p.name.toLowerCase())) return false;
  if (testGroupFilter === 'hide' && TEST_GROUP_NAMES.has(p.name)) return false;
  if (testGroupFilter === 'only' && !TEST_GROUP_NAMES.has(p.name)) return false;
  if (locationFilter === 'us' && p.isCanada) return false;
  if (locationFilter === 'ca' && !p.isCanada) return false;
  if (!passesTenure(p)) return false;
  if (mkt && p.market !== mkt) return false;
  return true;
}}

/* Onboarding counts as complete when the LMS marks the whole curriculum done (100%).
   NOTE: two Sales Skills courses (the KMPF ones) are still "(Coming Soon)" and can't be
   completed yet, so until they launch nobody reaches 100% -- the same reason the
   Onboarding tab currently shows 0% Completed. Change this one rule to count someone
   as done once every currently-available course is finished. */
function certOnboardingDone(p) {{ return !!p.overallDone || p.overallPct >= 100; }}

/* Case-insensitive lookup: report emails are often mixed-case (e.g. ARedondo@...). */
let _certIdx = null, _certIdxN = -1;
function certRec(p) {{
  const keys = Object.keys(CERT_JOURNEY);
  if (!_certIdx || _certIdxN !== keys.length) {{
    _certIdx = {{}};
    keys.forEach(k => {{ _certIdx[k.toLowerCase()] = CERT_JOURNEY[k]; }});
    _certIdxN = keys.length;
  }}
  return _certIdx[(p.email || '').toLowerCase()] || {{}};
}}

function certSteps(p) {{
  const r = certRec(p);
  const bc = r.bootCamp || {{}}, ss = r.socialSelling || {{}}, cp = r.capstone || {{}}, cr = r.certification || {{}};
  return {{
    onboarding:    {{done: certOnboardingDone(p), pct: p.overallPct}},
    bootCamp:      {{done: !!bc.attended, date: bc.date || '', reg: !!bc.registered, sessionDate: bc.sessionDate || '', sessionEnd: bc.sessionEnd || '', enrolled: bc.enrolledDate || ''}},
    social:        {{done: !!ss.attended, date: ss.date || '', reg: !!ss.registered, sessionDate: ss.sessionDate || '', sessionEnd: ss.sessionEnd || '', enrolled: ss.enrolledDate || ''}},
    capstone:      {{done: !!cp.attended, date: cp.date || '', reg: !!cp.registered, sessionDate: cp.sessionDate || '', sessionEnd: cp.sessionEnd || '', enrolled: cp.enrolledDate || ''}},
    /* Certified = the Accelerate Certification course has a completion date. */
    certification: {{done: !!(cr.completed || cr.date), date: cr.date || ''}},
  }};
}}

/* Registration. A milestone is in one of four states for each person:
   done (completed) > reg (registered, not yet completed) > none. Certification only has done / none. */
const CERT_REC_KEY = {{bootCamp: 'bootCamp', social: 'socialSelling', capstone: 'capstone'}};
const CERT_CLASS_KEYS = ['bootCamp', 'social', 'capstone'];
function certMsState(p, key) {{
  const d = certSteps(p)[key];
  if (d.done) return 'done';
  if (d.reg) return 'reg';
  return 'none';
}}
/* A session's dates: one day ("Oct 22, 2026"), a range in one month ("Nov 17–19, 2026"), or across months. */
function fmtRange(a, b) {{
  if (!a) return '';
  if (!b || b === a) return fmtDate(a);
  const da = new Date(a + 'T00:00:00'), db = new Date(b + 'T00:00:00');
  if (da.getFullYear() === db.getFullYear() && da.getMonth() === db.getMonth())
    return da.toLocaleDateString('en-US', {{month: 'short'}}) + ' ' + da.getDate() + '–' + db.getDate() + ', ' + da.getFullYear();
  return fmtDate(a) + ' – ' + fmtDate(b);
}}
/* Shown under "Registered": the session the learner is booked into. */
function certRegLabel(d) {{
  return d.sessionDate ? 'Session ' + fmtRange(d.sessionDate, d.sessionEnd) : '';
}}
function certMilestones(p) {{
  const s = certSteps(p);
  return CERT_STEPS.filter(st => s[st.key].done).length;
}}
/* Certified = the "Accelerate Certification" course has a completion date in the LMS. */
function isCertified(p) {{ return certSteps(p).certification.done; }}
/* Next step = first milestone not yet complete, in journey order. Certified people have none. */
function certNextStep(p) {{
  if (isCertified(p)) return null;
  const s = certSteps(p);
  return CERT_STEPS.find(st => !s[st.key].done) || null;
}}

function switchTab(which) {{
  certTabActive = (which === 'cert');
  document.getElementById('tab-onboarding').style.display = certTabActive ? 'none' : '';
  document.getElementById('tab-cert').style.display = certTabActive ? '' : 'none';
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === which));
  document.getElementById('f-status-wrap').style.display = certTabActive ? 'none' : 'contents';
  if (certTabActive) {{
    renderCert();
  }} else {{
    renderCharts();
    document.getElementById('result-count').textContent = filtered.length + ' shown';
  }}
}}

function resetCertControls() {{
  ['cert-ms', 'cert-search'].forEach(id => {{ const el = document.getElementById(id); if (el) el.value = ''; }});
  const b = document.getElementById('cert-sort');
  if (b) b.value = 'name';
}}

function setCertTableView(v) {{
  certTableView = v;
  document.getElementById('cert-view-individual').classList.toggle('active', v === 'individual');
  document.getElementById('cert-view-manager').classList.toggle('active', v === 'manager');
  const searchEl = document.getElementById('cert-search');
  searchEl.placeholder = v === 'manager' ? 'Search manager...' : 'Search name or manager...';
  searchEl.value = '';
  renderCertTable();
}}

function renderCert() {{
  const asof = document.getElementById('cert-asof');
  if (asof && typeof CERT_META !== 'undefined') asof.textContent = '· Completions from the ' + CERT_META.completionSource + ', through ' + fmtDate(CERT_META.completionThrough) +
      (CERT_META.registrationSource ? ' · Registrations from the ' + CERT_META.registrationSource + ', through ' + fmtDate(CERT_META.registrationThrough) : '');
  renderCertStats();
  renderCertTable();
  if (certTabActive) {{
    renderCertCharts();
    document.getElementById('result-count').textContent = certBase.length + ' shown';
  }}
}}

function renderCertStats() {{
  const total = certBase.length;
  const pctOf = n => total ? Math.round(n / total * 100) + '%' : '—';
  const onb  = certBase.filter(p => certSteps(p).onboarding.done).length;
  const boot = certBase.filter(p => certSteps(p).bootCamp.done).length;
  const soc  = certBase.filter(p => certSteps(p).social.done).length;
  const cap  = certBase.filter(p => certSteps(p).capstone.done).length;
  const cert = certBase.filter(isCertified).length;
  const avg  = total ? Math.round(certBase.reduce((s, p) => s + p.overallPct, 0) / total) : 0;
  const regN = key => certBase.filter(p => certMsState(p, key) === 'reg').length;
  const regTxt = key => ' · ' + regN(key) + ' registered';
  const set = (id, v) => {{ document.getElementById(id).textContent = v; }};
  set('cj-total', total);
  set('cj-onb', pctOf(onb));      set('cj-onb-sub', onb + ' of ' + total + ' · avg ' + avg + '% complete');
  set('cj-boot', pctOf(boot));    set('cj-boot-sub', boot + ' of ' + total + ' attended' + regTxt('bootCamp'));
  set('cj-soc', pctOf(soc));      set('cj-soc-sub', soc + ' of ' + total + ' attended' + regTxt('social'));
  set('cj-cap', pctOf(cap));      set('cj-cap-sub', cap + ' of ' + total + ' attended' + regTxt('capstone'));
  set('cj-cert', pctOf(cert));    set('cj-cert-sub', cert + ' of ' + total + ' · completed Accelerate Certification');
}}

/* ── Table ── */
function certChip(done, text, sub) {{
  return '<span class="cj-chip ' + (done ? 'done' : 'none') + '">' + text + '</span>' +
    (sub ? '<span class="cj-sub">' + sub + '</span>' : '');
}}
/* One table cell for a Boot Camp / Social Selling milestone. */
function certMsCell(p, key) {{
  const d = certSteps(p)[key];
  const st = certMsState(p, key);
  const sub = t => t ? '<span class="cj-sub">' + t + '</span>' : '';
  if (st === 'done') return '<span class="cj-chip done">✓ Attended</span>' + sub(d.date ? fmtDate(d.date) : '');
  if (st === 'reg')  return '<span class="cj-chip reg">Registered</span>' + sub(certRegLabel(d));
  return '<span class="cj-chip none">Not registered</span>';
}}
/* Certified column: plain Yes / No. */
function certCertifiedCell(p) {{
  return isCertified(p) ? '<span class="cj-chip done">Yes</span>' : '<span class="cj-chip none">No</span>';
}}
function certFilteredRows(opts) {{
  opts = opts || {{}};
  const applySearch = opts.applySearch !== false;
  const ms    = document.getElementById('cert-ms').value;
  const sort  = document.getElementById('cert-sort').value;
  const q     = document.getElementById('cert-search').value.trim().toLowerCase();
  let rows = certBase.filter(p => {{
    if (applySearch && q && !(p.name.toLowerCase().includes(q) || (p.manager || '').toLowerCase().includes(q))) return false;
    if (ms) {{ const parts = ms.split(':'); if (certMsState(p, parts[0]) !== parts[1]) return false; }}
    return true;
  }});
  if (sort === 'ms-desc')       rows.sort((a, b) => (isCertified(b) - isCertified(a)) || certMilestones(b) - certMilestones(a) || b.overallPct - a.overallPct || a.name.localeCompare(b.name));
  else if (sort === 'ms-asc')   rows.sort((a, b) => (isCertified(a) - isCertified(b)) || certMilestones(a) - certMilestones(b) || a.overallPct - b.overallPct || a.name.localeCompare(b.name));
  else if (sort === 'onb-desc') rows.sort((a, b) => b.overallPct - a.overallPct || a.name.localeCompare(b.name));
  else                          rows.sort((a, b) => a.name.localeCompare(b.name));
  return rows;
}}
function certVisibleRows() {{ return certFilteredRows(); }}

function certPersonRow(p) {{
  const onb = '<span class="pct-pill" style="' + overallPillStyle(p) + '">' + p.overallPct + '%</span>';
  return '<tr data-email="' + escHtml(p.email) + '" onclick="openCertModal(this.dataset.email)">' +
    '<td style="font-weight:600;white-space:nowrap">' + escHtml(p.name) + '<span class="cj-sub">' + escHtml(p.market) + '</span></td>' +
    '<td style="font-size:12px;color:var(--muted);white-space:nowrap">' + escHtml(p.manager || '—') + '</td>' +
    '<td class="cj-c">' + onb + '</td>' +
    '<td class="cj-c">' + certMsCell(p, 'bootCamp') + '</td>' +
    '<td class="cj-c">' + certMsCell(p, 'social') + '</td>' +
    '<td class="cj-c">' + certMsCell(p, 'capstone') + '</td>' +
    '<td class="cj-c">' + certCertifiedCell(p) + '</td></tr>';
}}

function renderCertTable() {{
  document.getElementById('cert-head').innerHTML = '<tr>' +
    '<th>Learner</th><th>Manager</th>' +
    '<th class="cj-c">Onboarding</th><th class="cj-c">Boot Camp</th><th class="cj-c">Social Selling</th><th class="cj-c" title="Accelerate: Capstone Presentation Certification Workshop">Capstone Workshop</th><th class="cj-c">Certified</th></tr>';
  const body = document.getElementById('cert-body');

  if (certTableView === 'manager') {{
    const rows = certFilteredRows({{applySearch: false}});
    if (!rows.length) {{
      body.innerHTML = '<tr style="cursor:default"><td colspan="7" style="text-align:center;color:var(--muted);padding:28px;">No learners match the current filters.</td></tr>';
      return;
    }}
    const q = document.getElementById('cert-search').value.trim().toLowerCase();
    const groups = {{}};
    rows.forEach(p => {{
      const mgr = p.manager || '(No Manager)';
      if (!groups[mgr]) groups[mgr] = [];
      groups[mgr].push(p);
    }});
    const mgrNames = Object.keys(groups).sort().filter(mgr => !q || mgr.toLowerCase().includes(q));
    if (!mgrNames.length) {{
      body.innerHTML = '<tr style="cursor:default"><td colspan="7" style="text-align:center;color:var(--muted);padding:28px;">No managers match the current filters.</td></tr>';
      return;
    }}
    let html = '';
    mgrNames.forEach(mgr => {{
      const team = groups[mgr];
      const cert = team.filter(isCertified).length;
      const certStr = cert > 0 ? ' &nbsp;·&nbsp; <span style="color:var(--green);font-weight:600">' + cert + ' certified</span>' : '';
      html += '<tr class="mgr-group-hdr-row" data-manager="' + escHtml(mgr.toLowerCase()) + '"><td colspan="7">' +
        '<div class="mgr-group-hdr open" onclick="toggleMgrGroup(this)">' +
        '<span style="font-weight:700;font-size:13px;">' + escHtml(mgr) + '</span>' +
        '<span style="font-size:11px;color:var(--muted);">' + team.length + ' learner' + (team.length !== 1 ? 's' : '') + certStr + ' &nbsp;<span class="mgr-chevron">&#9660;</span></span>' +
        '</div></td></tr>';
      team.forEach(p => {{ html += certPersonRow(p); }});
    }});
    body.innerHTML = html;
  }} else {{
    const rows = certVisibleRows();
    if (!rows.length) {{
      body.innerHTML = '<tr style="cursor:default"><td colspan="7" style="text-align:center;color:var(--muted);padding:28px;">No learners match the current filters.</td></tr>';
      return;
    }}
    body.innerHTML = rows.map(certPersonRow).join('');
  }}
}}

/* ── Market charts ── */
function renderCertCharts() {{
  const markets = [...new Set(certBase.map(p => p.market))].sort();
  const pc = (n, t) => t ? Math.round(n / t * 100) : 0;
  const stats = markets.map(m => {{
    const mp = certBase.filter(p => p.market === m);
    const t = mp.length;
    return {{
      t,
      cert: mp.filter(isCertified).length,
      boot: mp.filter(p => certSteps(p).bootCamp.done).length,
      soc:  mp.filter(p => certSteps(p).social.done).length,
      cap:  mp.filter(p => certSteps(p).capstone.done).length,
    }};
  }});
  // three bars per market on the attendance chart, so size both charts for that and keep the cards level
  const h = Math.max(260, markets.length * 62) + 'px';
  document.getElementById('certMarketCertChart').closest('.chart-wrap').style.height = h;
  document.getElementById('certMarketAttChart').closest('.chart-wrap').style.height = h;

  const grid = cv('border'), muted = cv('muted');
  const axisX = {{min: 0, max: 100, grid: {{color: grid}}, ticks: {{color: muted, font: {{size: 11}}, callback: v => v + '%'}}, border: {{color: 'transparent'}}}};
  const axisY = {{grid: {{display: false}}, ticks: {{color: muted, font: {{size: 11}}, autoSkip: false}}, border: {{color: 'transparent'}}}};
  const legend = {{display: true, position: 'bottom', labels: {{color: muted, font: {{size: 11}}, boxWidth: 12, padding: 12}}}};
  const total = i => stats[i].t;

  // 1) Certified by market (stacked to 100%: Certified / Not certified)
  const green = cv('green'), grey = '#6b7280';
  const c1 = document.getElementById('certMarketCertChart').getContext('2d');
  if (certMarketCertChartObj) certMarketCertChartObj.destroy();
  certMarketCertChartObj = new Chart(c1, {{
    type: 'bar',
    data: {{labels: markets, datasets: [
      {{label: 'Certified', data: stats.map(s => pc(s.cert, s.t)), backgroundColor: green + 'cc', borderColor: green, borderWidth: 1, maxBarThickness: 26}},
      {{label: 'Not certified', data: stats.map(s => 100 - pc(s.cert, s.t)), backgroundColor: grey + '99', borderColor: grey, borderWidth: 1, maxBarThickness: 26}},
    ]}},
    options: {{
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: {{legend, tooltip: {{callbacks: {{
        label: c => {{
          const s = stats[c.dataIndex];
          const n = c.datasetIndex === 0 ? s.cert : s.t - s.cert;
          return '  ' + c.dataset.label + ': ' + c.parsed.x + '% (' + n + ')';
        }},
        afterBody: items => ['  Total in this market: ' + total(items[0].dataIndex)]
      }}}}}},
      scales: {{x: Object.assign({{stacked: true}}, axisX), y: Object.assign({{stacked: true}}, axisY)}}
    }}
  }});

  // 2) Attendance by market (grouped: Boot Camp / Social Selling attended)
  const blue = cv('accent'), teal = cv('teal'), purple = cv('accent2');
  const c2 = document.getElementById('certMarketAttChart').getContext('2d');
  if (certMarketAttChartObj) certMarketAttChartObj.destroy();
  certMarketAttChartObj = new Chart(c2, {{
    type: 'bar',
    data: {{labels: markets, datasets: [
      {{label: 'Advanced Boot Camp attended', data: stats.map(s => pc(s.boot, s.t)), backgroundColor: blue + 'cc', borderColor: blue, borderWidth: 1, maxBarThickness: 14}},
      {{label: 'Social Selling Workshop attended', data: stats.map(s => pc(s.soc, s.t)), backgroundColor: teal + 'cc', borderColor: teal, borderWidth: 1, maxBarThickness: 14}},
      {{label: 'Capstone Workshop attended', data: stats.map(s => pc(s.cap, s.t)), backgroundColor: purple + 'cc', borderColor: purple, borderWidth: 1, maxBarThickness: 14}},
    ]}},
    options: {{
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: {{legend, tooltip: {{callbacks: {{
        label: c => {{
          const s = stats[c.dataIndex];
          const n = [s.boot, s.soc, s.cap][c.datasetIndex];
          return '  ' + c.dataset.label + ': ' + c.parsed.x + '% (' + n + ' of ' + s.t + ')';
        }}
      }}}}}},
      scales: {{x: axisX, y: axisY}}
    }}
  }});
}}

/* ── Detail modal (journey view) ── */
function openCertModal(email) {{
  const p = PEOPLE.find(x => x.email === email);
  if (!p) return;
  const s = certSteps(p);
  const n = certMilestones(p);
  const next = certNextStep(p);
  const sessionTxt = d => {{ const t = certRegLabel(d); return t ? ' · ' + t : ''; }};

  function step(i, st) {{
    const d = s[st.key];
    let cls = '', status = 'Not yet';
    if (st.key === 'onboarding') {{
      if (d.done) {{ cls = 'done'; status = '100% complete'; }}
      else if (d.pct > 0) {{ cls = 'active'; status = d.pct + '% complete'; }}
    }} else if (st.key === 'certification') {{
      if (d.done) {{ cls = 'done'; status = 'Certified' + (d.date ? ' · ' + fmtDate(d.date) : ''); }}
    }} else {{
      const ms = certMsState(p, st.key);
      if (ms === 'done')      {{ cls = 'done'; status = 'Attended' + (d.date ? ' · ' + fmtDate(d.date) : ''); }}
      else if (ms === 'reg')  {{ cls = 'reg'; status = 'Registered' + sessionTxt(d); }}
      else                    {{ status = 'Not registered'; }}
    }}
    return '<div class="cj-step ' + cls + '"><div class="cj-num">' + (cls === 'done' ? '✓' : (i + 1)) + '</div>' +
      '<div class="cj-step-label">' + escHtml(st.label) + '</div><div class="cj-step-status">' + escHtml(status) + '</div></div>';
  }}

  function classRow(label, key) {{
    const d = s[key];
    const ms = certMsState(p, key);
    let status, detail = '';
    if (ms === 'done')     {{ status = 'Attended'; detail = d.date ? fmtDate(d.date) : ''; }}
    else if (ms === 'reg') {{ status = 'Registered'; detail = certRegLabel(d) + (d.enrolled ? ' · enrolled ' + fmtDate(d.enrolled) : ''); }}
    else                   {{ status = 'Not registered'; }}
    return [label, status, detail];
  }}

  const detailRows = [
    ['Accelerate Onboarding curriculum', s.onboarding.done ? 'Complete' : s.onboarding.pct + '% complete', ''],
    classRow('Accelerate Advanced Boot Camp', 'bootCamp'),
    classRow('Social Selling Workshop: Building Your LinkedIn Brand & Driving Engagement', 'social'),
    classRow('Accelerate: Capstone Presentation Certification Workshop', 'capstone'),
    ['Accelerate Certification course', s.certification.done ? 'Complete' : 'Not yet', s.certification.date ? fmtDate(s.certification.date) : ''],
  ].map(x => '<tr><td>' + escHtml(x[0]) + '</td><td class="c">' + escHtml(x[1]) + '</td><td class="c">' + (x[2] ? escHtml(x[2]) : '—') + '</td></tr>').join('');

  const curricPills = CURRIC_IDS.map(cid => {{
    const c = p.curricula[cid];
    if (!c) return '';
    return '<div class="metric-box"><div class="mb-label">' + escHtml(CURRIC_NAMES[cid] || c.title) + '</div>' +
      '<div class="mb-value"><span class="pct-pill" style="' + curricPillStyle(c, null) + ';font-size:12px">' + c.pct + '%</span></div>' +
      '<div class="mb-sub">' + c.done + ' of ' + c.total + ' items</div></div>';
  }}).join('');

  const classLabel = {{bootCamp: 'Boot Camp', social: 'Social Selling', capstone: 'Capstone Workshop'}};
  const upcoming = CERT_CLASS_KEYS.filter(k => certMsState(p, k) === 'reg')
    .map(k => classLabel[k] + (certRegLabel(s[k]) ? ' (' + certRegLabel(s[k]) + ')' : ''));
  const regNote = upcoming.length ? ' · Registered: ' + upcoming.join(', ') : '';
  const banner = isCertified(p)
    ? '<div class="rc-riskbar risk-green"><span class="rc-risklabel">★ Certified</span><span class="rc-riskaction">Accelerate Certification completed' + (s.certification.date ? ' ' + escHtml(fmtDate(s.certification.date)) : '') + '.</span></div>'
    : '<div class="rc-riskbar risk-amber"><span class="rc-risklabel">Not certified yet</span><span class="rc-riskaction">Next step: <b>' + escHtml((next || CERT_STEPS[CERT_STEPS.length - 1]).full) + '</b>' + escHtml(regNote) + '</span></div>';

  document.getElementById('modal-content').innerHTML =
    '<div class="modal-header"><div class="modal-name">' + escHtml(p.name) + '</div>' +
      '<div class="modal-meta"><span>' + escHtml(p.jobTitle || '') + '</span><span>' + escHtml(p.market) + '</span>' +
      (p.manager ? '<span>Manager: ' + escHtml(p.manager) + '</span>' : '') + '</div></div>' +
    '<div class="modal-body">' + banner +
      '<div style="font-size:12px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px;">Certification Journey</div>' +
      '<div class="cj-stepper">' + CERT_STEPS.map((st, i) => step(i, st)).join('') + '</div>' +
      '<table class="sched-table"><thead><tr><th>Milestone</th><th class="c">Status</th><th class="c">Date / Detail</th></tr></thead><tbody>' + detailRows + '</tbody></table>' +
      '<div style="font-size:12px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px;">Onboarding Curriculum</div>' +
      '<div class="metric-grid">' + curricPills + '</div>' +
      '<button class="btn-reset" onclick="openModal(\\'' + p.email.replace(/'/g, "\\\\'") + '\\')">View full onboarding detail &rarr;</button>' +
    '</div>';
  document.getElementById('modal-overlay').classList.add('open');
}}

/* ── Exports ── */
function certExportRow(p) {{
  const s = certSteps(p);
  const txt = key => {{ const ms = certMsState(p, key); return ms === 'done' ? 'Attended' : ms === 'reg' ? 'Registered' : 'Not registered'; }};
  const dt = key => {{ const d = s[key], ms = certMsState(p, key); return ms === 'reg' ? (d.sessionDate ? (d.sessionEnd && d.sessionEnd !== d.sessionDate ? d.sessionDate + ' to ' + d.sessionEnd : d.sessionDate) : '') : (d.date || ''); }};
  return [
    p.name, p.email, p.market, p.manager || '',
    p.overallPct + '%',
    txt('bootCamp'), dt('bootCamp'),
    txt('social'),   dt('social'),
    txt('capstone'), dt('capstone'),
    certMilestones(p) + '/' + CERT_STEPS.length,
    isCertified(p) ? 'Yes' : 'No',
    s.certification.date || ''
  ];
}}
const CERT_EXPORT_HEAD = ['Name','Email','Market','Manager','Onboarding %','Boot Camp','Boot Camp Date / Session','Social Selling','Social Selling Date / Session','Capstone Workshop','Capstone Date / Session','Milestones','Certified','Certification Date'];

function runCertExportXLSX() {{
  document.getElementById('export-menu').classList.remove('open');
  const rows = [CERT_EXPORT_HEAD].concat(certVisibleRows().map(certExportRow));
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [28,32,18,26,13,13,15,14,18,16,18,11,10,17].map(w => ({{wch: w}}));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Certification Journey');
  XLSX.writeFile(wb, 'accelerate-certification-journey.xlsx');
}}

function runCertExport() {{
  document.getElementById('export-menu').classList.remove('open');
  const rows = certVisibleRows();
  const now = new Date().toLocaleDateString('en-US', {{month: 'long', day: 'numeric', year: 'numeric'}});
  document.getElementById('ph-report-type').textContent = 'Certification Journey';
  document.getElementById('ph-sub').textContent = rows.length + ' learners · ' + now + ' · Certified = Accelerate Certification course completed';
  const head = ['Name','Market','Manager','Onboarding %','Boot Camp','Social Selling','Capstone Workshop','Milestones','Certified','Certified On'];
  const body = rows.map(p => {{
    const r = certExportRow(p);
    return '<tr><td>' + escHtml(r[0]) + '</td><td>' + escHtml(r[2]) + '</td><td>' + escHtml(r[3]) + '</td><td>' + r[4] + '</td>' +
      '<td>' + r[5] + '</td><td>' + r[7] + '</td><td>' + r[9] + '</td><td>' + r[11] + '</td><td>' + r[12] + '</td><td>' + (r[13] || '') + '</td></tr>';
  }}).join('');
  document.getElementById('print-roster-wrap').innerHTML =
    '<table><thead><tr>' + head.map(x => '<th>' + x + '</th>').join('') + '</tr></thead><tbody>' + body + '</tbody></table>';
  window.print();
}}

/* ── Info popovers ── */
/* Add the Tenure filter to the shared filter help text. */
INFO['filters-info'] = INFO['filters-info'].replace('Use Reset to clear everything.',
  'Hire Status: New Hires shows reps hired within the last 65 days; Tenured shows reps hired more than 65 days ago (applies on both tabs). Use Reset to clear everything.');
Object.assign(INFO, {{
  "cert-total": "Everyone currently shown on this tab. The Location, Hire Status, Market and Group filters apply here too; the Status filter is specific to the Onboarding tab and is hidden on this one.",
  "cert-onboarding": "People who have completed 100% of the Accelerate Onboarding curriculum. This is live from the LMS data. The sub-line shows the average completion percentage for everyone shown.",
  "cert-bootcamp": "People who completed the Accelerate Advanced Skills Bootcamp. The sub-line also shows how many are Registered, meaning booked into an upcoming session in the registration report.",
  "cert-social": "People who completed the Social Selling Workshop: Building Your LinkedIn Brand & Driving Engagement. The sub-line also shows how many are Registered, meaning booked into a session in the registration report.",
  "cert-certified": "People who have completed the Accelerate Certification course in the LMS. Anyone with a completion date for that course is marked Certified, regardless of the other milestones.",
  "cert-market-certified": "Each market's learners split into Certified (completed the Accelerate Certification course) and Not certified. Hover a bar for exact counts and the market's total. Follows the Location, Market and Group filters.",
  "cert-market-attendance": "For each market, the share of its learners who attended the Advanced Boot Camp, the Social Selling Workshop and the Capstone Workshop. Registered learners aren't counted as attended until the class is completed. Hover a bar for exact counts.",
  "cert-capstone": "People who attended the Accelerate: Capstone Presentation Certification Workshop. The sub-line also shows how many are Registered, meaning booked into a session in the registration report.",
  "cert-table": "One row per learner. Boot Camp, Social Selling and Capstone Workshop each show Attended (with the completion date), Registered (booked into a session, with the session dates), or Not registered. Being assigned a class in the LMS without a booked session does not count as registered. Certified reads Yes once the Accelerate Certification course has a completion date, otherwise No. Use the Milestone filter to find, for example, everyone not yet registered for Boot Camp, or everyone who is Certified.",
  "cert-view-toggle": "Individual view shows one row per learner. By Manager groups learners under their manager so you can see how each team is tracking against the Certification Journey. The search box updates automatically — in Individual view it searches by learner name or manager, in By Manager view it filters to matching manager groups only.",
}});

/* ── Init ── */
function init() {{
  // Populate market dropdown (based on currently visible group, not full roster)
  populateMarketDropdown();

  // Close dropdowns on outside click
  document.addEventListener('click', e => {{
    if (!document.getElementById('export-drop').contains(e.target))
      document.getElementById('export-menu').classList.remove('open');
    if (!document.getElementById('info-popover').contains(e.target) && !e.target.classList.contains('info-btn'))
      document.getElementById('info-popover').classList.remove('visible');
  }});

  applyFilters();
}}

function runExportXLSX(type){{
  document.getElementById('export-menu').classList.remove('open');
  const now=new Date().toLocaleDateString('en-US',{{year:'numeric',month:'long',day:'numeric'}});
  function makeSheet(rows,colWidths){{
    const ws=XLSX.utils.aoa_to_sheet(rows);
    ws['!cols']=colWidths.map(w=>({{wch:w}}));
    return ws;
  }}
  function dlXLSX(name,wb){{ XLSX.writeFile(wb,name+'.xlsx'); }}
  const wb=XLSX.utils.book_new();
  if(type==='full'){{
    const rows=[['Name','Market','Status','Days Left','Overall %','Manager']];
    filtered.forEach(p=>{{
      const status=computeStatus(p);
      const d=daysLeft(p);
      const daysStr=p.overallDone?'Complete':d===null?'—':d<0?(Math.abs(d)+'d overdue'):(d+'d left');
      rows.push([p.name,p.market,status,daysStr,p.overallPct+'%',p.manager]);
    }});
    XLSX.utils.book_append_sheet(wb,makeSheet(rows,[28,18,14,14,12,28]),'Full Report');
    dlXLSX('onboarding-full-report',wb);
  }} else if(type==='overdue'){{
    const od=filtered.filter(p=>computeStatus(p)==='Overdue');
    const rows=[['Name','Email','Market','Completion %','Manager','Manager Email']];
    od.forEach(p=>rows.push([p.name,p.email,p.market,p.overallPct+'%',p.manager,p.mgrEmail]));
    XLSX.utils.book_append_sheet(wb,makeSheet(rows,[28,32,18,14,28,32]),'Overdue Only');
    dlXLSX('onboarding-overdue',wb);
  }} else if(type==='manager-summary'){{
    const byMgr={{}};
    filtered.forEach(p=>{{
      const m=p.manager||'Unknown';
      if(!byMgr[m]) byMgr[m]={{mgrEmail:p.mgrEmail,people:[]}};
      byMgr[m].people.push(p);
    }});
    const rows=[['Manager','Manager Email','Team Count','Overdue','Avg Completion %']];
    Object.entries(byMgr).sort((a,b)=>a[0].localeCompare(b[0])).forEach(([mgr,data])=>{{
      const avg=Math.round(data.people.reduce((s,p)=>s+p.overallPct,0)/data.people.length);
      const od=data.people.filter(p=>computeStatus(p)==='Overdue').length;
      rows.push([mgr,data.mgrEmail,data.people.length,od,avg+'%']);
    }});
    XLSX.utils.book_append_sheet(wb,makeSheet(rows,[28,32,12,10,16]),'Manager Summary');
    dlXLSX('onboarding-manager-summary',wb);
  }}
}}

init();
if (location.hash === '#cert') switchTab('cert');
(function(){{
  var n=0,t;
  var h=document.querySelector('.header h1');
  if(h) h.addEventListener('click',function(){{
    n++;clearTimeout(t);
    if(n>=3){{n=0;window.location.href='index.html';}}
    else t=setTimeout(function(){{n=0;}},1500);
  }});
}})();
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
    records = attach_registration_status(records)
    sales_map = _load_salesforce(records)
    print(f"Salesforce: {len(sales_map)} cohort members with a Closed Won deal")
    html = generate_html(records, sales_map)
    with open('onboarding.html', 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"Generated onboarding.html")


if __name__ == '__main__':
    main()

