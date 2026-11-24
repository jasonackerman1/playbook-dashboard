"""
Sales Training & L&D Engagement Dashboard generator.

Built 2026-11-24 from a new dashboard Resmie delivered (Sales_Training_LD_Dashboard.html +
"Sales Training & L&D Engagement Dashboard Guide.docx"). Shows executive leadership how much
sales training people completed this fiscal year (April 1 to the report date), who completed
it, and which live classes are coming up. Three tabs: All Groups, Sales Org (SF/SFC), Dealers.

INPUTS (sales-training-data/, newest file of each kind by MM.DD.YYYY in the name):
  LMS-Course-Status-Data-*.xlsx      LMS learning export (EmployeeLearningCSV_BD2_ACTIVE):
                                     every completion, delivery method, status, hire date, etc.
  Master-User-Data-*.xlsx            Master User Data: the SF/SFC roster for Reach
  Registration-Status-Future-*.xlsx  Registration Status, future classes: Upcoming live classes
  Catalog-Report-*.xlsx              Catalog: only used to note whether a left-out course is
                                     catalog-flagged SALES (helps decide if it should be added)
  sales-course-list.json             The sales course list + categories, frozen from Resmie's
                                     page. A course not on it is LEFT OUT and named in the log
                                     (her guide: "New courses need a decision").

PAGE: templates/sales-training-shell.html is Resmie's page with only the data swapped for
placeholders and the house look applied (KM logo header, "Data through", house palette with
light/dark via pb-theme, flat filter bar, full width, "?" help, Show/Hide TLG, Export PDF/Excel
from templates/sales-training-house.js). Her chart/table math is unchanged.

VERIFIED 2026-11-24: rebuilt from the 11.24 files, the completion data is identical row for row
to her page (13,640 completions, 1,918 people, 265 branches), and the roster (1,024) and
upcoming-class list (67) match exactly. Only the order of rows differs (hers came from another
copy of the export); that affects nothing but the order of exact ties.

RULES (from her guide):
  - Completion = Completion Status starts with COMPLETE, dated April 1 of the fiscal year
    through the report date; one per person-course (earliest date kept).
  - Report date = the newest completion date anywhere in the learning export.
  - Instructor-led = delivery ILT or VILT, OR a completion status of COMPLETE_VILT / COMPLETE_ILT even
    when the course is set up as ONLINE (added 2026-11-24 for Accelerate: Capstone Presentation
    Certification Workshop, set up ONLINE in the LMS but completed as VILT). Complete ILT (*) = status COMPLETE_ILT, except
    Pipeline Powerlifting for Managed IT, always treated as VILT.
  - Business unit: Dealers (DEALER_ domain) > IIM (job or manager title mentions IIM,
    Intelligent Information Management or BlueIrisIQ) > AC (allcovered.com email or an
    EMP_TYP_ID starting AC, which includes IT Weapons' ACC) > BCA (bt.konicaminolta.ca email
    or EMP_BCA) > BUS.
  - Sales field: EMP_TYP_ID SF = 1 (US), SFC = 2 (Canada). Dealer country: DEALER_US / DEALER_BCD.
  - Rep status: dealers and people with no hire date = none; new hire if the completion is
    within 65 days of the hire date; tenured after.
  - Blank branch = "Unassigned".
  - Roster = Master User Data rows with EMP_TYP_ID SF or SFC, in file order, matched by STUD_ID
    to the learning export's Employee ID.
  - Upcoming: one row per person per class, classes starting on or after the report date;
    group = Dealers / All Covered / Sales team (SF, SFC) / Other employees; BCA = Canadian email;
    new hire = class starts within 65 days of hire; instructors = every instructor on the class.
  - TLG (house rule): a "t" flag per completion, roster member and registration; hidden by
    default, Show TLG adds them back.
"""
import datetime
import glob
import json
import os
import re
from collections import Counter

import openpyxl

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(SCRIPT_DIR, 'sales-training-data')
SHELL = os.path.join(SCRIPT_DIR, 'templates', 'sales-training-shell.html')
HOUSE_JS = os.path.join(SCRIPT_DIR, 'templates', 'sales-training-house.js')
OUT = os.path.join(SCRIPT_DIR, 'sales-training.html')

TLG = ['Jason Ackerman', 'Bianca Davis', 'James Parker', 'Resmie Biba', 'Chris Curtis', 'Sara Thompson',
       'Jeremy MacBean', 'Bradley Pierce', 'Laura Sefcik', 'Samantha Maresca', 'Staci Musco', 'CJ Homer',
       'Rich Moore', 'Dale Kinsey', 'John Lechner', 'Resmie Nesimi', "Samantha D'Angelo", 'Bianca DiPasquale',
       'Doug Falk']
TLG_LC = {n.lower() for n in TLG}
NEW_HIRE_DAYS = 65
IIM_RE = re.compile(r'\bIIM\b|intelligent information management|blueirisiq', re.I)
MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September',
          'October', 'November', 'December']


def _file_date(fname):
    m = re.search(r'(\d{2})\.(\d{2})\.(\d{4})', os.path.basename(fname))
    return f'{m.group(3)}-{m.group(1)}-{m.group(2)}' if m else '0000-00-00'


def _newest(pattern):
    files = sorted(glob.glob(os.path.join(DATA_DIR, pattern)), key=lambda f: (_file_date(f), f))
    return files[-1] if files else None


def _long(iso_):
    d = datetime.date.fromisoformat(iso_)
    return f'{MONTHS[d.month - 1]} {d.day}, {d.year}'


def _short(iso_):
    d = datetime.date.fromisoformat(iso_)
    return f'{MONTHS[d.month - 1][:3]} {d.day}, {d.year}'


def _rows(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    it = wb.worksheets[0].iter_rows(values_only=True)
    hdr = [str(h).strip() if h is not None else '' for h in next(it)]
    out = [dict(zip(hdr, r)) for r in it if r and any(v is not None for v in r)]
    wb.close()
    return out


def _txt(v):
    return '' if v is None else str(v).strip()


def _iso(v):
    if v is None or v == '':
        return ''
    if hasattr(v, 'strftime'):
        return v.strftime('%Y-%m-%d')
    if isinstance(v, (int, float)):
        return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))).isoformat()
    m = re.match(r'\s*(\d{1,2})/(\d{1,2})/(\d{4})', str(v))
    if m:
        return f'{m.group(3)}-{int(m.group(1)):02d}-{int(m.group(2)):02d}'
    m = re.match(r'\s*(\d{4}-\d{2}-\d{2})', str(v))
    return m.group(1) if m else ''


def _day(iso_):
    return (datetime.date.fromisoformat(iso_) - datetime.date(1970, 1, 1)).days


def _is_tlg(first, last):
    return f'{_txt(first)} {_txt(last)}'.strip().lower() in TLG_LC


def _bu(r):
    dom, em, et = _txt(r['Domain']), _txt(r['Email Address']).lower(), _txt(r['EMP_TYP_ID'])
    if dom.startswith('DEALER'):
        return 4
    if IIM_RE.search(_txt(r['Job Title'])) or IIM_RE.search(_txt(r['Manager Job Title'])):
        return 3
    if em.endswith('allcovered.com') or et.startswith('AC'):
        return 2
    if em.endswith('bt.konicaminolta.ca') or dom == 'EMP_BCA':
        return 1
    return 0


def build_data(log=print):
    lms_f = _newest('LMS-Course-Status-Data-*.xlsx')
    mud_f = _newest('Master-User-Data-*.xlsx')
    reg_f = _newest('Registration-Status-Future-*.xlsx')
    cat_f = _newest('Catalog-Report-*.xlsx')
    if not (lms_f and mud_f and reg_f):
        raise FileNotFoundError('sales-training-data/ needs LMS-Course-Status-Data, Master-User-Data and '
                                'Registration-Status-Future files')
    cl = json.load(open(os.path.join(DATA_DIR, 'sales-course-list.json'), encoding='utf-8'))
    courses, cats = cl['courses'], cl['cats']
    cidx = {c[0]: i for i, c in enumerate(courses)}

    lms = _rows(lms_f)
    asof = max(_iso(r.get('Completion Date')) for r in lms if _txt(r.get('Completion Status')).startswith('COMPLETE'))
    a = datetime.date.fromisoformat(asof)
    fy0 = datetime.date(a.year if a.month >= 4 else a.year - 1, 4, 1).isoformat()

    recs, left_out = {}, Counter()
    for n, r in enumerate(lms):
        if not _txt(r['Completion Status']).startswith('COMPLETE'):
            continue
        dt = _iso(r['Completion Date'])
        if not dt or dt < fy0 or dt > asof:
            continue
        cid = _txt(r['Course ID'])
        if cid not in cidx:
            left_out[(cid, _txt(r['Course Title']))] += 1
            continue
        key = (_txt(r['Employee ID']), cid)
        if key not in recs or dt < recs[key][0]:
            recs[key] = (dt, n, r)

    emps = sorted({k[0] for k in recs})
    eidx = {e: i for i, e in enumerate(emps)}
    brname = lambda r: _txt(r['Branch']) or 'Unassigned'
    branches = sorted({brname(r) for _, _, r in recs.values()})
    bridx = {b: i for i, b in enumerate(branches)}
    rows = []
    for (emp, cid), (dt, n, r) in recs.items():
        dom, et, hire = _txt(r['Domain']), _txt(r['EMP_TYP_ID']), _iso(r['Hire Date'])
        if dom.startswith('DEALER') or not hire:
            s = 0
        elif _day(dt) - _day(hire) <= NEW_HIRE_DAYS:
            s = 1
        else:
            s = 2
        status = _txt(r['Completion Status'])
        rows.append(dict(
            e=eidx[emp], c=cidx[cid], b=_bu(r), o={'SF': 1, 'SFC': 2}.get(et, 0), r=0, br=bridx[brname(r)],
            v=1 if _txt(r['Course Delivery Method']) in ('ILT', 'VILT') or status in ('COMPLETE_VILT', 'COMPLETE_ILT') else 0,
            k=1, x=0, d=_day(dt), s=s,
            i=1 if status == 'COMPLETE_ILT' and 'pipeline powerlifting' not in _txt(r['Course Title']).lower() else 0,
            dc={'DEALER_US': 1, 'DEALER_BCD': 2}.get(dom, 0),
            t=1 if _is_tlg(r['First Name'], r['Last Name']) else 0, _n=n))
    rows.sort(key=lambda x: (x['c'], x['_n']))
    cols = {k: [x[k] for x in rows] for k in ('e', 'c', 'b', 'o', 'r', 'br', 'v', 'k', 'x', 'd', 's', 'i', 'dc', 't')}

    # roster: Master User Data SF/SFC, file order
    mud = _rows(mud_f)
    roster = [[eidx.get(_txt(r['STUD_ID']), -1), 1 if _txt(r['EMP_TYP_ID']) == 'SF' else 2, _iso(r['HIRE_DTE']),
               1 if _is_tlg(r['FNAME'], r['LNAME']) else 0]
              for r in mud if _txt(r['EMP_TYP_ID']) in ('SF', 'SFC')]

    # upcoming live classes
    reg = _rows(reg_f)
    mud_by = {_txt(r['STUD_ID']): r for r in mud}
    lms_by = {}
    for r in lms:
        lms_by.setdefault(_txt(r['Employee ID']), r)
    ins = {}
    for r in reg:
        ins.setdefault(_txt(r['Class ID']), set()).add(f"{_txt(r['Instructor First Name'])} {_txt(r['Instructor Last Name'])}".strip())
    up = {}
    for r in reg:
        key = (_txt(r['Class ID']), _txt(r['User ID']))
        start = _iso(r['Start Date/Time'])
        if key in up or not start or start < asof:
            continue
        em = _txt(r['Student Email']).lower()
        l, m = lms_by.get(key[1]), mud_by.get(key[1])
        et = _txt((l and l['EMP_TYP_ID']) or (m and m['EMP_TYP_ID']))
        dom = _txt((l and l['Domain']) or (m and m['DMN_ID']))
        if dom.startswith('DEALER'):
            grp = 'Dealers'
        elif em.endswith('allcovered.com') or et.startswith('AC'):
            grp = 'All Covered'
        elif et in ('SF', 'SFC'):
            grp = 'Sales team'
        else:
            grp = 'Other employees'
        hd = _iso(r['HIRE_DTE'])
        nh = hd and _day(start) - _day(hd) <= NEW_HIRE_DAYS
        up[key] = [key[0], _txt(r['Title']), start, '; '.join(sorted(ins[key[0]])), grp,
                   'BCA' if em.endswith('bt.konicaminolta.ca') else 'BUS', 'New hire' if nh else 'Tenured',
                   _txt(r['Status']), 1 if _is_tlg(r['First Name'], r['Last Name']) else 0]

    # log left-out courses (with the catalog's SALES flag as a hint)
    if left_out:
        sales_flag = {}
        if cat_f:
            for r in _rows(cat_f):
                sales_flag[_txt(r.get('CPNT_ID'))] = _txt(r.get('CC_CPNT_10'))
        log(f'  Left out (not on the sales course list): {len(left_out)} courses, {sum(left_out.values())} completions')
        for (cid, title), n in left_out.most_common():
            flag = sales_flag.get(cid, 'not in catalog')
            log(f'    {cid:24} {n:5}  [{flag}] {title}')

    D = {'asof': asof, 'cats': cats, 'bus': ['BUS', 'BCA', 'AC', 'IIM', 'Dealers'], 'regions': ['-'],
         'branches': branches, 'courses': courses, 'cols': cols, 'roster': roster, 'upcoming': list(up.values())}
    meta = {'lms': lms_f, 'mud': mud_f, 'reg': reg_f, 'left_out': len(left_out)}
    return D, meta


def stats(D, tlg=False):
    """All Groups tab, fiscal year to date, TLG hidden: the homepage card numbers."""
    C = D['cols']
    keep = [i for i in range(len(C['e'])) if tlg or not C['t'][i]]
    people = len({C['e'][i] for i in keep})
    il = sum(C['v'][i] for i in keep)
    return {'people': people, 'completions': len(keep), 'online': len(keep) - il, 'il': il,
            'online_pct': round((len(keep) - il) / len(keep) * 100) if keep else 0,
            'date_label': _long(D['asof'])}


def generate_html(D, meta):
    html = open(SHELL, encoding='utf-8').read()
    house = open(HOUSE_JS, encoding='utf-8').read()
    for k, v in {'__ST_REG_DATE__': _short(_file_date(meta['reg'])),
                 '__ST_LMS_DATE__': _short(_file_date(meta['lms'])),
                 '__ST_MASTER_DATE__': _short(_file_date(meta['mud'])),
                 '__ST_REG_DATE_LONG__': _short(_file_date(meta['reg'])),
                 '__ST_LEFT_OUT__': str(meta['left_out']),
                 '__ST_HOUSE_JS__': house}.items():
        html = html.replace(k, v)
    return html.replace('__ST_DATA__', json.dumps(D, separators=(',', ':'), ensure_ascii=False))


def load_sales_training_stats():
    """For generate_homepage.py."""
    D, _ = build_data(log=lambda *a: None)
    return stats(D)


def main():
    print('Sales Training & L&D Engagement:')
    D, meta = build_data()
    for k in ('lms', 'mud', 'reg'):
        print(f'  {os.path.basename(meta[k])}')
    s = stats(D)
    print(f"  Data through {s['date_label']}: {s['people']:,} people, {s['completions']:,} completions "
          f"({s['online_pct']}% online), TLG hidden. Roster {len(D['roster']):,}; {len(D['upcoming'])} upcoming registrations.")
    with open(OUT, 'w', encoding='utf-8') as fh:
        fh.write(generate_html(D, meta))
    print(f'  Written → {os.path.basename(OUT)}')


if __name__ == '__main__':
    main()
