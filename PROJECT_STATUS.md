# Playbook Dashboard — Project Status

Last updated: 2026-07-31

---

## Live Dashboards

| Dashboard | File | Data Through | Notes |
|---|---|---|---|
| Analytics Hub Homepage | `index.html` | — | Links to all dashboards |
| Playbook Traffic | `playbook.html` | June 29, 2026 | |
| Healthcare Certification | `cert-healthcare.html` | July 28, 2026 | 45 enrolled, 0 certified; Reps/Managers chart split live |
| Public Sector Certification | `cert-publicsector.html` | July 31, 2026 | 118 active people, 74 completed (63%) — "Curricula" report format |
| Accelerate Onboarding | `onboarding.html` | July 30, 2026 | 37 learners; 18 reps with Closed Won deals |
| Accelerate Leaderboard | `leaderboard.html` | July 27, 2026 | Beta cohort filter live; 37 hires, 26 cohort deals |
| Layered Security Curriculum | `cert-layered-security.html` | July 29, 2026 | 33 learners, 20 complete (61%) — **PENDING:** 07.31 file removed, waiting on Resmie to confirm scope |

---

## Open Question — LS July 31 File (as of 2026-08-01)

The `Layered-Security-Curricula-Report-07.31.2026.xlsx` file was removed from `cert-data/`
because it showed 517 enrolled (vs 33 from July 29) — a 10x jump. Jason and Resmie are
investigating whether this is correct or an accidental over-broad pull.

**What the file actually contains:** 6,658 rows, ~505 unique people, all taking LSFDS curriculum,
across all 6 Direct Sales regions (Central, Southeast, West, Mid-Atlantic, Heartland, Northeast).
517 enrolled = 505 from July 31 + ~12 from July 29 not in July 31, minus exclusions
(Solutions Consultants, Commercial Print market, TLG).

**Clarification — "Direct Sales only" was a script interpretation, not a column in the file.**
There is no "Direct Sales" column. The original exclusion rules (remove Solutions Consultants +
Commercial Print) produced 33 people from the July 29 file. That framing was an interpretation.
The actual scope is determined by whoever pulls the LMS report.

**Two possible outcomes:**
1. **File is correct** → LS curriculum was expanded to all Direct Sales nationwide; put the file
   back, dashboard shows ~517 enrolled, 29 complete. No code changes needed.
2. **File was over-broad** → Resmie re-pulls with the right scope and sends a new file; drop it
   in `cert-data/` and push. Dashboard regenerates automatically with `LSC_IAPO` already wired.

**Next step:** Resmie confirms scope → act on outcome 1 or 2 above.

---

## Recent Changes (2026-08-01)

### HC Cert — Two more hardcoded denominators fixed after adding LSC_IAPO
- **LS pill on roster cards:** left-column pill was showing "Layered Sec X/11" → fixed to "/12"
- **overallPct calculation:** `round(overall_done / 21 * 100)` → `/ 22` — people who finished
  all 22 courses (12 LS + 10 HCF) were showing 105% instead of 100%
- **Detail panel label:** "X of 21 courses complete" → "X of 22 courses complete"
- **Pattern note:** when adding a course to LS_ORDER, check ALL three hardcoded denominators:
  (1) left-column pill label `/N`, (2) Python `overallPct` divisor, (3) detail panel label "of N courses"

---

## Recent Changes (2026-07-31 — session 2)

### PS Cert — Curriculum items breakdown in roster detail panel
- Clicking any person now shows all 3 curriculum items with completion status and date
- Pending items shown as "Not yet completed" (Jason confirmed: include incomplete items too)
- Same pattern as Healthcare's per-course checklist in the detail panel
- Script: `update_cert_dashboard.py` → `load_rows_publicsector` now stores `p['items']`

### LS + HC Cert — New course added: LSC_IAPO
- **New course:** "Layered Security Certification Introduction and Program Overview" (ID: `LSC_IAPO`)
- Added as first item in `ITEM_ORDER` in `update_layered_security_dashboard.py`
- Added as first item in `LS_ORDER` in `update_cert_dashboard.py`
- Course count updated 11 → 12 in tooltips and roster display in both scripts
- **Warning mechanism added:** both scripts now print a WARNING line when a file contains an
  item ID not in the order list — catches new courses automatically on the next data drop

### LS — Incorrect-scope 07.31 file removed
- `Layered-Security-Curricula-Report-07.31.2026.xlsx` had 505 people (company-wide pull)
  instead of the expected ~33 Direct Sales learners — removed from `cert-data/`
- Dashboard reverted to 07.29 data (33 learners, 20 complete)
- **Root cause of silent drop:** both LS and HC scripts had hardcoded `ITEM_ORDER`/`LS_ORDER`
  lists; any item ID not in the list was silently skipped with no warning
- **Next step:** Resmie to send a correctly-scoped 07.31 LS file (Direct Sales only); when
  dropped in `cert-data/`, `LSC_IAPO` will appear automatically — script is already ready

---

## Recent Changes (2026-07-31)

### Public Sector Cert — New data file + loader rewrite + terminology fix

- **New data file:** `cert-data/PublicSector-Certification-Curricula-Report-07.31.2026.xlsx`
  — new "Curricula" format (one row per curriculum item per person, not one row per person)
- **118 people, 74 completed (63%)** — up from 33 people in the old FY26 file; old file was a
  completely different format/scope
- **120 → 118 fix:** old FY26 file was being accumulated alongside new file, dragging in
  Braden McCargar and Stephen Langdon (inactive/departed — they dropped off the new report
  because it only includes active users). PS now uses latest-file-only (not accumulation).
- **Loader rewritten:** `load_rows_publicsector` now groups items by person internally,
  returns 118 unique rows directly. No downstream dedup needed (though it's harmless).
- **Completions Over Time chart fixed:** col 21 in new format is "Days Remaining" (not a
  cert date). Now uses max Item Completion Date (col 27) across all item rows per person as
  the completion date for `CertDate`/`CertQtr`. Trend chart now has real quarterly data.
- **Terminology:** "Certified/Certification" → "Completed/Completion" throughout the entire
  PS dashboard (stat cards, charts, filter dropdown, roster, exports, tooltips, print reports).
  Healthcare and Layered Security dashboards unchanged — those ARE certifications.
- **`extract_file_date` fix:** cert script was only matching `YYYY-MM` filenames; new PS file
  uses `MM.DD.YYYY`. Same class of bug as the onboarding sort fix from 2026-07-27.

### PS column map (new Curricula format — verified 2026-07-31)
```
col 0: User            col 14: Hire Date             col 22: Item ID
col 2: First Name      col 16: Parent Curriculum ID   col 26: Item Title
col 3: Last Name       col 17: Curriculum ID           col 27: Item Completion Date ← CertDate proxy
col 4: Email Address   col 18: Curriculum Title        col 28: Item Completion Status ID
col 5: Job Title       col 19: Curriculum Complete     col 29: Item Completion Status Description
col 6: Region          col 20: Curriculum Assignment Date
col 7: Market          col 21: Days Remaining          ← NOT a cert date (was PS_COL_CERT_DATE=21)
col 9: ManagerFirstName
col 10: ManagerLastName
col 11: SUPEMAILADDR (Manager email)
col 12: Manager JobTitle
```

---

## Recent Changes (2026-07-30)

### Onboarding Dashboard — v4 changes (ported from Resmie)
- **New LMS file:** `Accelerate-Curriculum-Report-07.30.2026.xlsx` (37 learners — Jonathan Griffith + Laud Vidal no longer in export, likely graduated out)
- **New Salesforce file:** `onboarding-data/New-Opportunities-Report-07.30.2026.xlsx` — replaces old leaderboard-data approach entirely. 18 reps with Closed Won deals (up from 2). `Age` field used directly as days-to-close.
- **Schedule-aware Expected %:** Was linear (days elapsed ÷ 35). Now per-curriculum: GS/SW/CP day 1–7, Prospecting 1–14, Sales Skills 8–21, Pipeline Mgmt 22–35. Item-count weighted. Flows through table column, sorting, modal metric grid, and schedule table.
- **Overdue severity badge:** Progress column now shows "Overdue · 14d" inline. Amber at 10+ days, red at 20+ days, based on worst single curriculum. Hover names which curriculum is driving it.
- **Modal expanded:** Risk banner (green/amber/red + recommended next step), metric grid (Overall %, Pacing Target, Curricula Done, Est. Completion), per-curriculum schedule table (Due By / Should Be At / Actually At). Modal widened 760→860px.
- **Column reorder:** Learner → Actual % → Expected % → Gap → Progress → Curricula → Days to Close
- **Averages row removed** from bottom of heatmap table
- **Amber CSS variable centralized:** `--amber` is now theme-aware; dark=#eab308, light=#92400e
- **`_date()` extended:** Now handles integer Excel serial dates (days since 1899-12-30) — new LMS files return integers for date cells with `openpyxl read_only=True`
- **Days to First Sale stat card fixed:** Now respects active filters (was averaging all reps regardless of filter state)
- Cleaned up: deleted `onboarding_dashboard_v4.html` and `dashboard_changelog.md` after porting

### Onboarding — LMS loader scoped to correct files (2026-07-30)
- **Bug:** `load_lms()` was loading every `.xlsx` in `onboarding-data/`. When `New-Opportunities-Report-07.30.2026.xlsx` was added to that folder today, the LMS loader accidentally read it as an LMS file, injecting Salesforce job titles as fake learner records (Dir IP Packaging & Label, Lead Customer Success Manager, etc.).
- **Fix:** Scoped the glob to `Accelerate-Curriculum-Report-*.xlsx` only. The Salesforce loader already uses its own specific pattern and was unaffected. No corrected LMS report needed — the 07.30 file was fine all along.
- **Rule:** Any new file type dropped into `onboarding-data/` is safe — the LMS loader ignores anything that isn't an LMS report.

### Onboarding — LMS loader scoped to correct files (2026-07-30)
- **Bug:** `load_lms()` was loading every `.xlsx` in `onboarding-data/`. When `New-Opportunities-Report-07.30.2026.xlsx` was added to that folder today, the LMS loader accidentally read it as an LMS file, injecting Salesforce job titles as fake learner records (Dir IP Packaging & Label, Lead Customer Success Manager, etc.).
- **Fix:** Scoped the glob to `Accelerate-Curriculum-Report-*.xlsx` only. The Salesforce loader already uses its own specific pattern and was unaffected.
- **Rule:** Any new file type dropped into `onboarding-data/` is safe — the LMS loader ignores anything that isn't an LMS report.
- **Verified clean (2026-07-30):** All 37 learner records sanity-checked after the fix — per-curriculum percentages match item counts, no missing assign dates, no data gaps.

### Onboarding — extract_date sort bug fixed (2026-07-30)
- **Bug:** `extract_date()` regex looked for `YYYY-MM` format but filenames use `MM.DD.YYYY`. Every file returned the same sort key `'0000-00'`, making `files[-1]` non-deterministic — Actions could pick any file, not necessarily the latest one. Same class of bug as the CI file sort fix on 2026-07-27.
- **Fix:** Updated regex to parse `MM.DD.YYYY` first (converting to `YYYY-MM-DD` for correct sort), with `YYYY-MM` as a fallback. Now reliably picks `07.30.2026` as the latest file.

---

## Earlier Changes (2026-07-29)

### Layered Security Curriculum Dashboard — NEW
- New script: `update_layered_security_dashboard.py` reads `cert-data/Layered-Security-Curricula-Report-*.xlsx`
- Generates `cert-layered-security.html` — 33 Direct Sales learners (excludes Solutions Consultants + Commercial Print market)
- 11-module curriculum (`LSFDS`): 20 complete (61%), 13 in progress or not started
- Features: stat cards, completion pipeline + market charts, individual/manager roster views, course-level checklist, PDF + Excel exports, sales certification import
- Actions workflow updated — triggers on `update_layered_security_dashboard.py` + `cert-data/*.xlsx`
- Homepage card added showing completion stats

### Onboarding — Progress table simplified
- **Playbook dot removed** from learner name cells — engagement info still in the individual modal card
- **Legend simplified** — removed the "Playbook ●" section; Gap + Curricula ■ keys remain
- **Progress column** — shows only the status badge (On Track / Overdue / Completed); sub-detail lines (X courses past due, Xd until deadline, Behind pace, Also lagging) removed from table view

---

## Earlier Changes (2026-07-28)

### Healthcare Cert — Certifications Over Time chart split (Reps vs Managers)
- Ported from Resmie's prototype (`cert-healthcare-07-28-2026.html`, now deleted)
- `isManager(p)` — job title contains "Director of Sales" or "Vice President" → Manager; everyone else → Rep
- Chart 2 now renders as a **stacked bar** with two datasets: Reps (purple `#8b5cf6`) and Managers (blue)
- Legend visible at bottom; tooltip shows per-segment count + "Total:" line on hover
- Currently all 4 certified people are Managers (Directors of Sales), so only the Manager segment shows — Rep segment appears once reps start certifying
- New data files added: `Healthcare-Certification-Report-07.28.2026.xlsx` + Foundations Curricula variant

### Onboarding Dashboard — Canadian Learner Support
- **New data file:** `Accelerate-Curriculum-Report-07.28.2026.xlsx` — first file with Canadian learners (39 total: 37 US + 2 Canadian: Jonathan Griffith/Ontario, Laud Vidal/BC)
- **Bug fixed:** Canadian learners were showing 0% on all curricula because the script only knew US curriculum IDs (`ACCELERATE_GS` etc.). Fixed by detecting `Parent Curriculum ID = 'ACCELERATE_BCA'` and dynamically appending `_BCA` suffix when reading curriculum rows. Data is stored under US keys so the JS heatmap is unchanged.
- **New column constant:** `COL_PARENT_CURRIC = 16`
- **Location filter added:** All / 🇺🇸 US / 🇨🇦 Canada button group in the filter bar
- **Flag icon:** 🇨🇦 appears next to Canadian learner names in the progress table
- **Course differences confirmed:** CA has CASL (GS), TRUEBLUE (CP instead of BlueIrisIQ), KMpriceHUB + FINTRAC (SS); Prospecting is identical between US and CA

---

## Earlier Changes (2026-07-27 evening)

### Leaderboard updated for new Salesforce file format
- **New file names:** `New-Opportunities-Report-07.27.2026.xlsx` (Closed Won) and `New-Opportunity-History-Report-07.27.2026.xlsx` (Stage History) — previously `report*.xls` HTML-disguised exports
- **File detection:** switched from content-sniffing (reading first 2000 bytes) to filename-pattern matching. Falls back to old `.xls` content sniff if new-format files not present.
- **New parser:** `_parse_xlsx_sf()` uses openpyxl for real Excel files. `_parse_sf_file()` routes by extension.
- **LMS column shift:** new `Accelerate-Curriculum-Report-07.27.2026.xlsx` dropped the duplicate `Email Address` col at position 0, shifting all columns left by 1. Updated constants: COL_FIRST=2, COL_LAST=3, COL_EMAIL=4, COL_JOBTITLE=5, COL_REGION=6, COL_MARKET=7, COL_BRANCH=8, COL_HIRE_DATE=14, COL_CURRIC_COMPLETE=19, COL_ASSIGN_DATE=20.
- **Result:** 37 hires, cohort start 2026-06-04, 26 cohort deals, 37 verification entries.
- **Health check:** if a future run shows `cohort start None` or hires < 30, the LMS columns shifted again — re-verify against the new file.

---

## Earlier Changes (2026-07-27 morning)

### Healthcare Cert dashboard updated to 07.27 data
- New files: `cert-data/Healthcare-Certification-Report-07.27.2026.xlsx` and `cert-data/Healthcare-Certification-Foundations-Curricula-Report-07.27.2026.xlsx`
- Output: 45 people enrolled, 0 certified, 44 in progress, 1 not started

### CI file sort bug fixed (critical)
- **Bug:** `sorted(glob.glob(...), key=os.path.getmtime)` is non-deterministic in GitHub Actions — fresh checkout gives all files the same mtime, so the "latest" file was chosen randomly. This caused Actions to overwrite the July 27 HTML with July 20 data.
- **Fix:** Removed `key=os.path.getmtime` in `update_cert_dashboard.py`, `generate_homepage.py`, and `update_onboarding_dashboard.py`. Alphabetical sort by filename is correct since filenames contain `MM.DD.YYYY`.
- **Rule going forward:** Never use `os.path.getmtime` as a sort key for file selection in any script. Only use it as a date-label fallback.

### Workflow rule established
- **Never commit generated HTML manually.** GitHub Actions owns `cert-healthcare.html`, `cert-publicsector.html`, `index.html`, `onboarding.html`, `leaderboard.html`.
- Correct pattern: push data files (Excel) or script changes → Actions regenerates HTML → `git pull`.
- Committing HTML manually + Actions also committing = two writers, one overwrites the other.

---

## Previous session changes (2026-07-22 to 2026-07-24)

### Leaderboard — Beta Cohort filter (ported from Resmie)
- "Hide Beta Cohort" toggle button in header
- `BETA_NAMES` Set of 27 hardcoded names — immune to LMS `assignDate` resets
- `renderAll()` pattern: wraps all augment + stat + render logic; all section renders called through it
- `ALL_DEALS` / `ALL_HIRES` / `ALL_VERIFICATION` = source arrays; `DEALS` / `HIRES` / `VERIFICATION` = filtered working vars

### Onboarding — Test Group toggle fixed
- **Bug:** Filter used `p.assignDate === TEST_GROUP_DATE`. LMS resets `Curriculum Assignment Date` to report generation date on each pull — 23 of 27 June 4 cohort members showed `2026-07-10` instead of `2026-06-04`.
- **Fix:** Replaced `TEST_GROUP_DATE` constant with `TEST_GROUP_NAMES` Set (same 27 names as leaderboard `BETA_NAMES`). Filter now uses `TEST_GROUP_NAMES.has(p.name)`.
- **Rule:** Never use `assignDate` to identify the Test Group / Beta Cohort. Always use the hardcoded name Set.

---

## Key Rules (don't break these)

1. **Never edit generated HTML directly** — always edit the Python script and regenerate
2. **Never commit generated HTML manually** — push data files, let Actions handle the HTML
3. **Never use `assignDate` to filter the June 4 test cohort** — use `TEST_GROUP_NAMES` / `BETA_NAMES`
4. **Never use `os.path.getmtime` for file selection sort** — use alphabetical filename sort
5. **Git conflict fix:** `git pull --no-rebase -X ours && git push` when Actions auto-commit diverges

---

## Pending

- HC sub-certifications: still not live in LMS. Restore from `_snapshots/update_cert_dashboard-full-subcurricula-2026-06-29.py` when they go live.
