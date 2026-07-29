# Playbook Dashboard — Project Status

Last updated: 2026-07-29

---

## Live Dashboards

| Dashboard | File | Data Through | Notes |
|---|---|---|---|
| Analytics Hub Homepage | `index.html` | — | Links to all dashboards |
| Playbook Traffic | `playbook.html` | June 29, 2026 | |
| Healthcare Certification | `cert-healthcare.html` | July 28, 2026 | 45 enrolled, 0 certified; Reps/Managers chart split live |
| Public Sector Certification | `cert-publicsector.html` | May 2026 | 33 people, 6 certified |
| Accelerate Onboarding | `onboarding.html` | July 28, 2026 | 39 learners (37 US + 2 Canadian) |
| Accelerate Leaderboard | `leaderboard.html` | July 27, 2026 | Beta cohort filter live; 37 hires, 26 cohort deals |

---

## Recent Changes (2026-07-29)

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

- Onboarding 07.27 data: Resmie said 45 people enrolled but `onboarding-data/` only goes to 07.22 (36 learners). Waiting on `Accelerate-Curriculum-Report-07.27.2026.xlsx` from Resmie.
- HC sub-certifications: still not live in LMS. Restore from `_snapshots/update_cert_dashboard-full-subcurricula-2026-06-29.py` when they go live.
