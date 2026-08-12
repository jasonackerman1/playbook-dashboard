# Playbook Dashboard — Project Status

Last updated: 2026-08-12

---

## Live Dashboards

| Dashboard | File | Data Through | Notes |
|---|---|---|---|
| Analytics Hub Homepage | `index.html` | — | Links to all dashboards |
| Playbook Traffic | `playbook.html` | June 29, 2026 | |
| Healthcare Certification | `cert-healthcare.html` | August 10, 2026 | 65 people, 6 certified, 59 in progress |
| Public Sector Certification | `cert-publicsector.html` | July 31, 2026 | 118 active people, 74 completed (63%) — "Curricula" report format |
| Accelerate Onboarding | `onboarding.html` | August 11, 2026 | 47 learners; 20 reps with Closed Won deals; Coming Soon courses now counted toward % (see below); Days to Close now hire-date-based |
| Accelerate Leaderboard | `leaderboard.html` | July 27, 2026 | Beta cohort filter live; 37 hires, 26 cohort deals |
| Layered Security Certification | `cert-layered-security.html` | August 12, 2026 | 503 learners, 41 complete; rebuilt to merge Resmie's prototype (see below) |

---

## Recent Changes (2026-08-12) — Layered Security dashboard merge

Resmie hand-built a prototype (`Layered_Security_Certification_Dashboard.html` +
`LS_Dashboard_Documentation.html`) with a real bug fix and new features. Ported into
`update_layered_security_dashboard.py` rather than adopting her static file — she'd iterated
forward from the exact same person-object schema the script already emits, and this repo's rule
is edit-the-script-not-the-generated-HTML (same convention as Onboarding). Both her files deleted
after porting.

**3-file cadence, no more accumulation.** Old dated Layered Security Curricula files (07.29, 08.01,
08.10) deleted; going forward it's "drop 3 new dated files, delete the old ones," not multi-file
accumulation like Healthcare. Three files now: the Curricula Report (same 31-col schema, unchanged)
plus two new ones — a Certification Report (Sales Certification course, scoped to already-certified
people only — confirmed by Jason, so "any row present = certified" is the correct permanent rule,
currently empty because nobody's certified yet) and an OPS FY26 Salesforce opportunity export
(drives Closed Won deal tracking, $5,000+ threshold — **currently has no Amount/dollar column at
all**, flagged to Resmie separately; the loader prints a non-fatal `WARNING` and Closed Won just
stays empty until one appears).

**Real bug fixed: stale `Curriculum Complete` LMS flag.** Someone flagged "Yes" back when the
curriculum had 11 modules (before the 12th, `LSC_IAPO`, was added) stayed "Yes" while stuck at
11/12. Both `load_ls_data()` (Python) and `personStatus()` (JS) now derive completion from actual
module count instead of trusting the flag — fixed in every consumer (stat cards, roster, By-Manager
grouping, all 6 export functions) and in `generate_homepage.py`'s homepage card too, so nothing can
drift out of sync. Verified zero literal `p.Complete === "Yes"` comparisons remain in the generated
output.

**New Closed Won tracking, fully build-time-baked (not client-imported).** The old "Import Sales
Certification" browser-upload button is gone. `SALES_CERT`/`SALES_DEALS` are now baked into the
generated HTML the same way `PEOPLE` already is, via two new Python loaders
(`load_sales_cert()`/`load_sales_deals()`) that detect columns by name-substring match (robust to
Salesforce export column reordering).

**Roster redesigned from a card list to a table** (Learner · Manager · Layered Security · Overall %
· Closed Won), matching Resmie's shipped design. Stat cards reordered (Certified moved from 3rd to
last position). New requirements banner spelling out all 4 certification criteria. Detail panel
gained a Closed Won field and a simplified binary certification badge. The existing "Past Due"
indicator was deliberately retained — Resmie's prototype had silently dropped it with no stated
reason, and removing it wasn't requested.

Verified locally: script runs clean (`503 learners, 41 complete`), homepage card matches exactly
(`41 of 503 complete`), generated JS passes `node --check`, git status shows the expected file
adds/deletes. **Pushed live** (commit `65a53d4`) — merged cleanly with unrelated leaderboard-data
commits already on `origin/main` via `git pull --no-rebase -X ours`, no conflicts.

**Still open: waiting on Resmie's updated OPS FY26 file with a real Amount column.** Jason confirmed
a new export is coming (unclear why the original one omitted Amount — likely just a Salesforce
report-column selection, not a data/permissions issue) but it hadn't landed in `cert-data/` yet as
of this checkpoint. No code changes needed when it arrives — `load_sales_deals()` finds the Amount
column by name, not position.

**Push-time build hiccup, resolved same session — not a real regression.** The push triggered both
`Update Cert Dashboards` and `Update Onboarding Dashboard` simultaneously (both matched the merge
commit's changed paths); Onboarding's auto-commit landed first, so Cert Dashboards' own auto-commit
got rejected on push (`! [rejected] main -> main (fetch first)`) — same class of race as the documented
Git conflict pattern below. Investigated the failed run's logs directly via the GitHub API: the lost
commit only touched `index.html`, and diffing it against a fresh local regeneration showed it was
just the homepage's Leaderboard section flipping between real numbers and placeholder dashes —
traced to `generate_homepage.py`'s legacy `.xls` HTMLParser fallback behaving differently under local
Python 3.9.6 vs GitHub Actions' Python 3.11.15 (`subclasses of ParserBase must override error()`).
Confirmed via `git status` that `cert-healthcare.html`/`cert-publicsector.html`/
`cert-layered-security.html` had zero diff on fresh regeneration — nothing from the Layered Security
work was actually lost. Manually re-triggered the workflow via `workflow_dispatch`; it completed
`success` with nothing new to commit, confirming Actions' own environment already had the correct
state. **Pre-existing, out-of-scope issue surfaced but not fixed:** the leaderboard-stats reader in
`generate_homepage.py` is environment-sensitive due to the legacy `.xls` parser — worth a real fix
next time the homepage is touched, but not addressed here since it's unrelated to Layered Security
and didn't actually cause data loss.

Full detail in project memory: `layered_security_dashboard.md`.

---

## Recent Changes (2026-08-11) — Onboarding: Coming Soon courses + Days to Close redefinition

**Coming Soon courses now count toward completion %.** Previously any item with "coming soon"
in its title was skipped entirely — excluded from every curriculum's total/done counts. Jason
asked to bring them back into the % math (so a curriculum containing an unlaunched course won't
show artificially inflated to 100% for people who obviously can't complete it), while making
sure they can never be flagged Overdue since they have no real due date. Fix: items are no
longer skipped, but each Coming Soon item's `req` (LMS "Item Required Date") is always forced to
`None` in `update_onboarding_dashboard.py`, regardless of what the LMS column contains — this
keeps them permanently excluded from the deadline engine (`curricDaysLeft`/`overdueItemsCount`)
even after the LMS eventually starts populating real due dates for other courses. Verified: all
188 Coming Soon item entries across the current roster show `done=false, req=null`. Affects 5
courses as of the 08.10/08.11 data: How to Prepare for Effective Account Reviews, FINTRAC,
Introducing AllCovered/IT Weapons, Commission Confidence, KM Premier Finance Leasing
Fundamentals.

**Days to Close redefined: hire date → first Closed Won deal.** Was previously Salesforce's own
`Age` field (Close Date − opportunity Created Date — i.e. how long the *deal* took, not the
*rep*). Jason wanted it measured from the rep's actual hire date to their earliest Closed Won
deal instead — a truer "how fast did this person land their first sale after joining" metric.
Uses `hireDate` (LMS column), not `assignDate` (Accelerate program start) — those two diverge
significantly for this cohort since several reps were hired months before the program launched.
Shows a dash when there's no hire date on file, the deal predates the hire date (bad-data guard),
or no Closed Won deal exists yet. Both info-tooltip strings updated to match. Verified against
sample reps, e.g. Paolo Castellon: hired 2025-11-03, first Closed Won 2026-02-17 → 106 days.

Committed as `3ad4861`, pushed, confirmed the `Update Onboarding Dashboard` Actions run completed
`success`. Full detail in project memory: `onboarding_dashboard.md`.

---

## Recent Changes (2026-08-11) — same-date data file mix-up, fixed

**What happened:** Jason pushed a Healthcare data update (2 files, both dated 8.10.2026). The
GitHub Actions build failed (`Update Cert Dashboards`, exit code 1) at the Layered Security
step. Root cause: the updated (fuller, 65-person) `Healthcare-Certification-Report` file got
saved into the `Layered-Security-Curricula-Report-08.10.2026.xlsx` slot instead of its own —
same-date filenames made the destination easy to mix up. This overwrote the real 31-column
Layered Security item-level data (820KB → 21KB), so `update_layered_security_dashboard.py`
crashed with `IndexError: tuple index out of range` on `row[COL_ITEM_ID]`. The Healthcare
script itself ran fine, just against the stale 50-person version of its own file.

**Diagnosis:** Used the GitHub REST API directly (no `gh` CLI installed) with the token from
`git credential-osxkeychain get` to pull the failed run's job list and logs, which pinpointed
the exact failing step/line. Confirmed the swap by diffing the current file's columns against
the previous commit's version, and by checking that the misplaced 65-row Healthcare file was a
strict superset of the stale 50-row one (same underlying report, more complete).

**Fix (commit `3f6037f`):** restored the real Layered Security file from the commit before the
mix-up, moved the 65-row Healthcare data into `Healthcare-Certification-Report-08.10.2026.xlsx`,
verified both `update_cert_dashboard.py` and `update_layered_security_dashboard.py` run clean
locally, then pushed on Jason's explicit "push fix" instruction. Confirmed the resulting GitHub
Actions run completed with `success` via the API.

**Result:** Healthcare — 65 people, 6 certified (up from 45/0). Layered Security — 526
learners, 41 complete (up from 517/30).

Full incident writeup in project memory: `cert_data_pipeline.md`.

---

## Open Question — LS Audience Scope (as of 2026-08-01)

`Layered-Security-Curricula-Report-08.01.2026.xlsx` is now live in `cert-data/`. Dashboard
shows 517 learners, 30 complete. This is Resmie's replacement for the 07.31 file.

**Pending:** Resmie is still confirming whether 517 is the correct audience (all Direct Sales
nationwide) or whether the pull was too broad. If she re-pulls with a tighter scope, drop
the new file in `cert-data/` and push — script is ready, no code changes needed either way.

**Background:** "Direct Sales only" was never a column in the file. The original exclusion rules
(remove Solutions Consultants + Commercial Print) produced 33 people from the July 29 file.
Scope is entirely set by whoever pulls the LMS report.

---

## Recent Changes (2026-08-10) — session 2

### Layered Security — Days Remaining removed from roster detail cards
- Removed "Days Remaining" field from individual roster detail panels.
- `DaysRemaining` value still baked into the data (used by `isPastDue()` for the roster badge logic) — only the displayed field was removed.

### Layered Security — JS crash fixed (Past Due removal follow-up)
- `sel("s-overdue").textContent` was still called after the Past Due stat card was removed, crashing all stat rendering and leaving every card showing "—".
- Removed the dead assignment and the now-unused `overdue` variable.

### Layered Security — Past Due stat card removed
- Removed the "Past Due" stat card from the top of `cert-layered-security.html`.
- Remaining cards: Total Enrolled, Curriculum Complete, Certified, In Progress, Not Started, Completion Rate.

---

## Recent Changes (2026-08-10)

### Leaderboard — "Hide Beta Cohort" renamed to "Hide Test Group"
- Button label updated to match terminology used on the Accelerate Onboarding dashboard.
- Same 27 people, same filter logic — label only. `BETA_NAMES` / `hideBeta` variable names unchanged internally.

### Leaderboard — Part B added to 45-day window rule
- **New rule:** Reps roll off the Closed-Won Leaderboard once their 45-day program window expires
  (today's date > assignDate + 45 days), even if their deal closed within the window.
- **Filter change:** `leaderboardRows` now also requires `hireMap[d.name].eligible` (in-window today).
- **Effect today:** 6 expired-window reps removed from the leaderboard display:
  Morgan Bruno (day 48), Daniel Zepeda (day 67), Kelli Sorrentino (day 63),
  Michael Shields (day 67), Paolo Castellon (day 67), Randahl Bradley (day 67).
  The other 41 cohort members with active windows are unaffected.
- **Investigation note:** Morgan Bruno's deal (NORTHLIGHT THEATRE, $5,900, day 24) technically
  passed all prior rules — Morgan personally progressed through SQ and Engage — but the deal was
  Created By Adit Thakur (not Morgan). Jason flagged it; window-expiry rule removes it cleanly
  without adding a "self-created" requirement.
- Eligibility note and section description updated to mention roll-off.

---

## Recent Changes (2026-08-01) — session 3

### All workflows confirmed working
- New playbook traffic data file dropped; both `update-dashboard.yml` (playbook.html) and
  `update-onboarding.yml` (onboarding.html) triggered and ran successfully.
- Onboarding workflow already had `data/playbook-monthly-*.xlsx` as a trigger — no manual
  re-trigger needed when new playbook data is dropped.
- Workflows had some failures during the session but resolved; all dashboards are current.

---

## Recent Changes (2026-08-01) — session 2

### LS + HC Cert — Raw item ID showing instead of course name (fixed)
- **Bug:** When a person's LMS data had no row for `LSC_IAPO` (not yet started), the fallback
  was the raw item ID string — showing "LSC_IAPO" in the roster detail panel instead of the
  full course name.
- **Fix:** Added `ITEM_TITLES` dict in `update_layered_security_dashboard.py` and `LS_TITLES`
  dict in `update_cert_dashboard.py` — 12 item IDs → display names. Changed fallback:
  `title = iid` → `title = ITEM_TITLES.get(iid, iid)` / `raw_item.get('title') or LS_TITLES.get(iid, iid)`.
- **Pattern:** Any time a new course is added to `ITEM_ORDER`/`LS_ORDER`, also add its display
  name to the titles dict. This prevents raw IDs from ever appearing in the UI.

### LS — 07.31 file replaced with 08.01 file from Resmie
- Removed `Layered-Security-Curricula-Report-07.31.2026.xlsx`
- Added `Layered-Security-Curricula-Report-08.01.2026.xlsx` — 517 learners, 30 complete
- Scope still pending confirmation from Resmie (see Open Question above)

---

## Recent Changes (2026-08-01) — session 1

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
