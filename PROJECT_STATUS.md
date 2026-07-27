# Playbook Dashboard — Project Status

Last updated: 2026-07-27

---

## Live Dashboards

| Dashboard | File | Data Through | Notes |
|---|---|---|---|
| Analytics Hub Homepage | `index.html` | — | Links to all dashboards |
| Playbook Traffic | `playbook.html` | June 29, 2026 | |
| Healthcare Certification | `cert-healthcare.html` | July 27, 2026 | 45 enrolled, 0 certified |
| Public Sector Certification | `cert-publicsector.html` | May 2026 | 33 people, 6 certified |
| Accelerate Onboarding | `onboarding.html` | July 22, 2026 | 36 learners |
| Accelerate Leaderboard | `leaderboard.html` | July 22, 2026 | Beta cohort filter live |

---

## Recent Changes (this session — 2026-07-27)

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
