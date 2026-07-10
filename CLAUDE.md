# Playbook Dashboard — Claude Instructions

## DO NOT PUSH
Do not push to GitHub unless Jason explicitly says to. All development is tested locally first.

---

## Dashboards

| File | Script | Description |
|---|---|---|
| `index.html` | `generate_homepage.py` | Analytical Data Hub homepage |
| `playbook.html` | `update_dashboard.py` | Playbook Traffic Dashboard |
| `cert-healthcare.html` | `update_cert_dashboard.py` | Healthcare Certification Dashboard |
| `cert-publicsector.html` | `update_cert_dashboard.py` | Public Sector Certification Dashboard |
| `onboarding.html` | `update_onboarding_dashboard.py` | Accelerate Onboarding |
| `leaderboard.html` | `update_leaderboard_dashboard.py` | Accelerate Leaderboard |

**Never edit generated HTML files directly** — always edit the Python script and regenerate.

**Exception — Resmie:** Resmie sometimes edits `onboarding.html` directly to prototype JS changes. When she does, `diff` her version against what the script generates, identify the JS changes, then port them into `update_onboarding_dashboard.py`. Her edited files appear as `onboarding (2).html` (or similar) in the repo root — these are safe to delete after porting.

---

## Snapshots — 2026-06-29

Full-state backup of all dashboards and scripts saved to `_snapshots/` before simplifying the Healthcare cert dashboard.

| Snapshot file | What it contains |
|---|---|
| `cert-healthcare-full-subcurricula-2026-06-29.html` | HC dashboard with sub-cert columns (Layered Security, Healthcare, Ambulatory, Extended Care) |
| `cert-publicsector-2026-06-29.html` | PS dashboard at this date |
| `index-2026-06-29.html` | Playbook Traffic Dashboard at this date |
| `onboarding-2026-06-29.html` | Accelerate Onboarding at this date |
| `update_cert_dashboard-full-subcurricula-2026-06-29.py` | Python source for full HC dashboard — **restore point for sub-curricula** |
| `update_dashboard-2026-06-29.py` | Playbook script at this date |
| `update_onboarding_dashboard-2026-06-29.py` | Onboarding script at this date |

**Why:** HC sub-certifications are not yet live in the LMS. Shipping a simpler HC dashboard now; restoring the full sub-curricula version once they go live. When ready, start from `_snapshots/update_cert_dashboard-full-subcurricula-2026-06-29.py`.

---

## Navigation

All dashboards (Playbook, HC Cert, PS Cert, Onboarding, Leaderboard) use a **hidden triple-click Easter egg** on the `<h1>` to return to `index.html`. No visible hamburger nav.

The homepage (`index.html`) has a card for each dashboard — that's the primary navigation entry point. Always add a new card in `generate_homepage.py` when adding a new dashboard.

---

## KM Academy Logo SVGs

Both SVG files are local to the repo root. Do NOT use the CDN URL for `KMA-wht.svg` — a local copy exists here.

| File | Theme | `cls-2` fill |
|---|---|---|
| `KMA-wht.svg` | Dark backgrounds | `#fff` |
| `KMA-drk.svg` | Light backgrounds | `#1a1e2e` |

**"POWERED BY" text is baked into both SVGs** at SVG coordinates `x="93" y="9.5" font-size="10"` — centered in the negative space between the hat icon (right edge x≈35) and "KONICA MINOLTA" small text (left edge x≈151). Do not add a "Powered by" wrapper div in any HTML — it's part of the SVG.

All Python scripts reference:
- Dark mode logo: `<img src="KMA-wht.svg" class="kma-logo kma-logo-dark">`
- Light mode logo: `<img src="KMA-drk.svg" class="kma-logo kma-logo-light">`

---

## Onboarding Script — Developer Notes

**`_date()` helper (critical):** LMS exports dates in two formats — datetime objects (openpyxl) and `"M/D/YYYY timezone"` strings (e.g. `"6/11/2026 US/Alaska"`). `new Date('M/D/YYYY US/Alaska')` returns Invalid Date in browsers → NaN in all math. Always run LMS cell values through `_date()` before embedding in JSON. It normalizes to `YYYY-MM-DD`.

**`COL_ITEM_REQ = 30`:** "Item Required Date" column — currently blank/space in the LMS export (not yet populated). The per-curriculum deadline engine is built and waiting; it activates automatically once the LMS starts populating this column.

**JS onclick escaping inside f-string:** Single quotes inside single-quoted JS strings break the template. For `onclick="showInfo(event,'key')"` patterns inside `hRow += '...'` strings, use `\\'` in Python so the output JS has `\'`:
```python
# CORRECT:
hRow += '...<span onclick="showInfo(event,\\'key\\')">?</span>...';
```
For user-supplied values (names, etc.) use `data-*` attributes and read them in a JS handler — never trust single-quote escaping for dynamic content.

**Current data file:** `onboarding-data/Accelerate-Curriculum-Report-07.07.2026.xlsx` (29 learners, 24 matched to playbook). Verify column positions against each new file from Resmie before regenerating.

**TLG always hidden:** `hideTLG = true` permanently. No toggle button — same behavior as healthcare cert dashboard. Do not add a Show/Hide TLG button.

**Top filter bar (updated 2026-07-08):** Market · Status · Sort · Reset — that's it. No "search name or manager" input (removed — the progress report has its own search box). The Sort dropdown syncs `tableSort` when changed, so it also re-sorts the progress report table. Options: Name A→Z · Completion High→Low · Completion Low→High · Most Urgent First.

**Progress report search (dual-mode, updated 2026-07-08):** The `table-search` input at the top of the progress report section is context-aware:
- **Individual view** — placeholder "Search name..." — filters learner rows by `data-name`
- **Manager view** — placeholder "Search manager..." — filters entire manager groups by `data-manager` (hides header + all learner rows for non-matching managers)
- Switching views clears the search and updates the placeholder. Manager group header rows carry `data-manager="lowercase name"` to enable this.

---

## Homepage Visual Effects (generate_homepage.py — added 2026-07-09)

Three animated effects are embedded in the generated `index.html` via `generate_homepage.py`:

1. **Particle network** — `<canvas id="particles-bg">` fixed behind all content. 65 dots, connect within 130px. Mode-aware colors (dark/light).
2. **Aurora header glow** — 3 blurred orbs (`filter:blur(70px); border-radius:50%`) inside `.aurora-wrap` div in `.header`. Animated with `@keyframes aurora-drift`.
3. **Animated stat counters** — `.stat-num` elements count up on DOMContentLoaded, 1400ms ease-out cubic easing.

**Dark mode aurora: DO NOT CHANGE** — Jason confirmed it's "perfect." Three-orb positions at top:-120px/left:1%, top:-85px/left:40%, top:-100px/right:3% are intentional.
**Light mode aurora:** Centered cluster, all three orbs at left:18–38%, top:-45 to -60px — glow sits behind the center logo, leaving title and buttons on clean white.

All JS/CSS is inside a Python f-string — use `{{` / `}}` for literal braces.

---

## Cert Dashboard — `generate_html_healthcare_v2` CSS (Fixed 2026-07-09)

`update_cert_dashboard.py` has two healthcare generator functions. The v2 function (currently used) has its own self-contained CSS block — it does NOT inherit from the v1 function.

**Every time CSS is added to the v1 function, check whether it's also needed in v2.** The kma-logo CSS was missing from v2, causing the logo to be invisible. Required in v2's CSS block:
```css
.header { display:grid; grid-template-columns:1fr auto 1fr; align-items:center; gap:12px; }
.header-center { display:flex; justify-content:center; align-items:center; }
.kma-logo { height:38px; width:auto; display:block; }
.kma-logo-light { display:none; }
.light-mode .kma-logo-dark { display:none; }
.light-mode .kma-logo-light { display:block; }
```

---

## Leaderboard Script — Developer Notes

**Window anchor (critical):** The 45-day window runs from `assignDate` (LMS col 21, Accelerate program assignment date), NOT `hireDate`. This cohort was hired months before the program launched — the two dates diverge significantly. For future cohorts they may align, but the logic must always use `assignDate`.

**Window timing:** Eligibility is checked at deal close date (`assignToCloseDays = closeDate - assignDate`), not today. This means a deal that closed on day 30 stays on the board even after the rep's window has since expired.

**Both HIRES and DEALS carry `assignDate`:** HIRES uses it to compute `daysSince` for the Window Tracker (today-based). DEALS uses it for `assignToCloseDays`.

**On Deck section:** Shows in-window reps (today check) with non-qualifying deals. "What's Missing" column — spell out all labels fully, no abbreviations. "SQ" → "Sales Qualified." Also shows curriculum-complete in-window reps with no deals yet.

**Window Tracker split:** Active reps (in window) in the main table. Expired reps in a collapsed toggle (`expiredTrackerWrap` div, hidden by default). Count badge shows active count only.

**History table column names:** "Program Start" (`assignDate`) and "Program Day" (`assignToCloseDays`) — not "Hire Date" / "Hire→Close."

**LMS column constants (0-based):**
```
COL_FIRST=3, COL_LAST=4, COL_EMAIL=5
COL_JOBTITLE=6, COL_REGION=7, COL_MARKET=8, COL_BRANCH=9
COL_HIRE_DATE=15, COL_CURRIC_COMPLETE=20, COL_ASSIGN_DATE=21
```
CAUTION: verify against each new file from Resmie before regenerating.

**`curriculumComplete = 'Yes'` is a multi-step final state:** Courses done → workshops attended → Boot Camp → capstone presentation to John Lechner with his sign-off. Reps finishing individual courses do NOT flip this flag. "On board = 0" in the homepage card while reps are actively enrolled is accurate and expected.

**`leaderboard_stats()` in `generate_homepage.py` uses `assignDate`** for the 45-day window check (fixed 2026-07-09). Before the fix, it was using `hireDate`, which made all 29 reps appear out-of-window since they were hired months before Accelerate launched.

**Git push conflict pattern:** GitHub Actions auto-commits `leaderboard.html` after each push, causing diverged branches. Fix: `git pull --no-rebase -X ours && git push`.

---

## Accelerate Playbook — Page & Course Map (confirmed 2026-07-09)

Used in `update_onboarding_dashboard.py` for two purposes: (1) timeline sorting — within same-day clusters, playbook visits appear immediately before courses on that page; (2) "In sequence" label — courses in prescribed curriculum order.

| Section | Page | Slug | Courses |
|---|---|---|---|
| Homepage | Home | `index` | — |
| Getting Started | Program Overview | `overview` | Register: Accelerate Advanced Skills Boot Camp |
| Getting Started | Welcome to Konica Minolta | `welcome` | — (doc downloads only) |
| Getting Started | Understanding Salesforce | `understandingsalesforce` | KM Sales Experience · Register: Salesforce Live Workshop |
| Sales Workflow | Sales Workflow | `salesworkflow` | Mastering Your Daily Sales Workflow |
| Core Portfolio | Core Portfolio | `coreportfolio` | Why Konica Minolta · bizhub One i Series · Layered Security: Introducing the Model · Layered Security: User Auth · Layered Security: Document Security · Introduction to Production Print · Introducing Blue Iris IQ · Introducing All Covered · Register: Managed IT New Hire Kickstart |
| Prospecting | Prospecting Skills | `prospectingskills` | Identifying Real Sales Opportunities · Lease Upgrade Sheet Pt 1 & 2 · High Payoff Activities · Intro to Microsoft Copilot · Copilot: Art of Prompting · Targeted Message: Foundations · Targeted Message: Persona Worksheet · Design Better Prospecting Calls (DPQ) · Build Better Voicemails · Register: Accelerate Advanced Skills Boot Camp · Register: Prospecting Live Workshop |
| Prospecting | Salesforce Prospecting | `salesforceprospecting` | Prospecting Foundations: Turning Skills into Actions |
| Sales Skills | Call Prep Essentials | `callprep` | From Prep to Performance · Discovery That Delivers · Presentation Best Practices · How to Prepare for Account Reviews (Coming Soon) |
| Sales Skills | Working with Numbers | `workingwithnumbers` | Introducing KM Premier Finance · Leasing Fundamentals · Sales Math: Numbers Don't Lie · Commission Confidence (Coming Soon) · Register: KM Premier Finance Leasing Workshop |
| Pipeline Mgmt | Moving Deals Forward | `movingdeals` | Managing and Moving Your Deal Forward |
| Pipeline Mgmt | Pipeline Ownership | `pipelineownership` | How to Build Accurate Forecasts and Strong Pipelines |
| — | Resources | `resources` | — |
| — | Managers | `managers` | — |

**Prescribed sequence:** Getting Started → Sales Workflow → Core Portfolio → Prospecting → Sales Skills → Pipeline Management

---

## TLG (hide from dashboards)

Jason Ackerman, Bianca Davis, James Parker, Resmie Biba, Chris Curtis, Sara Thompson, Jeremy MacBean, Bradley Pierce, Laura Sefcik, Samantha Maresca, Staci Musco, CJ Homer, Rich Moore, Dale Kinsey, John Lechner, Resmie Nesimi, Samantha D'Angelo, Bianca DiPasquale, Doug Falk
