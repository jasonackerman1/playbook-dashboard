# Playbook Dashboard — Claude Instructions

## DO NOT PUSH
Do not push to GitHub unless Jason explicitly says to. All development is tested locally first.

---

## Dashboards

| File | Script | Description |
|---|---|---|
| `index.html` | `update_dashboard.py` | Playbook Traffic Dashboard |
| `cert-healthcare.html` | `update_cert_dashboard.py` | Healthcare Certification Dashboard |
| `cert-publicsector.html` | `update_cert_dashboard.py` | Public Sector Certification Dashboard |
| `onboarding.html` | `update_onboarding_dashboard.py` | Accelerate Onboarding |

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

## Navigation (hamburger menus)

Every dashboard links to every other dashboard. When adding a new dashboard, update the hamburger menu in **all** scripts.

| Dashboard | Links to |
|---|---|
| Playbook | HC Cert, PS Cert, Accelerate Onboarding |
| Healthcare Cert | Playbook, PS Cert, Accelerate Onboarding |
| Public Sector Cert | Playbook, HC Cert, Accelerate Onboarding |
| Accelerate Onboarding | Playbook, HC Cert, PS Cert |

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

---

## TLG (hide from dashboards)

Jason Ackerman, Bianca Davis, James Parker, Resmie Biba, Chris Curtis, Sara Thompson, Jeremy MacBean, Bradley Pierce, Laura Sefcik, Samantha Maresca, Staci Musco, CJ Homer, Rich Moore, Dale Kinsey, John Lechner, Resmie Nesimi, Samantha D'Angelo, Bianca DiPasquale, Doug Falk
