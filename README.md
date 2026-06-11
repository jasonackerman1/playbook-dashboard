# Playbook Traffic Dashboard

Live dashboard tracking Konica Minolta playbook page views across all regions.

## Monthly Update Workflow

1. Drop the new Excel file into the `/data` folder
   - File must follow this naming pattern: `playbook-monthly-YYYY-MM.xlsx`
   - Example: `playbook-monthly-2026-06.xlsx`

2. Push to GitHub:
   ```
   git add data/playbook-monthly-YYYY-MM.xlsx
   git commit -m "Add [Month] [Year] report"
   git push
   ```

3. GitHub Actions runs automatically and updates the live dashboard within ~60 seconds.

## Local Development

To rebuild the dashboard locally:

```
pip install -r requirements.txt
python update_dashboard.py
```

Then open `index.html` in your browser.

## Folder Structure

```
playbook-dashboard/
├── .github/
│   └── workflows/
│       └── update-dashboard.yml
├── data/
│   └── playbook-monthly-YYYY-MM.xlsx
├── index.html          ← auto-generated, do not edit manually
├── update_dashboard.py
├── requirements.txt
└── README.md
```

## Playbooks Tracked

- Salesforce Playbook
- Healthcare Playbook
- Public Sector Playbook
- Accelerate
- DX Playbook
- Legal Playbook
- IQ501
- GC/IP Sales Playbook
