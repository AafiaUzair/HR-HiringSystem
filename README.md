# HR Hiring System

A working model of an HR talent acquisition and onboarding workflow, built as a capstone project.
Move candidates through a hiring pipeline, score interviews, track new-hire onboarding and watch the KPIs update.

> All people, emails and figures are simulated (fictional company: Northwind Services Pvt. Ltd.).

## Features

- **Pipeline board:** 6 stages, duplicate-email check, 60-point screening cut-off, mandatory rejection reason, auto-created onboarding checklist on hire
- **Scorecards:** 5 weighted criteria, weighted score, decision guide, 24-hour feedback SLA flag, debrief-gap flag
- **Onboarding tracker:** checklist, LMS module progress, Day 10 reminder
- **KPI dashboard:** time-to-hire, retention, feedback turnaround, SLA compliance, pilot score
- **Excel workbook** (`data/`) with the same logic as live formulas

## Tech

| Part               | Technology                                                  |
| ------------------ | ----------------------------------------------------------- |
| App                | HTML, CSS, vanilla JavaScript (no framework, no build step) |
| Storage            | Browser `localStorage` (no server database)               |
| Font               | Google Fonts (Public Sans)                                  |
| Workbook generator | Python 3 with `openpyxl`                                  |
| Hosting            | Vercel (static site)                                        |

## Run locally

Open `index.html` in a browser, or run a local server:

```bash
python3 -m http.server 8000
# then open http://localhost:8000
```

## Deploy to Vercel

1. Push this folder to a GitHub repository.
2. In Vercel choose **Add New > Project**, import the repository.
3. Framework Preset: **Other**. Leave build command and output directory empty. Deploy.

## Rebuild the Excel workbook

```bash
pip install openpyxl
cd tools && python3 build_workbook.py
```

Open the file in Excel or Google Sheets so the formulas calculate.

## Data note

The app saves to the browser's `localStorage`, so data stays on one device and browser. Use **Reset demo data** to restore the starting data.
