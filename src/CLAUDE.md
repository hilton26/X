# CLAUDE.md — Project Context

## Who I Am
Hilton Netta — Investment Operations / Compliance at Prescient Investment Management (Cape Town, South Africa).

## What This Repo Is
Python automation scripts and Jupyter notebooks for compliance, reporting, and data processing tasks in an investment operations context.

## Key Domain
- **Regulation 28 / Reg 30**: South African financial regulations for pension and investment funds
- **Derivative exposure checking**: Daily derivative cover monitoring across funds
- **PGF UCITS Share Class Hedges**: Hedging reporting
- **Medical scheme reporting**: Circ 3/6/11/12 categorisation
- **Eagle**: Portfolio management system (web portal at `eagleportal.prescient.co.za`) — reports pulled via Selenium
- **Prime**: Internal data platform at `prime.prescient.co.za`

## Project Structure
- `constants.py` — single source of truth for ALL file paths, URLs, credentials (loaded from Excel), and domain constants. Always import from here.
- `*.py` / `*.ipynb` — scripts/notebooks usually correspond 1:1 (same logic, notebook for dev)
- Scheduled notebooks use `SCHEDULED_` prefix

## Key Paths (from constants.py)
- Network drive root: `\\PIM-CPT-FS.prescient.local\PIM-Documents$` (mapped as `P:`)
- Compliance folder: `pthCmp = pthPIM\Investment Operations\GRC\Compliance`
- Main Excel control file: `pthPy = pthCmp\Daily\py_reports.xlsm`
- Working folder: `pthW = pthPIM\Working Folders\Hilton\W`
- Git repo root: `C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo`
- FX rates output: `pth_fx = pthW\fx_rates.csv`

## Tech Stack
- Python (pandas, requests, openpyxl, selenium, streamlit, scipy)
- Jupyter notebooks (JupyterLab / VS Code)
- Excel (.xlsm, .xlsx) as control files and data sources
- Selenium for Eagle/Prime portal automation
- `frankfurter.app` for FX rates (USD/ZAR)

## Credentials
Loaded at runtime from `pthPy` sheet `creds` col J — never hardcoded.
`p_al` and `p_xe` are the credential variables in `constants.py`.

## Style Notes
- All shared constants live in `constants.py` — add new paths/files there
- Scripts are standalone (`if __name__ == "__main__"`) and also importable as modules
- SSL verification is disabled for internal network calls (`urllib3.disable_warnings`)

## Database access rules

- ALWAYS connect with `username='read.only'`, `password='read.only'`.
- Host: `pim-cpt-mysql-prod.prescient.local`
- Port 3306 for main databases; port 3307 for EAV / time-series (prime_eav, prime_profile, pim_ts, pim_static, benchmarks).
- NEVER use a credential with write access unless I explicitly tell you to for one specific, approved task.
- Prefer `ppym` for connections and queries. Connection patterns and worked examples:
  ppym source — https://gitlab-pim.prescient.co.za/dsqa/pim-python-module (ppym/data/db.py)
  ppym docs   — https://docs-ppym.prescient.co.za/ppym.html
- Add a `LIMIT` to exploratory queries (start with `LIMIT 100`).
- Avoid `SELECT *` on wide tables; name the columns you need.
- Don't run queries in a tight loop.
- I am the human in the loop: show me the SQL before you run it.
