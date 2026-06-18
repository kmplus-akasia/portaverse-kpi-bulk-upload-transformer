# KPI Planning Dashboard Production

Local Streamlit dashboard for monitoring KPI planning coverage in Portaverse production.

## Environment Variables

The app reads database credentials only from environment variables. Do not commit credentials.

```bash
export DB_HOST="..."
export DB_PORT="3306"
export DB_NAME="defaultdb"
export DB_USER="..."
export DB_PASSWORD="..."
```

Optional overrides:

```bash
export KPI_DASHBOARD_YEAR="2026"
export KPI_DASHBOARD_COMPANY_ID="1"
```

## Install

```bash
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements-dashboard.txt
```

## Smoke Test

```bash
python3 dashboard/kpi_planning_dashboard.py --check
```

## Run

```bash
streamlit run dashboard/kpi_planning_dashboard.py
```

If you want a single click / launch target from the editor, use the `Run analytic dashboard` config in `.vscode/launch.json` or run:

```bash
make run
```

If the live production database is temporarily unreachable, the dashboard will fall back to a cached snapshot so the page still opens instead of stopping on a raw MySQL error.
