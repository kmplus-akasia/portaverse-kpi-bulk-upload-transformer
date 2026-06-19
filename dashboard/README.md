# KPI Planning Dashboard Production

Local Streamlit dashboard for monitoring KPI planning coverage and worker portfolio progress across Pelindo production companies.

The default scope is **Seluruh Pelindo**. Administrators can drill down to one active Company ID, filter worker readiness/origin, and download the current follow-up list as CSV.

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
export KPI_DASHBOARD_COMPANY_ID="all" # or one numeric Company ID
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
python3 dashboard/kpi_planning_dashboard.py --check --company-id 1
```

## Run

```bash
streamlit run dashboard/kpi_planning_dashboard.py
```

If you want a single click / launch target from the editor, use the `Run analytic dashboard` config in `.vscode/launch.json` or run:

```bash
make run
```

## Worker Progress Contract

Portfolio origin is derived from the backend write contract:

- `KAMUS_KPI`: the active position has KPI ownership created with `created_by_pov=SYSTEM`.
- `MANUAL_TANPA_KAMUS`: no system KPI exists for the position and the active worker has KPI ownership created by `WORKER` or `SUPERIOR`.
- `BELUM_ADA_PORTFOLIO`: neither system nor manual KPI exists.
- `ORIGIN_TIDAK_DIKENAL`: KPI rows exist with an unsupported or missing origin value.

Readiness uses the same five workflow buckets as My Team Performance: Belum Ada Draft, Draft Perencanaan, Menunggu Review Bawahan, Menunggu Keputusan Anda, and Disetujui. Detail is one row per active worker-position assignment; worker summary uses the least advanced status when a worker has multiple active assignments.

The dashboard includes definitive, Lakhar, and job-sharing assignments. Structural positions use PMID; non-structural positions use PNID mapping for the active company and group.

## Failure Behavior

If the live database is unavailable, the dashboard reports `unavailable` and does not show cached Head Office metrics. Employee data and database credentials are never written to the repository. Upload audit remains global because import logs do not contain a reliable Company ID for filtering.
