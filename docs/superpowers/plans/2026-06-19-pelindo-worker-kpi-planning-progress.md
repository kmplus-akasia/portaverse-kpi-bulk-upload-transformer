# Expand KPI Planning Progress Across Pelindo

This ExecPlan is a living document. Keep `Progress`, `Surprises & Discoveries`, `Decision Log`, and `Outcomes & Retrospective` current while implementation proceeds.

## Purpose / Big Picture

Give administrators one read-only dashboard for tracking KPI planning readiness for every active worker-position assignment across Pelindo. The dashboard must distinguish system-provided KPI dictionaries from manual planning, expose workers who need follow-up, and support a company drill-down without presenting stale Head Office fallback data as live information.

## Context and Orientation

The implementation lives in `dashboard/kpi_planning_dashboard.py`, with operator instructions in `dashboard/README.md` and regression coverage in `tests/test_kpi_planning_dashboard.py`. The dashboard reads MySQL directly using environment credentials and must remain SELECT-only. Structural positions use PMID (`position_master_type_id = 5`); non-structural positions use PNID from `position_nomenclature_mapping`.

## Scope and Approach

Use grouped read-only SQL to return one row per active worker-position assignment, then pure Python helpers to derive portfolio origin, readiness, worker-level rollups, filters, summaries, and CSV output. Company scope defaults to all active Pelindo companies and may be narrowed to one Company ID. No backend endpoint, schema migration, new dependency, committed employee snapshot, or secret is in scope.

## Milestones

### Milestone 1: Tested progress domain

Add failing tests, then implement origin/readiness classification, worker rollup, filters, summaries, Company ID parsing, and CSV serialization.

### Milestone 2: All-company data contract

Extend the active assignment query and KPI aggregates, add company options, include progress data in `--check`, and keep upload audit explicitly global.

### Milestone 3: Administrator workflow

Add company/origin/readiness/search controls, progress cards and charts, the follow-up table and CSV download, then remove the misleading cached Head Office fallback.

### Milestone 4: Verification and handoff

Run unit tests, compile checks, fixture/AppTest coverage, live read-only smoke checks when credentials exist, and browser verification of the local Streamlit page.

## Validation

- `.venv/bin/python -m unittest discover -s tests -v`
- `.venv/bin/python -m compileall dashboard`
- `.venv/bin/python dashboard/kpi_planning_dashboard.py --check`
- `.venv/bin/python dashboard/kpi_planning_dashboard.py --check --company-id 1`
- `.venv/bin/python -m streamlit run dashboard/kpi_planning_dashboard.py`

## Progress

- [x] 2026-06-19: Approved product/data/UI plan reviewed against dashboard and PMS source contracts.
- [x] 2026-06-19: Feature branch created and baseline verified (42 tests passing after installing the existing `openpyxl` requirement into `.venv`).
- [x] 2026-06-19: Milestone 1 completed with red-green coverage for origin, readiness, rollup, filtering, CSV, and SQL scope.
- [x] 2026-06-19: Milestone 2 completed with all-company query builders, worker progress aggregation, and extended check payload.
- [x] 2026-06-19: Milestone 3 completed with company/origin/readiness filters, progress analytics, follow-up CSV, and fail-closed DB behavior.
- [x] 2026-06-19: Milestone 4 completed for local verification: 52 tests passed, compile and diff checks passed, AppTest covered full/failed states, and browser DOM confirmed fail-closed behavior.
- [ ] Live production SQL remains blocked because `DB_HOST`, `DB_PORT`, `DB_NAME`, `DB_USER`, and `DB_PASSWORD` are absent from this session.

## Surprises & Discoveries

- `kpi_v3.source` cannot separate uploaded dictionary content from manual drafts because the bulk importer persists `source='MANUAL'`; `created_by_pov` is the reliable contract.
- `kpi_template_import_log` has actor company text but no reliable Company ID, so import audit remains global.
- The prior cached snapshot contains Head Office-only values and cannot safely represent all-company or worker-level views.
- Coverage previously counted all KPI rows as dictionary availability; it now counts only `created_by_pov=SYSTEM`, while worker readiness still evaluates the complete employee portfolio.
- The local environment has no `ruff` executable; syntax, unittest, AppTest, and `git diff --check` are the available verification paths.

## Decision Log

- Decision: Use `created_by_pov=SYSTEM` for dictionary/Performance Tree origin and `WORKER/SUPERIOR` for manual origin. Rationale: this matches actual write paths. Date/Author: 2026-06-19, Codex with user approval.
- Decision: Include definitive, Lakhar, and job-sharing assignments; summarize distinct workers but keep detail per worker-position. Rationale: no active responsibility is hidden. Date/Author: 2026-06-19, Codex with user approval.
- Decision: Default company scope to all Pelindo and stop on DB failure without cached metrics. Rationale: prevents stale Head Office numbers from appearing under another scope. Date/Author: 2026-06-19, Codex with user approval.

## Outcomes & Retrospective

The all-company worker progress workflow, filters, CSV export, CLI payload, and fail-closed behavior are implemented. Local verification passes. Production result counts and the 30-second query acceptance criterion remain unverified until DB credentials are available in the runtime environment.
