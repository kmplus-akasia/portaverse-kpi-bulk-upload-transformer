from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


def main() -> int:
    repo_root = Path(__file__).resolve().parents[1]
    dashboard_entry = repo_root / "dashboard" / "kpi_planning_dashboard.py"

    host = os.getenv("KPI_DASHBOARD_HOST", "0.0.0.0")
    port = os.getenv("KPI_DASHBOARD_PORT", "8501")

    cmd = [
        sys.executable,
        "-m",
        "streamlit",
        "run",
        str(dashboard_entry),
        "--server.address",
        host,
        "--server.port",
        port,
    ]

    return subprocess.call(cmd, cwd=str(repo_root))


if __name__ == "__main__":
    raise SystemExit(main())
