"""
build.py
--------
One-command refresh: parse all xlsx → rebuild CSV → rebuild dashboard HTML.

Usage:
    python scripts/build.py
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent


def run(script: str) -> None:
    print(f"\n>>> {script}")
    r = subprocess.run([sys.executable, str(HERE / script)], cwd=HERE.parent)
    if r.returncode != 0:
        sys.exit(f"{script} failed with exit code {r.returncode}")


def main() -> int:
    run("build_db.py")
    run("build_dashboard.py")
    print("\nDone. Open index.html in any browser.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
