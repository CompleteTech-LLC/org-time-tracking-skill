#!/usr/bin/env python3
"""Regenerate the committed demonstration files in assets/examples/ from the synthetic Northwind fixture.

Renders with the CompleteTech preset (it identifies this library; the registry bundle excludes assets/examples/).
The workbook needs a spreadsheet engine to store calculated values: desktop Excel via scripts/recalc_excel.ps1 on
Windows. Elsewhere, recalculate the workbook yourself (LibreOffice) and re-run build_report.py.
"""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import make_fixtures as fx  # noqa: E402

ROOT = fx.ROOT
OUT = fx.OUT


def main() -> int:
    shutil.rmtree(OUT, ignore_errors=True)
    OUT.mkdir(parents=True)
    (OUT / "lines.txt").write_text(fx.LINES, encoding="utf-8")
    (OUT / "meetings.json").write_text(json.dumps(fx.MEETINGS), encoding="utf-8")
    (OUT / "tasks.json").write_text(json.dumps(fx.TASKS), encoding="utf-8")
    fx.run(str(ROOT / "scripts/build_tree.py"), "--in", "lines.txt", "--out", "tree", "--start", "2026-07-01", "--tz", "America/New_York")
    (OUT / "example.json").write_text(json.dumps(fx.config("example", {"preset": "completetech"})), encoding="utf-8")
    print(fx.run(str(ROOT / "scripts/make_workbook.py"), "--config", "example.json").strip())
    powershell = shutil.which("powershell") or shutil.which("pwsh")
    if not powershell:
        print("no PowerShell/Excel here: recalculate tests/out/example.xlsx yourself, then run scripts/build_report.py")
        return 1
    result = subprocess.run([powershell, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ROOT / "scripts/recalc_excel.ps1"),
                             "-Path", str(OUT / "example.xlsx"), "-Cells", "Time Spent!D7|Time Spent!E7"], capture_output=True, text=True)
    print(result.stdout[-400:])
    if result.returncode or "error cells: 0" not in result.stdout:
        sys.stderr.write(result.stderr)
        raise SystemExit("recalculation failed or reported errors")
    fx.run(str(ROOT / "scripts/build_report.py"), "--config", "example.json", "--out", "example.html")
    target = ROOT / "assets" / "examples"
    target.mkdir(parents=True, exist_ok=True)
    shutil.copy(OUT / "example.xlsx", target / "example.xlsx")
    shutil.copy(OUT / "example.html", target / "example.html")
    print("wrote assets/examples/example.xlsx and example.html")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
