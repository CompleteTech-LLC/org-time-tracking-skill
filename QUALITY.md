# Quality Checks

Run the production-readiness checks from this repository root:

```bash
python3 scripts/validate_quality.py
```

The validator compiles Python files, runs Ruff, parses YAML/JSON/INI files, renders Mermaid sources when Mermaid tooling is available, runs `--help` smoke checks on the entry points, runs the synthetic fixture suite (`tests/make_fixtures.py`, which must print `ALL OK`), checks that committed example pages are self-contained and carry a Content-Security-Policy, syntax-checks the browser and PowerShell scripts when Node or PowerShell are available, and checks ClawHub bundle readiness.

GitHub Actions runs the same validator on push and pull request. The fixture suite does not evaluate workbook formulas; recalculate a workbook in Excel or LibreOffice and compare totals when changing `scripts/make_workbook.py`.
