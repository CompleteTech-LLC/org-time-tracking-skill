# ClawHub Publishing

This repository is prepared for ClawHub publishing as a text-based OpenClaw skill bundle.

## Included In ClawHub

The ClawHub bundle is intended to include text-based skill material only:

- `SKILL.md`, `README.md`, `QUALITY.md`, and this file
- `agents/openai.yaml`
- `pyproject.toml`
- `.github/workflows/quality.yml`
- `scripts/*.py`, `scripts/*.js`, `scripts/*.ps1`
- `references/` text files and `templates/` JSON
- `assets/diagrams/*.mmd`

## Excluded From ClawHub

`.clawhubignore` excludes binary and generated assets from the publish candidate: PNG logos and previews, workbooks, PDFs, DOCX files, fonts, the rendered `assets/examples/` demonstrations, working directories (`exports/`, `trees/`, `tests/out/`), caches, virtual environments, local env files and `scripts/validate_quality.py`.

Those files remain part of the GitHub repository for brand presentation and demonstration. The registry bundle is neutral: it carries no logo and no committed company-branded artifact.

## License And Brand Boundary

ClawHub publishes skills under MIT-0. The text/code bundle can be used under ClawHub's publishing terms, but CompleteTech LLC names, logos, seals, and other brand assets remain reserved. Publishing this text bundle does not grant a trademark or brand-asset license and does not relicense excluded binary brand assets. The `completetech` branding preset is configuration text that names the company; it does not include or license the logo.

## Runtime Dependencies

Runtime requirements are declared in `SKILL.md` under `metadata.openclaw`: `python3`, `openpyxl`, and `pyyaml` for the quality validator. Pillow is an optional convenience for embedding a local logo in the workbook. Desktop Excel is optional and only used by `scripts/recalc_excel.ps1` on Windows.

## Local Readiness Check

Run before publishing:

```bash
python3 scripts/validate_quality.py
```

The validator checks lint, Python compilation, structured-file parsing, Mermaid rendering, smoke tests, the fixture suite and ClawHub bundle readiness. It does not publish to ClawHub.

## Publishing

Do not publish automatically. Use the ClawHub CLI only after explicit approval and an authenticated owner context, for example:

```bash
clawhub skill publish . --owner <owner> --version <semver>
```
