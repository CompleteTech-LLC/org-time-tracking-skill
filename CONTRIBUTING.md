# Contributing to CompleteTech LLC Skills

[Start here](ONBOARDING.md) · [Branding and handoffs](BRANDING.md) · [Maintainer instructions](AGENTS.md) · [Specialist instructions](SKILL.md)

## Compatibility

Keep each skill independently usable. Preserve activation keys, root generator CLIs, templates, runtime dependencies, consent and approval gates. This skill's neutral default branding and its parameter-only design (no organization, person, date, rule or brand value in a script) are intentional, not gaps to remove. Before adding a constant to a script, add a config parameter and a row to the table in SKILL.md.

## Verification

From a full checkout using Python 3.12, the package CI baseline:

```bash
python -m unittest discover -s tests -p 'test_package_*.py' -v
python scripts/validate_package.py
```

Also run the fixture suite and the full Quality gate after installing its documented development tools:

```bash
python tests/make_fixtures.py
python scripts/validate_quality.py
```

The repository's `QUALITY.md` and [README.md](README.md) describe the checks. Package checks are read-only and do not execute specialist generators. Their link scan covers ordinary inline destinations in ONBOARDING.md, CONTRIBUTING.md, BRANDING.md and AGENTS.md, not the entire documentation tree, reference-style links, HTML or anchor validity. PNG checks cover bounded chunk structure and CRC integrity, not decompression, pixel decoding or artwork. The fixture suite evaluates no workbook formulas (that needs Excel or LibreOffice); recalculate and compare totals after any change to `make_workbook.py`. Do not weaken checks to obtain a green result.

## Coordinated changes

Keep `skill-package.json`, onboarding and actual entry points synchronized. Shared validator, tests, CI, maintainer instructions and branding guidance must match the orchestrator's shared-file audit; `scripts/validate_package.py`, `tests/test_package_contract.py`, the workflows and `AGENTS.md` are shared files. Use focused PRs and report exact checks, failures and limitations. Respect branch protection, review requirements and changed heads.

Use synthetic fixtures and new output paths instead of overwriting committed previews. Never commit secrets, real mail, chats, calendars, client names, workbooks or reports. Code remains under [LICENSE](LICENSE); branding remains under [BRAND_ASSETS.md](BRAND_ASSETS.md). Do not redistribute private assets or fonts. Registry releases and repository visibility changes require separate authorization.
