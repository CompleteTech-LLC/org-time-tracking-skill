**CompleteTech LLC Skills** · [Start here](ONBOARDING.md) · [Agent instructions](SKILL.md) · [Branding](BRANDING.md) · [Skill library](https://github.com/CompleteTech-LLC/agentic-services-orchestrator-skill/blob/main/references/skill-family.md) · [Contributing](CONTRIBUTING.md)

# Org Comms Export Skill

<p align="center">
  <img src="assets/logo.png" alt="CompleteTech LLC logo" width="260">
</p>

A CompleteTech LLC skill that exports one organization's Outlook mail, Teams chats and calendar meetings from a start date into a redacted folder tree, an Excel time-accounting workbook and a self-contained branded HTML report.

## About

Part of the CompleteTech LLC skill library. It reads the operator's own signed-in Microsoft 365 web session through a browser, turns the evidence into month/day files, and documents it with the operator's own billing rules: a minimum per lapse, a billing increment, a meeting-day minimum and a held-meeting rule. Every organization, person, date, rule and brand value is a parameter; identity is neutral unless an approved one is chosen.

## OpenClaw / ClawHub Metadata

- Skill key: `org-comms-export`
- Version-ready metadata: `1.0.0`
- Homepage: https://github.com/CompleteTech-LLC/org-comms-export-skill
- README: https://github.com/CompleteTech-LLC/org-comms-export-skill#readme
- Runtime binaries: `python3`
- Python packages: `openpyxl==3.1.5`, `pyyaml==6.0.3` (optional logo embedding: `pillow==12.2.0`)
- Windows only, optional: desktop Excel for `scripts/recalc_excel.ps1`
- License: repository code, templates, and documentation use MIT; brand assets are reserved, see `BRAND_ASSETS.md`.

## Workflow Diagram

Source: [assets/diagrams/workflow.mmd](assets/diagrams/workflow.mmd).

```mermaid
flowchart LR
  A[Parameters and consent] --> B[Browser collection: mail, chats, calendar]
  B --> C[Redact secrets]
  C --> D[Tree: staff, related, alerts, teams, meetings]
  D --> E[Workbook: sessions, meetings, tasks, time]
  E --> F{Recalculated and tied?}
  F -->|No| G[Fix and rebuild]
  F -->|Yes| H[Branded HTML report]
  classDef source fill:#eef2ff,stroke:#1e3a8a,color:#0f172a;
  classDef gate fill:#fff7e6,stroke:#c97a12,color:#3d2600;
  classDef output fill:#eefaf0,stroke:#2f8f46,color:#12351d;
  class A,B,C,D,E source;
  class F gate;
  class G,H output;
```

## What It Does

| Capability | Details |
|---|---|
| Collection | Outlook search results, Teams chat messages and calendar events through small, parameterized browser scripts. |
| Redaction | Passwords, tokens and keys are replaced before anything is written; extra patterns are configurable. |
| Tree | `staff/`, `related/`, `alerts/`, `teams/`, `meetings/` by `YYYY-MM/DD.md`. |
| Workbook | Communications, Teams sessions and messages, meetings, tasks with effort estimates, time spent, daily calendar and clickable event detail. All totals are formulas driven by the billing inputs. |
| Meetings | Calendar plus Teams meeting cards; a call counts as held only when there was Teams conversation that day (rule is a parameter). |
| Report | One self-contained HTML file with brand tokens in light and dark, a monthly heat calendar, meetings and tasks. |
| Branding | Neutral by default; the CompleteTech preset or the operator's own identity on request. |

## Contents

- `SKILL.md` - operating instructions, consent table and the parameter table.
- `references/` - browser mechanics, meeting and billing semantics, workbook and report reference.
- `scripts/` - `build_tree.py`, `make_workbook.py`, `build_report.py`, `branding.py`, browser collectors (`*.js`), `recalc_excel.ps1`, validators.
- `templates/` - `config.example.json`, `meetings.example.json`, `tasks.example.json`, `branding.neutral.json`, `branding.completetech.json`.
- `examples/` and `tests/` - synthetic fixtures and the end-to-end suite.

## Quick Start

```bash
pip install -r requirements.txt
python3 tests/make_fixtures.py          # synthetic end-to-end run, prints ALL OK
python3 scripts/build_tree.py --in lines.txt --out ./acme --start 2026-07-01 --tz "America/New_York"
python3 scripts/make_workbook.py --config config.json
python3 scripts/build_report.py --config config.json --theme auto
```

Copy `templates/*.example.json` to `config.json`, `meetings.json` and `tasks.json` and edit them; nothing is read from the repository's own examples at run time. The workbook has no cached values until Excel or LibreOffice recalculates it; the report needs those values.

## Example

![Branded report preview, light theme](assets/examples/example.png)

Rendered from the synthetic Northwind fixture with the CompleteTech preset: [report](assets/examples/example.html) (light and dark, follows the reader's system) and [workbook](assets/examples/example.xlsx). The fixture data is fictional. Regenerate with `python tests/make_examples.py` (Windows with desktop Excel).

## Runtime Permissions

| Area | Runtime behavior |
|---|---|
| Execution | Local Python entry points and, optionally, a PowerShell script that drives desktop Excel; browser collectors run inside the operator's own tab. |
| Reads | The operator's signed-in Outlook, Teams and calendar pages; the tree, config files and the local logo. |
| Writes | Only the operator-chosen tree, workbook and report paths. |
| Network | None from the scripts. The browser session reads Microsoft 365 as the operator would. |

| Not Included | Boundary |
|---|---|
| Mailbox changes | Never sends, deletes, moves or edits mail, chats or events. |
| Credentials | Never stores or prints passwords, tokens or keys; redacts those found in chats. |
| Billing | Does not invoice or authorize billing; estimates are evidence for the operator's own decision. |
| System changes | No persistence, no scheduled tasks, no privilege escalation. |

## License

Code, templates, and documentation are licensed under the MIT License. CompleteTech LLC names, logos, seals, and brand assets are reserved and are not licensed for reuse except to identify this project. See `LICENSE` and `BRAND_ASSETS.md`.
