# Workbook and report reference

`python scripts/make_workbook.py --config config.json` builds the time-tracking workbook from the evidence tree plus `meetings.json` and `tasks.json`. All totals are formulas; billing inputs live on Time Spent and everything downstream follows. Follow the xlsx conventions: a professional font, no hard-coded totals, blue/yellow inputs, assumptions visible where the reader sees them.

| Tab (in workbook order) | Contents |
|---|---|
| Summary | Brand banner (eyebrow, title, tagline, optional logo), then time first: billed hours (logged and best case), days with an actual meeting, Teams sessions, emails written, task effort; evidence counts last; notes naming the active billing parameters. |
| Time Spent | The billing engine: line items (Teams sessions, emails written, held meetings, meeting-day top-up, optional unconfirmed), totals, a best-case total, and the inputs: I1 unconfirmed toggle, I2 meeting basis, I3 minimum per lapse, I4 increment, I5 billed minutes per email, I7 meeting-day minimum. |
| Calendar View, Daily Time, Day Detail | Month grids whose day cells link to that day's events; one row per day with a column per line item (the check row must tie to Time Spent); every time-bearing event with minutes, decimal hours, h:mm, a day total and a source hyperlink. A calendar that only shows totals is not enough. |
| Meetings | Calendar rows: scheduled minutes, status, Teams actual minutes, evidence, qualification flags, Teams messages that day. |
| Teams Sessions, Teams Messages | Sessions from a gap rule (only sessions where the operator posted); every message with a link back to its session. Day Detail links to the messages, not just the summary row. |
| Tasks | Id, task, source, requested by, raised, status, low/high/likely hours (likely = average), basis, evidence (Requested vs Suggested). Estimates come from the scope described in the messages and are labeled as estimates. Mark unrequested work "Suggested". |
| Communications, Alert Rollup | Evidence: one row per mail or Teams item parsed from the tree so sheet and files agree; automated mail counted by type and month. |

## Branding and theme in the workbook
Identity resolves through `scripts/branding.py`: neutral defaults, then the chosen preset (`templates/branding.<preset>.json`), then explicit `branding` values from the config. The accent colors table headers, tab colors and the calendar heat scale (white to brand tints); readable header text is chosen by contrast. The banner shows `eyebrow`, title and `tagline`; `logo` must be a local image file (shown only when Pillow is installed). `footer` prints on every page. Yellow input cells stay yellow in every brand: they are a semantic convention, not a theme color. Excel has no light/dark switch, so the workbook uses the light tokens; the HTML report carries both.

## HTML report
`python scripts/build_report.py --config config.json` renders a single self-contained HTML file from the **recalculated** workbook: brand header, hours-first KPI tiles, billed time by line item, a per-day heat calendar, held and not-counted meetings, tasks. Tokens `bg, surface, surface-2, ink, ink-2, ink-3, line, line-2` come from the brand's `light` and `dark` sets; `style.theme` (`auto`, `light`, `dark`) or `--theme` picks the mode, and `auto` follows the reader's system. No external fonts, scripts or images; the logo is inlined; a Content-Security-Policy meta tag forbids everything else.

## Recalculation and verification
- With desktop Excel: `scripts/recalc_excel.ps1 -Path book.xlsx -Cells "Time Spent!D7|Summary!B6:B15"` (full recalculation, error scan, save so cached values are stored). Never run it on a workbook that is open in Excel with unsaved changes; write a new file name instead.
- Without Excel use LibreOffice; without either, verify totals independently and tell the operator the cached values are empty until Excel recalculates.
- Test any change to `make_workbook.py` by regenerating the previous workbook from the same inputs and comparing the totals.
- Email times are not available from the Outlook list view, so email effort is an assumption.
