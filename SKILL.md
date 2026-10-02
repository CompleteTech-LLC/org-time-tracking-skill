---
name: org-time-tracking
description: >-
  Track and account for the time spent with one organization from the evidence in Outlook-on-the-web mail, Teams-on-the-web chats and the calendar: an Excel time-tracking workbook (time spent under operator-set billing rules such as a minimum per lapse, a billing increment and a meeting-day minimum; a daily calendar and clickable event detail; tasks with effort estimates; meetings counted as held only when there was Teams conversation that day), backed by a redacted month/day evidence tree and a self-contained branded light/dark HTML report. Every organization, person, date, billing rule and brand value is a parameter; identity is neutral unless the operator chooses an approved one (the CompleteTech preset or their own). Disclosure: this skill reads the operator's own signed-in mailbox, chats and calendar through a browser session, writes redacted files and a workbook to a folder the operator chooses, and transmits nothing. Use when asked to track, tally, justify or bill time spent with a company ("how much time did I spend with acme.com since July", "put my time on a spreadsheet and a calendar", "apply my billing minimums", "show hours per day"), or to gather the supporting mail, Teams and meeting evidence.
version: 1.0.0
metadata:
  openclaw:
    skillKey: org-time-tracking
    homepage: https://github.com/CompleteTech-LLC/org-time-tracking-skill
    requires:
      bins:
        - python3
    install:
      - kind: uv
        package: openpyxl==3.1.5
      - kind: uv
        package: pyyaml==6.0.3
---

# Org Time Tracking Skill

## Purpose

| Use | Scope |
|---|---|
| Time tracking | Turn one organization's Outlook mail, Teams chats and calendar meetings, from a start date on, into billable time: sessions and meetings become lapses that are billed under the operator's own rules, shown per day on a calendar and traceable event by event. |
| Evidence | The same pass produces a redacted month/day tree (the proof behind every number) and a branded report. Communications are the evidence, not the product. |
| Starting point | Any request to track, tally, justify or bill time spent with a company domain, or to show hours on a spreadsheet, calendar or report. |
| Operating boundary | Read-only over the operator's own signed-in Microsoft 365 web session. Nothing is sent, deleted or modified in the mailbox, and nothing is transmitted anywhere. Hours are estimates from evidence, not invoices, and attendance and billing parameters are the operator's decisions. |

## Before You Start

Relay this table to the operator, in chat, before the first collection, and again if an opt-in below is about to be switched on.

| Topic | What the operator must hear |
|---|---|
| What will be read | Messages visible in the operator's own Outlook web mailbox that match the organization (sender list and preview text; full bodies only if requested), the Teams chats the operator confirms belong to the organization (every message since the start date), and the Outlook calendar plus Teams meeting cards. Nothing is opened that the signed-in session cannot already see. |
| What is written | A tree of Markdown files, a workbook and an HTML report under a folder the operator chooses. Passwords, tokens and keys found in chats are replaced by a placeholder before anything is written; the operator is told that such secrets exist and should be rotated. Real exports contain client names and activity: keep them out of Git and out of shared locations. |
| Leaving the machine | Nothing. The scripts make no network calls. The browser session reads Microsoft 365 pages as the operator would. Data is moved from the page to the agent through the page-text channel, never through URLs or other hosts. |
| Decisions that are the operator's | Which Teams chats belong to the organization; billing rules (minimum per lapse, increment, meeting-day minimum); whether a scheduled meeting counts as held; whether recurring unrelated items are personal; and which identity (brand) the output carries. |
| Coverage limits | The mail list view carries no times for older mail; searches can hide months; cancelled calendar occurrences vanish; Teams meeting cards are the only real durations; recaps and transcripts are not read. These are reported, not hidden. |
| How to remove everything | Delete the output folder, the workbook and the report. Nothing else is stored. |

## System Boundary

| Boundary | Use |
|---|---|
| This skill | Collection, redaction, the tree, the workbook, the report and the evidence behind each number. |
| `agentic-invoice-skill` | Billing a client: feed it verified hours from the workbook; an estimate here is not authority to bill. |
| `agentic-security-review-skill` | Questions about who may read the exported chats or how they are stored. |
| `ai-usage-ledger` | Local AI-agent usage and cost, a different data source. |

## Parameters (nothing is hard-coded)

**Never write an organization name, domain, person, date, billing number, threshold or brand value into a script.** Every such value lives in `config.json`, `meetings.json`, `tasks.json` (templates in `templates/`) or a script's `PARAMS` block. Ask for any value you do not have; do not silently reuse a previous run's values.

| Parameter (config path) | Meaning | Default / template value |
|---|---|---|
| `org.name`, `org.domain` | label and email domain | ask |
| `tree_root`, `output_xlsx` | where the tree and workbook live | ask |
| `period.start`, `period.end`, `period.timezone` | inclusive date range and the operator's zone | ask |
| `me.teams_names`, `me.mail_markers`, `me.draft_prefix`, `me.calendar_subject_prefixes` | how to tell my Teams posts, mail I wrote or drafted, and calendar items I organized | ask |
| `alerts.sender_label` | sender shown for automated mail | ask |
| `billing.min_lapse_min`, `billing.increment_min` | minimum billed per lapse (a Teams session, an email written, a meeting) and the block it rounds up to | ask (template 15 / 15) |
| `billing.email_actual_min` | assumed real minutes per email written | ask (template 5) |
| `billing.session_gap_min`, `billing.session_padding_min` | gap that splits Teams sessions; padding per session | template 20 / 2 |
| `billing.meeting_day_min` | minimum billed for a day with an actual meeting (0 = none) | ask (template 240) |
| `billing.meeting_basis`, `billing.include_unconfirmed` | `Logged` or `Best case`; `Yes`/`No` for unconfirmed items | Logged / No |
| `meetings.held_rule` | `teams_conversation_same_day` (no Teams conversation that day = no meeting) or `none` | ask |
| `meetings.work_group_label`, `unconfirmed_group_label`, `labels.*` | group and status names | template |
| `meetings_file`, `tasks_file` | JSON lists of meeting series and task estimates | templates |
| `branding.preset`, `branding.*` | identity: `neutral`, `completetech`, or explicit name, eyebrow, tagline, contact, logo (local file), accent, footer, font, light/dark tokens | `neutral` |
| `style.*` | font, input color, calendar color scale (minutes), minute unit, report theme `auto`/`light`/`dark` | template |
| `PARAMS` in each `*.js` | start date (required), scroll step/pause, waits, text limits, row-skip regex, months back | edit the block or set `window.__PARAMS` |
| `build_tree.py` flags | `--start --tz --patterns --placeholder --categories` | see the script header |
| `recalc_excel.ps1` args | `-Path`, `-Cells "Sheet!A1|Sheet!B2:B9"`, `-NoSave` | pass them |

When reviewing or extending this skill, grep the scripts for literals (names, dates, minute counts, colors) and move any you find into the config.

## Branding and theme

Resolve identity before rendering. An approved engagement brand takes precedence; preserve the operator's explicit settings and ask about unresolved identity instead of silently applying company branding. Output is **neutral by default**. Use `branding.preset: "completetech"` only for authorized CompleteTech output; it applies the CompleteTech accent, palette, tagline and footer and the local `assets/logo.png`. Do not synthesize a logo, fetch remote artwork or fonts, or move brand assets between repositories; `BRAND_ASSETS.md` governs the marks. See [BRANDING.md](BRANDING.md) for the palette and handoffs. The HTML report themes light and dark from the brand tokens; the workbook uses the light tokens.

## Workflow

1. **Set the rules.** Confirm the parameters above and relay Before You Start. Ask for the period, time zone, the billing rules (minimum per lapse, increment, meeting-day minimum, held-meeting rule), whether recurring unrelated items are personal, and which identity (brand) the output carries.
2. **Gather the evidence.** Collect Outlook mail, Teams chats and (optionally) calendar and meeting cards with the `scripts/*.js` collectors; mechanics in [references/browser-collection.md](references/browser-collection.md). Redact, write lines, run `python scripts/build_tree.py --in lines.txt --out <out>/<org> --start <start> --tz "<tz>"`. Line format: `YYYY-MM-DD|HH:MM|category|source|who|text`, where category is staff, related, alerts or teams, and `source` is the chat, the alert type, or "mail".
3. **Decide what counts.** Write `meetings.json` (series and dates; explicit status only where chat or mail proves a cancellation or postponement), `tasks.json` (effort estimates) and `config.json`; rules in [references/meetings-and-billing.md](references/meetings-and-billing.md).
4. **Compute the time.** `python scripts/make_workbook.py --config config.json`; recalculate (`scripts/recalc_excel.ps1` or LibreOffice) and check the Daily Time and Day Detail "ties" cells. Tabs in [references/workbook.md](references/workbook.md): Time Spent, Calendar View, Daily Time and Day Detail come first; the evidence tabs follow.
5. **Present it.** Open the workbook for the operator; `python scripts/build_report.py --config config.json` renders the branded HTML report (needs the recalculated workbook).
6. **Report** (see the end of this file), including every borderline case that needs the operator's decision.

## Hard rules (learned the hard way)

1. **Never write secrets to disk.** Chats routinely contain passwords, client secrets, API keys and tokens. Redact before writing (`build_tree.py` applies built-in patterns; `--patterns FILE` adds more), read the output for leftovers, tell the operator credentials exist, and never quote them in a reply.
2. **Do not run broad process-kill or cleanup commands** (`taskkill /IM python.exe`, `pkill python`). Never `cd` into a temp folder you intend to delete. Keep scratch files in the scratchpad directory. Edit scripts by exact anchors, not by deleting lines that match a phrase.
3. Do not bypass page restrictions; use the page-text transfer.
4. If a workbook is open in Excel with unsaved changes, never overwrite or recalculate it: write a new version name.
5. Meeting attendance, billing and identity are the operator's decisions. A clean build, a recalculated workbook or a rendered report does not authorize sending, billing or sharing.

## Final report to the operator (short)

Counts per category, folder path, the parameters used (billing rules, held-meeting rule, identity), gaps (months with no results, details not captured), chats and meetings excluded and why, borderline cases needing a decision, any credentials found (never repeat them), and anything unverified (for example sender classification from display names). Review the actual output for clipping, contrast, placeholders and factual accuracy before approved use.
