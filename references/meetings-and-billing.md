# Meetings, the held-meeting rule and billing semantics

Every number here is a parameter in `config.json` (see the table in [SKILL.md](../SKILL.md)); nothing is a constant of the skill.

## Meetings
Ask early whether meetings are wanted; they change the time numbers a lot. Show the calendar and the Teams cards side by side and never assume a scheduled slot was attended.

`meetings.json` is a list of series: `{title, group: work|unconfirmed, start, end, dates: [...], status?, actual_min?, evidence?}`. The builder expands them, applies the held-meeting rule and writes the Meetings tab.

- **Held-meeting rule** (`meetings.held_rule`). With `teams_conversation_same_day` a scheduled meeting counts as held only if there was Teams conversation that day, even if the calendar shows it or a meeting card exists; list cards that fall on no-conversation days for the operator to confirm. With `none` the calendar is trusted.
- **Explicit status** in a series (for example `Cancelled`, `Likely not held (chat)`) always wins over the rule. Use it when chat or mail shows a postponement ("can we move it to tomorrow") or a cancellation.
- **Unconfirmed group**: recurring entries that do not look like the organization's work (personal blocks, unrelated webinars) are listed but excluded from totals behind a Yes/No toggle (`billing.include_unconfirmed`) and raised with the operator. Remove them only when the operator says they are personal.

## Billing semantics
- Each lapse (a Teams session, an email written, a meeting) is `CEILING(MAX(actual, min_lapse_min), increment_min)`; lapses are rounded individually, then summed. Elapsed minutes are rounded to whole minutes first so floating-point noise cannot push a value into the next block.
- Cancelled and not-held items stay 0.
- A day with a held meeting bills at least `meeting_day_min`, topped up from the meeting: the day minimum minus the meeting minutes already billed that day. Other chat and email that day add on top.
- Show elapsed and billed minutes side by side so the uplift is visible, and confirm borderline cases (a 2-minute call, a call held per chat only, a card on a no-conversation day) with the operator.
- Estimates are not invoices. Hand verified hours to the billing skill; do not present this workbook as a bill.
