#!/usr/bin/env python3
"""Build the time-accounting workbook from an org-time-tracking tree. Everything org- or user-specific comes from a
JSON config (see config.example.json); nothing is hard-coded here.

Usage:
  python make_workbook.py --config config.json
Reads: <tree_root>/{staff,related,alerts,teams}/YYYY-MM/DD.md (format written by build_tree.py),
       the meetings JSON and tasks JSON named in the config.
Writes: the .xlsx named in the config (all totals are formulas; recalculate in Excel or LibreOffice afterwards).
"""

import argparse
import datetime as dt
import glob
import json
import os
import re
import sys
import calendar as _cal
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import branding as _brand  # noqa: E402

ap = argparse.ArgumentParser()
ap.add_argument("--config", required=True)
cfg_path = ap.parse_args().config
CFG = json.load(open(cfg_path, encoding="utf-8"))
BASE_DIR = os.path.dirname(os.path.abspath(cfg_path))


def path(p):
    return p if os.path.isabs(p) else os.path.join(BASE_DIR, p)


def get(*keys, default=None):
    cur = CFG
    for k in keys:
        if not isinstance(cur, dict) or k not in cur:
            return default
        cur = cur[k]
    return cur


# ---------------- parameters ----------------
ORG = get("org", "name", default="Org")
ROOT = path(get("tree_root"))
OUT = path(get("output_xlsx"))
START = dt.datetime.strptime(get("period", "start"), "%Y-%m-%d")
END = dt.datetime.strptime(get("period", "end"), "%Y-%m-%d")
TZ = get("period", "timezone", default="")
ME_FIRST = get("me", "teams_names", default=[])  # names used for my posts in the Teams tree lines
ME_MAIL = get("me", "mail_markers", default=[])  # substrings of a sender that mean "I wrote it"
DRAFT_PREFIX = get("me", "draft_prefix", default="[Draft]")
CAL_PREFIXES = get("me", "calendar_subject_prefixes", default=["Meeting"])
ALERT_SENDER = get("alerts", "sender_label", default="automated sender")
B = CFG.get("billing", {})
MIN_LAPSE = B.get("min_lapse_min", 0)
INCREMENT = B.get("increment_min", 1)
EMAIL_ACTUAL = B.get("email_actual_min", 5)
GAP = B.get("session_gap_min", 20)
PAD = B.get("session_padding_min", 2)
DAY_MIN = B.get("meeting_day_min", 0)
BASIS = B.get("meeting_basis", "Logged")
INCL_UNCONF = B.get("include_unconfirmed", "No")
G_WORK = get("meetings", "work_group_label", default=f"{ORG} work")
G_UNC = get("meetings", "unconfirmed_group_label", default="Unconfirmed (personal?)")
LBL = get("meetings", "labels", default={})
S_HELD_CARD = LBL.get("held_card", "Held (Teams record + conversation)")
S_HELD_CONV = LBL.get("held_conversation", "Held (conversation that day)")
S_NOT_HELD = LBL.get("not_held", "Not held (no conversation that day)")
S_CANCEL = LBL.get("cancelled", "Cancelled")
S_LIKELY = LBL.get("likely_not_held", "Likely not held (chat)")
S_DEFAULT = LBL.get("scheduled", "Scheduled, no record")
RULE = get("meetings", "held_rule", default="teams_conversation_same_day")  # or "none" (trust the calendar)
STY = CFG.get("style", {})
CLIENT_LOGO = _brand.client_logo(CFG.get("org"), BASE_DIR)  # optional client mark (org.logo)
BR = _brand.resolve(CFG.get("branding"), BASE_DIR)  # neutral unless a preset / explicit values are chosen
F = STY.get("font") or BR.get("font") or "Arial"
HDR_COLOR = BR["accent"].lstrip("#")  # table headers, tab colour, banner
ON_HDR = _brand.on_colour(BR["accent"]).lstrip("#")  # readable text on the accent
INP_COLOR = STY.get("input_color", "FFFF00")  # semantic: yellow = cells the operator may edit
MINE_COLOR = BR["light"]["surface-2"].lstrip("#")  # my Teams messages
SCALE = STY.get("calendar_color_scale", {"mid_min": 60, "max_min": 240})
CAL_UNIT = STY.get("minute_unit", "m")
TASK_ROWS = json.load(open(path(get("tasks_file")), encoding="utf-8")) if get("tasks_file") else []
MEET_SERIES = json.load(open(path(get("meetings_file")), encoding="utf-8")) if get("meetings_file") else []

HDR = PatternFill("solid", fgColor=HDR_COLOR)
INP = PatternFill("solid", fgColor=INP_COLOR)
MINE = PatternFill("solid", fgColor=MINE_COLOR)
base = Font(name=F, size=10)
bold = Font(name=F, size=10, bold=True)
white = Font(name=F, size=10, bold=True, color=ON_HDR)
blue = Font(name=F, size=10, color="0000FF")
note = Font(name=F, size=9, italic=True)
LINKF = Font(name=F, size=10, color="0563C1", underline="single")
wrap = Alignment(wrap_text=True, vertical="top")
thin = Side(style="thin", color="D9D9D9")


def D(s):
    return dt.datetime.strptime(s, "%Y-%m-%d")


def T(s):
    return dt.datetime.strptime(s, "%H:%M").time()


def day_of(folder, fname):
    return f"{os.path.basename(folder)}-{fname[:2]}"


def link(cell, sheet, ref):
    cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{sheet}'!{ref}", display=str(cell.value))


# ---------------- parse the tree ----------------
comms, teams_msgs = [], []
LINE = re.compile(r"^- (?:(\d\d:\d\d) )?\*\*(.*?)\*\*: (.*)$")
for cat in ("staff", "related", "alerts", "teams"):
    for p in sorted(glob.glob(f"{ROOT}/{cat}/*/*.md")):
        day = day_of(os.path.dirname(p), os.path.basename(p))
        src = ""
        if not (START <= D(day) <= END):
            continue
        for L in open(p, encoding="utf-8").read().splitlines():
            if L.startswith("## "):
                src = L[3:].strip()
                continue
            m = LINE.match(L)
            if not m:
                continue
            tm, who, text = m.groups()
            if cat == "teams":
                teams_msgs.append((day, tm or "00:00", src, who, text))
            elif cat == "alerts":
                comms.append(
                    (D(day), None, "Email", "Alert / system", ALERT_SENDER if who in ("", "-") else who, src, text, "Recipient (automated)")
                )
            else:
                mine = any(k in who for k in ME_MAIL) or who.startswith(DRAFT_PREFIX)
                inv = (
                    ("Organizer (calendar)" if any(text.startswith(x) for x in CAL_PREFIXES) else "Author (sent/draft)")
                    if mine
                    else "Recipient"
                )
                comms.append((D(day), None, "Email", "Staff" if cat == "staff" else "Related", who, text, "", inv))
for day, tm, chat, who, txt in teams_msgs:
    comms.append((D(day), T(tm), "Teams", "Chat", chat, who, txt, "Author" if who in ME_FIRST else "Participant"))
comms.sort(key=lambda r: (r[0], r[1] or dt.time(0, 0), r[2]))


# ---------------- Teams sessions (gap rule) ----------------
def mins(hhmm):
    return int(hhmm[:2]) * 60 + int(hhmm[3:])


sessions, by = [], {}
for day, tm, chat, who, txt in teams_msgs:
    by.setdefault((day, chat), []).append((tm, who))
for (day, chat), ms in sorted(by.items()):
    ms.sort()
    cur = [ms[0]]

    def flush(c):
        if any(x[1] in ME_FIRST for x in c):
            sessions.append((D(day), chat, T(c[0][0]), T(c[-1][0]), len(c), sum(1 for x in c if x[1] in ME_FIRST)))

    for m in ms[1:]:
        if mins(m[0]) - mins(cur[-1][0]) > GAP:
            flush(cur)
            cur = [m]
        else:
            cur.append(m)
    flush(cur)

# ---------------- meetings: expand series, apply the held rule ----------------
conv_dates = {D(m[0]) for m in teams_msgs}
meet = []
for sr in MEET_SERIES:
    grp = G_WORK if sr.get("group", "work") == "work" else G_UNC
    for ds in sr["dates"]:
        md = D(ds)
        status = sr.get("status")
        ev = sr.get("evidence", "Outlook calendar (recurring)")
        actual = sr.get("actual_min")
        if status is None:
            if grp != G_WORK:
                status = S_DEFAULT
            elif RULE == "teams_conversation_same_day":
                if md in conv_dates:
                    status = S_HELD_CARD if actual is not None else S_HELD_CONV
                else:
                    status = S_NOT_HELD
                    ev += "; rule: no Teams conversation that day = no meeting" + (
                        "; Teams card exists but is not counted" if actual is not None else ""
                    )
            else:
                status = S_DEFAULT
        meet.append((md, sr["title"], grp, T(sr["start"]), T(sr["end"]), status, actual, ev))
meet.sort(key=lambda r: (r[0], r[3]))

# ---------------- workbook ----------------
wb = Workbook()
ws0 = wb.active
ws0.title = "Summary"
wc = wb.create_sheet("Communications")
wt = wb.create_sheet("Teams Sessions")
wtm = wb.create_sheet("Teams Messages")
wk = wb.create_sheet("Tasks")
wmt = wb.create_sheet("Meetings")
wm = wb.create_sheet("Time Spent")


def header(ws, cols, widths):
    for i, (c, w) in enumerate(zip(cols, widths), 1):
        x = ws.cell(row=1, column=i, value=c)
        x.font = white
        x.fill = HDR
        x.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


# Communications
header(
    wc,
    ["Date", "Time", "Channel", "Category", "From / Chat", "Subject / Author / Type", "Summary / Text", "My involvement"],
    [11, 9, 9, 15, 34, 40, 90, 20],
)
for r, row in enumerate(comms, 2):
    for c, v in enumerate(row, 1):
        x = wc.cell(row=r, column=c, value=v)
        x.font = base
        x.alignment = wrap
    wc.cell(row=r, column=1).number_format = "yyyy-mm-dd"
    wc.cell(row=r, column=2).number_format = "hh:mm"
NC = len(comms) + 1
wc.auto_filter.ref = f"A1:H{NC}"
wc.cell(
    row=NC + 2, column=1, value=f"Email times are not recorded (list view only). Teams times are {TZ}. Secrets are redacted. Source: {ROOT}"
).font = note

# Teams Sessions
header(wt, ["Date", "Chat", "Start", "End", "Messages", "My messages", "Elapsed (min)", "Billed (min)"], [11, 52, 8, 8, 10, 12, 13, 13])
wt["J1"] = "Gap that splits sessions (min), applied when the workbook was built"
wt["K1"] = GAP
wt["J2"] = "Padding per session (min)"
wt["K2"] = PAD
wt["J3"] = "Counted only for sessions in which I posted"
wt["J4"] = "Billed = elapsed raised to the minimum per lapse (Time Spent I3), rounded up to the increment (Time Spent I4)"
for c in ("J1", "J2", "J3", "J4"):
    wt[c].font = base
wt["K2"].font = blue
wt["K2"].fill = INP
wt["K1"].font = base
wt.column_dimensions["J"].width = 60
for r, s in enumerate(sessions, 2):
    for c, v in enumerate(s, 1):
        wt.cell(row=r, column=c, value=v).font = base
    wt.cell(row=r, column=1).number_format = "yyyy-mm-dd"
    wt.cell(row=r, column=3).number_format = "hh:mm"
    wt.cell(row=r, column=4).number_format = "hh:mm"
    wt.cell(row=r, column=7, value=f"=ROUND((D{r}-C{r})*1440,0)+$K$2").font = base
    wt.cell(row=r, column=8, value=f"=CEILING(MAX(G{r},'Time Spent'!$I$3),'Time Spent'!$I$4)").font = base
    wt.cell(row=r, column=7).number_format = "0"
    wt.cell(row=r, column=8).number_format = "0"
NS = len(sessions) + 1
wt.cell(row=NS + 2, column=1, value="Total").font = bold
for c, L in ((5, "E"), (6, "F"), (7, "G"), (8, "H")):
    wt.cell(row=NS + 2, column=c, value=f"=SUM({L}2:{L}{NS})").font = bold
wt.cell(
    row=NS + 4,
    column=1,
    value="Billed time = elapsed raised to the minimum billable lapse (Time Spent I3) and rounded up to the increment (I4). Elapsed is a proxy for effort: first to last message of a back-and-forth plus padding.",
).font = note

# Teams Messages
header(wtm, ["Date", "Time", "Chat", "Author", "Message", "Session (click)"], [11, 9, 40, 12, 130, 16])
tm_rows = sorted(teams_msgs, key=lambda m: (m[0], m[2], m[1]))
first_msg_row = {}
for ri, (day, tm, chat, who, txt) in enumerate(tm_rows, 2):
    sidx = next((si for si, sx in enumerate(sessions) if sx[0] == D(day) and sx[1] == chat and sx[2] <= T(tm) <= sx[3]), None)
    for c, v in enumerate([D(day), T(tm), chat, who, txt], 1):
        x = wtm.cell(row=ri, column=c, value=v)
        x.font = base
        x.alignment = wrap
        if who in ME_FIRST:
            x.fill = MINE
    wtm.cell(row=ri, column=1).number_format = "yyyy-mm-dd"
    wtm.cell(row=ri, column=2).number_format = "hh:mm"
    if sidx is not None:
        lc = wtm.cell(row=ri, column=6, value=f"Session row {sidx + 2}")
        lc.font = LINKF
        link(lc, "Teams Sessions", f"A{sidx + 2}")
        first_msg_row.setdefault(sidx, ri)
NTM = len(tm_rows) + 1
wtm.auto_filter.ref = f"A1:F{NTM}"
wtm.cell(
    row=NTM + 2,
    column=1,
    value="Every Teams message in the period for the included chats. Shaded rows are mine. Attachments and secrets are placeholders.",
).font = note
wt.cell(row=1, column=9, value="Messages (click)").font = white
wt.cell(row=1, column=9).fill = HDR
wt.column_dimensions["I"].width = 18
for si in range(len(sessions)):
    if si in first_msg_row:
        lc = wt.cell(row=si + 2, column=9, value="View messages")
        lc.font = LINKF
        link(lc, "Teams Messages", f"A{first_msg_row[si]}")

# Tasks
header(
    wk,
    [
        "#",
        "Task",
        "Source",
        "Requested by",
        "Raised",
        "Status",
        "Est. low (h)",
        "Est. high (h)",
        "Est. likely (h)",
        "Basis for estimate",
        "Evidence",
    ],
    [5, 70, 36, 24, 11, 34, 11, 11, 12, 50, 11],
)
for r, tk in enumerate(TASK_ROWS, 2):
    vals = [
        r - 1,
        tk["task"],
        tk.get("source", ""),
        tk.get("requested_by", ""),
        D(tk["raised"]) if tk.get("raised") else None,
        tk.get("status", ""),
        tk.get("low_h", 0),
        tk.get("high_h", 0),
        f"=AVERAGE(G{r}:H{r})",
        tk.get("basis", ""),
        tk.get("evidence", "Requested"),
    ]
    for c, v in enumerate(vals, 1):
        x = wk.cell(row=r, column=c, value=v)
        x.font = base
        x.alignment = wrap
    wk.cell(row=r, column=5).number_format = "yyyy-mm-dd"
    for c in (7, 8):
        wk.cell(row=r, column=c).font = blue
        wk.cell(row=r, column=c).fill = INP
    for c in (7, 8, 9):
        wk.cell(row=r, column=c).number_format = "0.0"
NT = len(TASK_ROWS) + 1
wk.cell(row=NT + 1, column=2, value="Total estimated effort (hours)").font = bold
for c, col in ((7, "G"), (8, "H"), (9, "I")):
    x = wk.cell(row=NT + 1, column=c, value=f"=SUM({col}2:{col}{max(NT, 2)})")
    x.font = bold
    x.number_format = "0.0"
wk.cell(
    row=NT + 3,
    column=2,
    value="Hour ranges are estimates from the scope described in the messages, not tracked time. Yellow cells are inputs. 'Suggested' rows were not requested by the client.",
).font = note
wk.auto_filter.ref = f"A1:K{max(NT, 2)}"
wk.freeze_panes = "C2"

# Meetings
MINB = "'Time Spent'!$I$3"
INCB = "'Time Spent'!$I$4"
header(
    wmt,
    [
        "Date",
        "Day",
        "Meeting",
        "Group",
        "Sched. start",
        "Sched. end",
        "Sched. min",
        "Status",
        "Teams actual (min)",
        "Evidence",
        "Counted min (logged)",
        "Best-case min (if held)",
        "Day minimum applies (logged)",
        "Day minimum applies (best case)",
        "Teams messages that day",
    ],
    [11, 6, 52, 24, 11, 11, 10, 34, 12, 80, 13, 14, 14, 14, 12],
)
for r, (md, title, grp, s_, e_, status, actual, ev) in enumerate(meet, 2):
    vals = [
        md,
        f'=TEXT(A{r},"ddd")',
        title,
        grp,
        s_,
        e_,
        f"=ROUND((F{r}-E{r})*1440,0)",
        status,
        actual,
        ev,
        f'=IF(AND(D{r}="{G_WORK}",LEFT(H{r},4)="Held",I{r}<>""),CEILING(MAX(I{r},{MINB}),{INCB}),0)',
        f'=IF(D{r}<>"{G_WORK}",0,IF(OR(H{r}="{S_CANCEL}",H{r}="{S_LIKELY}"),0,IF(I{r}<>"",CEILING(MAX(I{r},{MINB}),{INCB}),CEILING(MAX(G{r},{MINB}),{INCB}))))',
        f'=IF(AND(D{r}="{G_WORK}",LEFT(H{r},4)="Held"),1,0)',
        f"=IF(L{r}>0,1,0)",
        f"=COUNTIFS('Teams Messages'!$A$2:$A${NTM},A{r})",
    ]
    for c, v in enumerate(vals, 1):
        x = wmt.cell(row=r, column=c, value=v)
        x.font = base
        x.alignment = wrap
    wmt.cell(row=r, column=1).number_format = "yyyy-mm-dd"
    wmt.cell(row=r, column=5).number_format = "hh:mm"
    wmt.cell(row=r, column=6).number_format = "hh:mm"
    for c in (7, 9, 11, 12, 13, 14, 15):
        wmt.cell(row=r, column=c).number_format = "0"
    if actual is not None:
        wmt.cell(row=r, column=9).font = blue
        wmt.cell(row=r, column=9).fill = INP
NM = max(len(meet) + 1, 2)
wmt.auto_filter.ref = f"A1:O{NM}"
wmt.cell(
    row=NM + 2,
    column=3,
    value=f"Rule ({RULE}): a scheduled meeting counts as held only if there was Teams conversation that day (Teams Messages tab); otherwise it is marked '{S_NOT_HELD}', even when a Teams meeting card exists. Cancelled occurrences vanish from the calendar, so cancellations come from emails/chat. Times are {TZ}.",
).font = note
wmt.cell(
    row=NM + 3,
    column=3,
    value=f"Group '{G_UNC}' = recurring items that do not look like {ORG} work. They are excluded from totals unless the toggle on Time Spent (I1) is Yes.",
).font = note

days = [START + dt.timedelta(days=i) for i in range((END - START).days + 1)]
ND = len(days) + 1

# Time Spent
header(wm, ["Category", "Count", "Minutes each", "Total (min)", "Total (h)", "Basis"], [58, 9, 14, 13, 10, 100])
inputs = [
    ("H1", "Include unconfirmed items? (Yes/No)", "I1", INCL_UNCONF),
    ("H2", "Meeting basis (Logged / Best case)", "I2", BASIS),
    ("H3", "Minimum billable per lapse of time (min)", "I3", MIN_LAPSE),
    ("H4", "Billing increment: round each lapse up to (min)", "I4", INCREMENT),
    ("H7", "Minimum billed for a day with an actual meeting (min)", "I7", DAY_MIN),
]
for hl, label, il, val in inputs:
    wm[hl] = label
    wm[hl].font = base
    wm[il] = val
    wm[il].font = blue
    wm[il].fill = INP
wm["H5"] = "Billed minutes per email written"
wm["H5"].font = base
wm["I5"] = "=CEILING(MAX(C3,I3),I4)"
wm["I5"].font = base
wm["H6"] = (
    "A lapse = one Teams session, one email written, or one meeting. Each lapse is raised to the minimum (I3), then rounded up to the increment (I4). Set I3 to 0 and I4 to 1 to see actual minutes."
)
wm["H6"].font = note
wm.column_dimensions["H"].width = 50
rows = [
    (
        "Teams chat sessions I took part in",
        f"=COUNTA('Teams Sessions'!A2:A{max(NS, 2)})",
        None,
        f"='Teams Sessions'!H{NS + 2}",
        "Computed from message timestamps (see Teams Sessions), billed per session",
    ),
    (
        "Email messages I wrote or drafted",
        f'=COUNTIFS(Communications!C2:C{NC},"Email",Communications!H2:H{NC},"Author*")',
        EMAIL_ACTUAL,
        "=B3*$I$5",
        "Actual effort assumed per message (yellow); billed per message follows I3/I4",
    ),
    (
        "Held meetings (Teams card minutes where recorded)",
        f'=COUNTIFS(Meetings!K2:K{NM},">0")',
        None,
        f"=SUM(Meetings!K2:K{NM})",
        "Sum of 'Meeting ended' durations in Teams for held meetings; see Meetings tab",
    ),
    (
        "Meeting-day minimum top-up (a day with an actual meeting bills at least the day minimum, I7)",
        f"=COUNTIFS('Daily Time'!M2:M{ND},\">0\")",
        None,
        f"=SUM('Daily Time'!M2:M{ND})",
        "Held meetings only (scheduled and Teams conversation that day; not cancelled or postponed). Top-up = day minimum minus the meeting minutes already billed that day. Edit which meetings qualify on the Meetings tab, columns M and N.",
    ),
    (
        "Unconfirmed recurring items (only counted if toggle = Yes)",
        f'=COUNTIFS(Meetings!D2:D{NM},"Unconfirmed*",Meetings!A2:A{NM},"<="&DATE({END.year},{END.month},{END.day}))',
        None,
        f'=IF($I$1="Yes",SUMIFS(Meetings!G2:G{NM},Meetings!D2:D{NM},"Unconfirmed*",Meetings!A2:A{NM},"<="&DATE({END.year},{END.month},{END.day})),0)',
        "Scheduled minutes in the period; awaiting confirmation they are work",
    ),
]
for r, (cat, cnt, each, tot, basis) in enumerate(rows, 2):
    wm.cell(row=r, column=1, value=cat).font = base
    wm.cell(row=r, column=2, value=cnt).font = base
    if each is not None:
        x = wm.cell(row=r, column=3, value=each)
        x.font = blue
        x.fill = INP
    wm.cell(row=r, column=4, value=tot).font = base
    wm.cell(row=r, column=5, value=f"=D{r}/60").font = base
    wm.cell(row=r, column=6, value=basis).font = base
    wm.cell(row=r, column=4).number_format = "0"
    wm.cell(row=r, column=5).number_format = "0.0"
wm.cell(row=7, column=1, value="Total billed time (logged basis)").font = bold
wm.cell(row=7, column=4, value="=SUM(D2:D6)").font = bold
wm.cell(row=7, column=5, value="=D7/60").font = bold
wm.cell(
    row=8,
    column=1,
    value="Best-case meetings incl. day minimum (ignores the held rule: every scheduled, non-cancelled call counted as held)",
).font = base
wm.cell(row=8, column=2, value=f'=COUNTIFS(Meetings!L2:L{NM},">0")').font = base
wm.cell(row=8, column=4, value=f"=SUM(Meetings!L2:L{NM})+SUM('Daily Time'!P2:P{ND})").font = base
wm.cell(row=8, column=5, value="=D8/60").font = base
wm.cell(row=8, column=6, value="Upper bound only").font = base
wm.cell(row=9, column=1, value="Total using the best-case meeting figure").font = bold
wm.cell(row=9, column=4, value="=D2+D3+D8+D6").font = bold
wm.cell(row=9, column=5, value="=D9/60").font = bold
for rr in (7, 8, 9):
    wm.cell(row=rr, column=4).number_format = "0"
    wm.cell(row=rr, column=5).number_format = "0.0"
wm.cell(
    row=11,
    column=1,
    value="Not included: reading time, phone calls, and work done after a chat or meeting (see Tasks). Meeting transcripts/recaps were not retrieved. Estimates, not tracked time.",
).font = note

# Daily Time
wdt = wb.create_sheet("Daily Time")
header(
    wdt,
    [
        "Date",
        "Day",
        "='Time Spent'!A2",
        "='Time Spent'!A3",
        "='Time Spent'!A4",
        "='Time Spent'!A5",
        "='Time Spent'!A6",
        "Total (min)",
        "Total (h)",
        "",
        "helper: meeting min (logged)",
        "helper: qualifying meetings (logged)",
        "Top-up (logged)",
        "helper: meeting min (best)",
        "helper: qualifying meetings (best)",
        "Top-up (best)",
    ],
    [11, 6, 30, 28, 32, 40, 32, 12, 10, 3, 14, 14, 12, 14, 14, 12],
)
MA = f"Meetings!$A$2:$A${NM}"
for r, day in enumerate(days, 2):
    wdt.cell(row=r, column=1, value=day).number_format = "yyyy-mm-dd"
    wdt.cell(row=r, column=2, value=f'=TEXT(A{r},"ddd")')
    wdt.cell(row=r, column=3, value=f"=SUMIFS('Teams Sessions'!$H$2:$H${max(NS, 2)},'Teams Sessions'!$A$2:$A${max(NS, 2)},A{r})")
    wdt.cell(
        row=r,
        column=4,
        value=f'=COUNTIFS(Communications!$A$2:$A${NC},A{r},Communications!$C$2:$C${NC},"Email",Communications!$H$2:$H${NC},"Author*")*\'Time Spent\'!$I$5',
    )
    wdt.cell(row=r, column=5, value=f"=IF('Time Spent'!$I$2=\"Best case\",N{r},K{r})")
    wdt.cell(row=r, column=6, value=f"=IF('Time Spent'!$I$2=\"Best case\",P{r},M{r})")
    wdt.cell(
        row=r,
        column=7,
        value=f'=IF(\'Time Spent\'!$I$1="Yes",SUMIFS(Meetings!$G$2:$G${NM},Meetings!$D$2:$D${NM},"Unconfirmed*",{MA},A{r}),0)',
    )
    wdt.cell(row=r, column=8, value=f"=SUM(C{r}:G{r})")
    wdt.cell(row=r, column=9, value=f"=H{r}/60")
    wdt.cell(row=r, column=11, value=f"=SUMIFS(Meetings!$K$2:$K${NM},{MA},A{r})")
    wdt.cell(row=r, column=12, value=f"=SUMIFS(Meetings!$M$2:$M${NM},{MA},A{r})")
    wdt.cell(row=r, column=13, value=f"=IF(L{r}>0,MAX(0,'Time Spent'!$I$7-K{r}),0)")
    wdt.cell(row=r, column=14, value=f"=SUMIFS(Meetings!$L$2:$L${NM},{MA},A{r})")
    wdt.cell(row=r, column=15, value=f"=SUMIFS(Meetings!$N$2:$N${NM},{MA},A{r})")
    wdt.cell(row=r, column=16, value=f"=IF(O{r}>0,MAX(0,'Time Spent'!$I$7-N{r}),0)")
    for c in range(1, 17):
        if c != 10:
            wdt.cell(row=r, column=c).font = base
    for c in list(range(3, 9)) + list(range(11, 17)):
        wdt.cell(row=r, column=c).number_format = "0"
    wdt.cell(row=r, column=9).number_format = "0.0"
wdt.cell(row=ND + 1, column=1, value="Total").font = bold
for c in range(3, 10):
    L = get_column_letter(c)
    x = wdt.cell(row=ND + 1, column=c, value=f"=SUM({L}2:{L}{ND})")
    x.font = bold
    x.number_format = "0.0" if c == 9 else "0"
wdt.cell(row=ND + 2, column=1, value="Check vs Time Spent").font = bold
wdt.cell(row=ND + 2, column=8, value="=IF('Time Spent'!$I$2=\"Best case\",'Time Spent'!D9,'Time Spent'!D7)").font = bold
wdt.cell(row=ND + 2, column=9, value=f'=IF(ROUND(H{ND + 1}-H{ND + 2},0)=0,"ties","MISMATCH")').font = bold
wdt.cell(
    row=ND + 4,
    column=1,
    value="Each column is a formula over the same source tab as its Time Spent line item. The day-minimum top-up = Time Spent I7 minus that day's billed meeting minutes, only on days with a qualifying meeting (Meetings columns M and N). Columns K to P are helpers.",
).font = note
wdt.freeze_panes = "C2"

# Day Detail
detail = []
for i, sx in enumerate(sessions):
    detail.append(
        (
            sx[0],
            "Teams chat",
            f"{sx[1]} ({sx[4]} messages, {sx[5]} mine)",
            sx[2],
            sx[3],
            f"='Teams Sessions'!H{i + 2}",
            "Teams Messages",
            f"A{first_msg_row[i]}",
        )
    )
for i, row_ in enumerate(comms):
    if row_[2] == "Email" and str(row_[7]).startswith("Author"):
        detail.append((row_[0], "Email written", f"{row_[4]}: {row_[5]}", None, None, "='Time Spent'!$I$5", "Communications", f"A{i + 2}"))
for i, (md, title, grp, s_, e_, status, actual, ev) in enumerate(meet):
    r = i + 2
    if grp == G_WORK:
        detail.append(
            (
                md,
                "Meeting",
                f"{title} [{status}]",
                s_,
                e_,
                f"=IF('Time Spent'!$I$2=\"Best case\",Meetings!L{r},Meetings!K{r})",
                "Meetings",
                f"A{r}",
            )
        )
    elif md <= END:
        detail.append(
            (md, "Unconfirmed item", f"{title} [{status}]", s_, e_, f"=IF('Time Spent'!$I$1=\"Yes\",Meetings!G{r},0)", "Meetings", f"A{r}")
        )
cand = sorted({m_[0] for m_ in meet if m_[2] == G_WORK and m_[5] not in (S_CANCEL, S_LIKELY)})
for md_ in cand:
    if md_ in days:
        drow = days.index(md_) + 2
        detail.append(
            (
                md_,
                "Meeting-day minimum",
                "Top-up to the day minimum for a day with an actual meeting (0 unless the meeting qualifies)",
                None,
                None,
                f"='Daily Time'!F{drow}",
                "Daily Time",
                f"A{drow}",
            )
        )
detail.sort(key=lambda x: (x[0], x[3] or (dt.time(23, 59) if x[1] == "Meeting-day minimum" else dt.time(0, 0)), x[1]))
wdd = wb.create_sheet("Day Detail")
header(
    wdd,
    [
        "Date",
        "Day",
        "Type",
        "Event",
        "Start",
        "End",
        "Minutes counted",
        "Hours (decimal)",
        "Time (h:mm)",
        "Day total (h)",
        "Source (click)",
    ],
    [11, 6, 20, 90, 8, 8, 14, 12, 11, 12, 18],
)
first_row = {}
NDD = len(detail) + 1
for r, (dd_, kind, ev_, s_, e_, f, ssheet, sref) in enumerate(detail, 2):
    for c, v in enumerate([dd_, f'=TEXT(A{r},"ddd")', kind, ev_, s_, e_, f], 1):
        x = wdd.cell(row=r, column=c, value=v)
        x.font = base
        x.alignment = wrap
    wdd.cell(row=r, column=1).number_format = "yyyy-mm-dd"
    wdd.cell(row=r, column=5).number_format = "hh:mm"
    wdd.cell(row=r, column=6).number_format = "hh:mm"
    wdd.cell(row=r, column=7).number_format = "0"
    hc = wdd.cell(row=r, column=8, value=f"=G{r}/60")
    hc.font = base
    hc.number_format = "0.00"
    tc = wdd.cell(row=r, column=9, value=f'=TEXT(G{r}/1440,"h:mm")')
    tc.font = base
    tc.alignment = Alignment(horizontal="right")
    dc = wdd.cell(row=r, column=10, value=f"=SUMIFS($G$2:$G${NDD},$A$2:$A${NDD},A{r})/60")
    dc.font = bold
    dc.number_format = "0.00"
    lc = wdd.cell(row=r, column=11, value=f"{ssheet} row")
    lc.font = LINKF
    link(lc, ssheet, sref)
    first_row.setdefault(dd_, r)
wdd.cell(row=NDD + 1, column=1, value="Total").font = bold
x = wdd.cell(row=NDD + 1, column=7, value=f"=SUM(G2:G{NDD})")
x.font = bold
x.number_format = "0"
x = wdd.cell(row=NDD + 1, column=8, value=f"=G{NDD + 1}/60")
x.font = bold
x.number_format = "0.00"
x = wdd.cell(row=NDD + 1, column=9, value=f'=INT(G{NDD + 1}/60)&":"&TEXT(MOD(G{NDD + 1},60),"00")')
x.font = bold
x.alignment = Alignment(horizontal="right")
wdd.cell(row=NDD + 2, column=1, value="Check vs Daily Time").font = bold
wdd.cell(row=NDD + 2, column=7, value=f"='Daily Time'!H{ND + 1}").font = bold
wdd.cell(row=NDD + 2, column=11, value=f'=IF(ROUND(G{NDD + 1}-G{NDD + 2},0)=0,"ties","MISMATCH")').font = bold
wdd.cell(
    row=NDD + 4,
    column=1,
    value="Every row is a formula over its source tab and follows the Time Spent settings. Click a Source cell to jump to the underlying row; click a day on Calendar View to land here.",
).font = note
wdd.auto_filter.ref = f"A1:K{NDD}"
wdd.freeze_panes = "A2"
bl = wdd.cell(row=1, column=13, value="Back to Calendar View")
bl.font = LINKF
link(bl, "Calendar View", "B1")
wdd.column_dimensions["M"].width = 22
for r, day in enumerate(days, 2):
    if day in first_row:
        link(wdt.cell(row=r, column=1), "Day Detail", f"A{first_row[day]}")
        wdt.cell(row=r, column=1).font = LINKF

# Calendar View
wcv = wb.create_sheet("Calendar View")
wcv.sheet_view.showGridLines = False
for i, w in enumerate([3, 12, 12, 12, 12, 12, 12, 12, 13], 1):
    wcv.column_dimensions[get_column_letter(i)].width = w
wcv["B1"] = "Time spent per day (minutes)"
wcv["B1"].font = Font(name=F, size=14, bold=True)
wcv["B2"] = (
    "Click any day to jump to its events on Day Detail. Linked to Daily Time, which sums the Time Spent line items. Billing settings, basis and toggle are on the Time Spent tab."
)
wcv["B2"].font = note
wcv["B3"] = (
    "=\"Meeting basis: \"&'Time Spent'!I2&\"   |   Unconfirmed items: \"&'Time Spent'!I1&\"   |   Grand total: \"&TEXT('Daily Time'!H"
    + str(ND + 1)
    + '/60,"0.0")&" h"'
)
wcv["B3"].font = bold
FMT = f'0" {CAL_UNIT}";-0;;@'
HEAT = tuple(c.lstrip("#") for c in _brand.heat_scale(BR["accent"]))
row = 5
months = []
y, mo = START.year, START.month
while (y, mo) <= (END.year, END.month):
    months.append((y, mo))
    mo += 1
    if mo == 13:
        y, mo = y + 1, 1
for yy, mm in months:
    wcv.cell(row=row, column=2, value=dt.datetime(yy, mm, 1)).number_format = "mmmm yyyy"
    wcv.cell(row=row, column=2).font = Font(name=F, size=12, bold=True)
    row += 1
    vcells = []
    for i, dn in enumerate(["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Week"]):
        x = wcv.cell(row=row, column=2 + i, value=dn)
        x.font = white
        x.fill = HDR
        x.alignment = Alignment(horizontal="center")
    row += 1
    first_val = row + 1
    for wk_ in _cal.Calendar(firstweekday=6).monthdatescalendar(yy, mm):
        for i, dd in enumerate(wk_):
            c = 2 + i
            dcell = wcv.cell(row=row, column=c)
            vcell = wcv.cell(row=row + 1, column=c)
            inmonth = dd.month == mm
            if inmonth:
                dkey = dt.datetime(dd.year, dd.month, dd.day)
                dcell.value = dkey
                dcell.number_format = "d"
                vcell.value = f"=IFERROR(INDEX('Daily Time'!$H$2:$H${ND},MATCH({get_column_letter(c)}{row},'Daily Time'!$A$2:$A${ND},0)),0)"
                vcell.number_format = FMT
                vcells.append(vcell.coordinate)
                if dkey in first_row:
                    link(vcell, "Day Detail", f"A{first_row[dkey]}")
                    link(dcell, "Day Detail", f"A{first_row[dkey]}")
                elif dkey in days:
                    link(vcell, "Daily Time", f"A{days.index(dkey) + 2}")
                    link(dcell, "Daily Time", f"A{days.index(dkey) + 2}")
            dcell.font = Font(name=F, size=9, color="595959", underline="single" if inmonth else None)
            dcell.alignment = Alignment(horizontal="left")
            vcell.font = Font(name=F, size=11, bold=True)
            vcell.alignment = Alignment(horizontal="center")
            for rr in (row, row + 1):
                wcv.cell(row=rr, column=c).border = Border(
                    left=thin, right=thin, top=thin if rr == row else None, bottom=thin if rr == row + 1 else None
                )
        wtc = wcv.cell(row=row + 1, column=9, value=f"=SUM(B{row + 1}:H{row + 1})")
        wtc.font = bold
        wtc.number_format = FMT
        wtc.alignment = Alignment(horizontal="center")
        wcv.row_dimensions[row + 1].height = 24
        row += 2
    wcv.cell(row=row, column=8, value="Month total").font = bold
    mt = wcv.cell(row=row, column=9, value="=SUM(" + ",".join(vcells) + ")")
    mt.font = bold
    mt.number_format = f'0" {CAL_UNIT}"'
    wcv.cell(row=row + 1, column=8, value="(hours)").font = base
    mh = wcv.cell(row=row + 1, column=9, value=f"=I{row}/60")
    mh.font = bold
    mh.number_format = '0.0" h"'
    wcv.conditional_formatting.add(
        " ".join(f"B{r_}:H{r_}" for r_ in range(first_val, row, 2)),
        ColorScaleRule(
            start_type="num",
            start_value=0,
            start_color=HEAT[0],
            mid_type="num",
            mid_value=SCALE["mid_min"],
            mid_color=HEAT[1],
            end_type="num",
            end_value=SCALE["max_min"],
            end_color=HEAT[2],
        ),
    )
    row += 3
wcv.cell(
    row=row,
    column=2,
    value=f"Shading: white = none, light brand tint about {SCALE['mid_min']} min, full tint {SCALE['max_min']} min or more. See Daily Time for the split.",
).font = note

# Alert Rollup
wa = wb.create_sheet("Alert Rollup")
mcols = [dt.datetime(y_, m_, 1) for (y_, m_) in months]
header(wa, ["Alert type"] + [m_.strftime("%b %Y") for m_ in mcols] + ["Total"], [28] + [10] * len(mcols) + [9])
kinds = sorted({r[5] for r in comms if r[3] == "Alert / system"})
for r, k in enumerate(kinds, 2):
    wa.cell(row=r, column=1, value=k).font = base
    for c, m_ in enumerate(mcols, 2):
        last = _cal.monthrange(m_.year, m_.month)[1]
        da = (m_ - dt.datetime(1899, 12, 30)).days
        db = (m_.replace(day=last) - dt.datetime(1899, 12, 30)).days
        wa.cell(
            row=r,
            column=c,
            value=f'=COUNTIFS(Communications!$F$2:$F${NC},$A{r},Communications!$D$2:$D${NC},"Alert / system",Communications!$A$2:$A${NC},">={da}",Communications!$A$2:$A${NC},"<={db}")',
        ).font = base
    wa.cell(row=r, column=len(mcols) + 2, value=f"=SUM(B{r}:{get_column_letter(len(mcols) + 1)}{r})").font = base
NA = len(kinds) + 1
wa.cell(row=NA + 1, column=1, value="Total").font = bold
for c in range(2, len(mcols) + 3):
    L = get_column_letter(c)
    wa.cell(row=NA + 1, column=c, value=f"=SUM({L}2:{L}{max(NA, 2)})").font = bold

# Summary
ws0.column_dimensions["A"].width = 52
ws0.column_dimensions["B"].width = 14
ws0.column_dimensions["C"].width = 70
eyebrow = BR.get("eyebrow", "")
ws0["A1"] = eyebrow
ws0["A1"].font = Font(name=F, size=9, bold=True, color=HDR_COLOR)
ws0["A2"] = f"{ORG} time tracking, {START:%Y-%m-%d} to {END:%Y-%m-%d}"
ws0["A2"].font = Font(name=F, size=14, bold=True)
ws0["A3"] = BR.get("tagline") or "Billable time from Outlook, Teams and calendar evidence. Tab guide in column C."
ws0["A3"].font = Font(name=F, size=10, italic=True, color=BR["light"]["ink-3"].lstrip("#"))
_col = 4
for _logo in [p for p in (BR.get("logo_path"), CLIENT_LOGO) if p]:  # brand mark first, then the client's mark
    try:
        from openpyxl.drawing.image import Image as _XLImage

        _img = _XLImage(_logo)
        _ratio = _img.height / _img.width if _img.width else 1
        _img.height = 56
        _img.width = max(1, int(56 / _ratio))
        ws0.add_image(_img, f"{get_column_letter(_col)}1")
        _col += _img.width // 64 + 2
    except Exception as exc:  # Pillow missing, SVG or unreadable image: the workbook is still valid without the logo
        print("logo not embedded:", os.path.basename(_logo), exc, file=sys.stderr)
S0 = 6
for _c, _h in enumerate(("Metric", "Value", "Where to look"), 1):
    _x = ws0.cell(row=S0 - 1, column=_c, value=_h)
    _x.font = white
    _x.fill = HDR
items = [
    ("My billed time (h), logged basis", "='Time Spent'!E7", "Time Spent tab; Calendar View shows it per day"),
    ("  same, best case (ignores the held rule) (h)", "='Time Spent'!E9", "Upper bound only"),
    ("Days with an actual meeting (day minimum applies)", "='Time Spent'!B5", "Meetings tab: held meetings only"),
    ("Teams sessions I took part in", "='Time Spent'!B2", "Teams Sessions tab; every message on Teams Messages"),
    ("Emails I wrote or drafted", "='Time Spent'!B3", "Communications tab"),
    ("Estimated task effort, likely (h)", f"=Tasks!I{NT + 1}", "Range on Tasks tab; estimates, not tracked time"),
    ("Tasks identified", f"=COUNTA(Tasks!B2:B{max(NT, 2)})", "Tasks tab"),
    ("  of which open or unknown status", f'=COUNTIFS(Tasks!F2:F{max(NT, 2)},"Open*")+COUNTIFS(Tasks!F2:F{max(NT, 2)},"Unknown*")', ""),
    ("Evidence: email from the organization (staff)", f'=COUNTIFS(Communications!C2:C{NC},"Email",Communications!D2:D{NC},"Staff")', "Communications tab"),
    (
        "Evidence: related email (not from the domain)",
        f'=COUNTIFS(Communications!C2:C{NC},"Email",Communications!D2:D{NC},"Related")',
        "Communications tab",
    ),
    (
        "Evidence: automated alerts / system mail",
        f'=COUNTIFS(Communications!C2:C{NC},"Email",Communications!D2:D{NC},"Alert / system")',
        "Alert Rollup tab by type and month",
    ),
    ("Evidence: Teams messages", f'=COUNTIFS(Communications!C2:C{NC},"Teams")', "Teams Messages tab"),
    ("Evidence: Teams messages written by me", f'=COUNTIFS(Communications!C2:C{NC},"Teams",Communications!H2:H{NC},"Author")', ""),
]
for r, (a_, f_, n_) in enumerate(items, S0):
    ws0.cell(row=r, column=1, value=a_).font = base
    ws0.cell(row=r, column=2, value=f_).font = bold
    ws0.cell(row=r, column=3, value=n_).font = base
    if "(h)" in a_:
        ws0.cell(row=r, column=2).number_format = "0.0"
ws0.cell(row=S0 + len(items) + 1, column=1, value="Notes").font = bold
for r, n in enumerate(
    [
        f"Billing parameters (Time Spent tab): minimum per lapse {MIN_LAPSE} min, increment {INCREMENT} min, meeting-day minimum {DAY_MIN} min, assumed {EMAIL_ACTUAL} min per email. Yellow cells are inputs; all totals are formulas.",
        "Evidence rows (email, Teams) support the time figures above; email times are not captured, so email dates are day-level only.",
        "A meeting counts as held only if there was Teams conversation that day (rule parameter in the config).",
        "Secrets found in chats were redacted; rotate them if still live.",
    ],
    S0 + len(items) + 2,
):
    ws0.cell(row=r, column=1, value=n).font = base

ORDER = ["Summary", "Time Spent", "Calendar View", "Daily Time", "Day Detail", "Meetings", "Teams Sessions", "Teams Messages", "Tasks", "Communications", "Alert Rollup"]
wb._sheets.sort(key=lambda w: ORDER.index(w.title) if w.title in ORDER else len(ORDER))
footer = BR.get("footer") or ""
for _ws in wb.worksheets:
    _ws.sheet_properties.tabColor = HDR_COLOR
    if footer:
        _ws.oddFooter.left.text = footer
    _ws.oddFooter.right.text = "Page &P of &N"
wb.save(OUT)
print(f"saved {OUT}: {len(sessions)} Teams sessions, {len(meet)} meeting rows, {len(TASK_ROWS)} tasks, {len(comms)} evidence rows")
