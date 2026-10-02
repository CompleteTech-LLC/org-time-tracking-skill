#!/usr/bin/env python3
"""Render a self-contained, branded light/dark time-tracking HTML report from a RECALCULATED org-time-tracking workbook.

Usage:
  python build_report.py --config config.json [--workbook book.xlsx] [--out report.html] [--theme auto|light|dark]

The workbook must carry cached values (recalculate it in Excel or LibreOffice first: scripts/recalc_excel.ps1).

The page has From / To dropdowns (and a quick month picker) that recompute every figure, the calendar, the meetings and
the tasks for the chosen range in the browser; a link can preselect a range with ?from=YYYY-MM-DD&to=YYYY-MM-DD. The
report is a single HTML file: no external fonts, scripts or images; the logo (local file only) is inlined as a data URI;
the Content-Security-Policy allows no network access and permits exactly one inline script, pinned by its SHA-256 hash.
Branding is neutral unless a preset or explicit values are set (see scripts/branding.py and templates/branding.*.json).
"""

from __future__ import annotations

import argparse
import base64
import datetime as dt
import hashlib
import html
import json
import mimetypes
import os
import sys

from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import branding as brandlib  # noqa: E402

# Daily Time columns this report reads by position (see make_workbook.py): 3 Teams, 4 Email, 7 Unconfirmed,
# 11 meeting min (logged), 13 top-up (logged), 14 meeting min (best), 16 top-up (best).
DAILY = {"teams": 2, "email": 3, "unc": 6, "meetL": 10, "topL": 12, "meetB": 13, "topB": 15}


def esc(value) -> str:
    return html.escape("" if value is None else str(value))


def hm(minutes) -> str:
    minutes = int(round(minutes or 0))
    return f"{minutes // 60}:{minutes % 60:02d}"


def table(ws) -> list[dict]:
    """Rows of a tab whose first row is a header; stops at the first row with an empty first cell."""
    headers = [c.value for c in ws[1]]
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] in (None, ""):
            break
        rows.append({h: v for h, v in zip(headers, r) if h})
    return rows


def day(value) -> str:
    return value.date().isoformat() if isinstance(value, (dt.datetime, dt.date)) else str(value)[:10]


def css(brand: dict, accent_on: str) -> str:
    def block(tokens: dict) -> str:
        return "".join(f"--{k}:{v};" for k, v in tokens.items())

    font = brand.get("font") or "Arial"
    text_light = brandlib.readable_on(brand["accent"], brand["light"]["surface"])
    text_dark = brandlib.readable_on(brand["accent"], brand["dark"]["surface"])
    return (
        f":root{{{block(brand['light'])}--accent:{brand['accent']};--on-accent:{accent_on};--accent-text:{text_light};}}"
        f'@media (prefers-color-scheme:dark){{:root:not([data-theme="light"]){{{block(brand["dark"])}--accent-text:{text_dark};}}}}'
        f':root[data-theme="dark"]{{{block(brand["dark"])}--accent-text:{text_dark};}}'
        f"*{{box-sizing:border-box}}body{{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 '{esc(font)}',system-ui,sans-serif}}"
        "main{max-width:1080px;margin:0 auto;padding:24px 16px 48px}"
        "header.brand{display:flex;gap:16px;align-items:center;padding:20px 0;border-bottom:3px solid var(--accent);margin-bottom:16px}"
        "header.brand img{height:56px;width:auto}.eyebrow{font-size:12px;letter-spacing:.12em;font-weight:700;color:var(--accent-text)}"
        "h1{margin:2px 0;font-size:26px}.tagline{color:var(--ink-3);margin:0}h2{margin:32px 0 10px;font-size:18px;color:var(--ink-2)}"
        "h3{margin:20px 0 8px;font-size:15px;color:var(--ink-2)}"
        ".filters{display:flex;gap:12px;align-items:flex-end;flex-wrap:wrap;background:var(--surface);border:1px solid var(--line);"
        "border-radius:10px;padding:12px 14px;margin-bottom:16px;position:sticky;top:0;z-index:5}"
        ".filters label{display:flex;flex-direction:column;font-size:12px;color:var(--ink-3);gap:2px}"
        ".filters select,.filters button{font:inherit;font-size:14px;color:var(--ink);background:var(--surface-2);border:1px solid var(--line-2);"
        "border-radius:6px;padding:6px 8px}.filters button{cursor:pointer;color:var(--on-accent);background:var(--accent);border-color:var(--accent)}"
        ".filters select:focus-visible,.filters button:focus-visible{outline:2px solid var(--accent-text);outline-offset:2px}"
        ".filters .range{margin-left:auto;color:var(--ink-3);font-size:13px}"
        ".tiles{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px}"
        ".tile{background:var(--surface);border:1px solid var(--line);border-radius:10px;padding:14px}"
        ".tile b{display:block;font-size:26px;color:var(--accent-text)}.tile span{color:var(--ink-3);font-size:13px}"
        "table{width:100%;border-collapse:collapse;background:var(--surface);border:1px solid var(--line);border-radius:8px;overflow:hidden}"
        "th{background:var(--accent);color:var(--on-accent);text-align:left;padding:8px 10px;font-size:13px}"
        "td{padding:8px 10px;border-top:1px solid var(--line);vertical-align:top;font-size:14px}tr:nth-child(even) td{background:var(--surface-2)}"
        "td.n,th.n{text-align:right;white-space:nowrap}tr.total td{font-weight:700}"
        ".cal{display:grid;grid-template-columns:repeat(7,1fr);gap:4px;margin-bottom:8px}"
        ".cal div{border:1px solid var(--line);border-radius:6px;min-height:46px;padding:4px 6px;background:var(--surface);font-size:12px}"
        ".cal .dow{background:var(--surface-2);min-height:0;text-align:center;font-weight:700;color:var(--ink-2)}"
        ".cal .d{color:var(--ink-3);display:block;font-size:11px}.cal .m{font-weight:700;font-size:13px}"
        ".cal .hot{color:var(--on-accent)}.cal .hot .d{color:inherit;opacity:.8}.cal .blank{border:none;background:none}"
        ".cal .out{opacity:.35}"
        ".note{color:var(--ink-3);font-size:13px}footer{margin-top:40px;padding-top:12px;border-top:1px solid var(--line);color:var(--ink-3);font-size:12px}"
        "@media print{body{background:#fff;color:#000}.filters{display:none}.tile,table,.cal div{break-inside:avoid}}"
    )


# The only script in the page. Everything it shows comes from the embedded DATA object; all text is escaped.
SCRIPT = r"""
(function () {
  "use strict";
  var D = __DATA__, days = D.days, SCALE = D.scale || 240;
  var DOW = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  var MON = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  function $(id) { return document.getElementById(id); }
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
  function pad(n) { return (n < 10 ? "0" : "") + n; }
  function hm(m) { m = Math.round(m || 0); return Math.floor(m / 60) + ":" + pad(m % 60); }
  function one(n) { return (Math.round(n * 10) / 10).toFixed(1); }
  function parts(d) { var p = d.split("-"); return [+p[0], +p[1], +p[2]]; }
  function dow(d) { var p = parts(d); return new Date(p[0], p[1] - 1, p[2]).getDay(); }
  function sum(rows, k) { var t = 0; rows.forEach(function (x) { t += x[k] || 0; }); return t; }
  if (!days.length) { return; }
  var fromSel = $("from"), toSel = $("to"), monthSel = $("month"), byDate = {}, monthKeys = [];
  days.forEach(function (x) {
    byDate[x.d] = x;
    var k = x.d.slice(0, 7); if (monthKeys.indexOf(k) < 0) { monthKeys.push(k); }
    [fromSel, toSel].forEach(function (s) { var o = document.createElement("option"); o.value = x.d; o.textContent = x.d + " " + DOW[dow(x.d)]; s.appendChild(o); });
  });
  monthSel.innerHTML = '<option value="">Whole period</option>' + monthKeys.map(function (k) { var p = k.split("-"); return '<option value="' + k + '">' + MON[+p[1] - 1] + " " + p[0] + "</option>"; }).join("");
  var first = days[0].d, last = days[days.length - 1].d;
  fromSel.value = first; toSel.value = last;
  try {
    var q = new URLSearchParams(location.search);
    if (byDate[q.get("from")]) { fromSel.value = q.get("from"); }
    if (byDate[q.get("to")]) { toSel.value = q.get("to"); }
  } catch (e) { /* no URL parameters available */ }
  $("filters").hidden = false;

  function inRange(d, f, t) { return d >= f && d <= t; }
  function render() {
    var f = fromSel.value, t = toSel.value;
    if (f > t) { var swap = f; f = t; t = swap; fromSel.value = f; toSel.value = t; }
    var sel = days.filter(function (x) { return inRange(x.d, f, t); });
    var meet = D.meetings.filter(function (m) { return inRange(m.d, f, t); });
    var held = meet.filter(function (m) { return m.kind === "held"; }), skipped = meet.filter(function (m) { return m.kind === "skipped"; });
    var tasks = D.tasks.filter(function (k) { return !k.raised || inRange(k.raised, f, t); });
    var teams = sum(sel, "teams"), email = sum(sel, "email"), meetL = sum(sel, "meetL"), topL = sum(sel, "topL"), unc = sum(sel, "unc");
    var meetB = sum(sel, "meetB"), topB = sum(sel, "topB");
    var logged = teams + email + meetL + topL + unc, best = teams + email + meetB + topB + unc;
    var topDays = sel.filter(function (x) { return x.topL > 0; }).length, bestDays = sel.filter(function (x) { return x.topB > 0; }).length;
    var effort = 0; tasks.forEach(function (k) { effort += k.likely || 0; });
    $("range").textContent = f + " to " + t + " · " + sel.length + " days";
    var tiles = [
      [one(logged / 60), "My billed time (h), logged basis"], [one(best / 60), "same, best case (ignores the held rule) (h)"],
      [topDays, "Days with an actual meeting (day minimum applies)"], [sum(sel, "sessions"), "Teams sessions I took part in"],
      [sum(sel, "emails"), "Emails I wrote or drafted"], [one(effort), "Estimated task effort, likely (h)"],
      [tasks.length, "Tasks identified"], [sum(sel, "staff"), "Evidence: email from the organization (staff)"],
      [sum(sel, "related"), "Evidence: related email (not from the domain)"], [sum(sel, "tmsg"), "Evidence: Teams messages"]
    ];
    $("tiles").innerHTML = tiles.map(function (x) { return '<div class="tile"><b>' + esc(x[0]) + "</b><span>" + esc(x[1]) + "</span></div>"; }).join("");
    function row(label, count, min, cls) { return '<tr class="' + (cls || "") + '"><td>' + esc(label) + '</td><td class="n">' + esc(count) + '</td><td class="n">' + Math.round(min) + '</td><td class="n">' + hm(min) + "</td></tr>"; }
    $("billed").innerHTML = '<h2>Billed time</h2><table><tr><th>Line item</th><th class="n">Count</th><th class="n">Minutes</th><th class="n">h:mm</th></tr>' +
      row(D.labels.teams, sum(sel, "sessions"), teams) + row(D.labels.email, sum(sel, "emails"), email) +
      row(D.labels.meet, held.length, meetL) + row(D.labels.topup, topDays, topL) + row(D.labels.unc, "", unc) +
      row("Total billed time (logged basis)", "", logged, "total") +
      row("Best-case meetings incl. day minimum (ignores the held rule)", bestDays, meetB + topB) + row("Total using the best-case meeting figure", "", best, "total") + "</table>";
    var cal = "";
    monthKeys.forEach(function (k) {
      var p = k.split("-"), y = +p[0], m = +p[1] - 1, count = new Date(y, m + 1, 0).getDate(), lead = new Date(y, m, 1).getDay(), mt = 0, cells = "", any = false;
      for (var i = 0; i < lead; i++) { cells += '<div class="blank"></div>'; }
      for (var d = 1; d <= count; d++) {
        var key = k + "-" + pad(d), x = byDate[key], ok = inRange(key, f, t), v = 0;
        if (ok) { any = true; }
        if (ok && x) { v = x.teams + x.email + x.meetL + x.topL + x.unc; mt += v; }
        var pct = SCALE ? Math.min(100, Math.round(100 * v / SCALE)) : 0;
        cells += '<div class="c' + (pct >= 55 ? " hot" : "") + (ok ? "" : " out") + '"' + (v ? ' style="background:color-mix(in srgb,var(--accent) ' + pct + '%,var(--surface))"' : "") + '><span class="d">' + d + '</span><span class="m">' + (v ? Math.round(v) + " m" : "") + "</span></div>";
      }
      if (any) { cal += "<h3>" + MON[m] + " " + y + " · " + hm(mt) + " h:mm</h3><div class=\"cal\">" + DOW.map(function (w) { return '<div class="dow">' + w + "</div>"; }).join("") + cells + "</div>"; }
    });
    $("cal").innerHTML = "<h2>Time per day</h2>" + cal + '<p class="note">Shade = billed minutes that day relative to the configured maximum; days outside the chosen range are dimmed. See the workbook (Daily Time, Day Detail) for the events.</p>';
    $("held").innerHTML = "<h2>Meetings counted as held (" + held.length + ')</h2><table><tr><th>Date</th><th>Meeting</th><th>Status</th><th class="n">Teams card (min)</th><th>Evidence</th></tr>' +
      held.map(function (m) { return "<tr><td>" + esc(m.d) + "</td><td>" + esc(m.title) + "</td><td>" + esc(m.status) + '</td><td class="n">' + esc(m.card) + "</td><td>" + esc(m.evidence) + "</td></tr>"; }).join("") + "</table>";
    $("skipped").innerHTML = "<h2>Scheduled but not counted (" + skipped.length + ')</h2><table><tr><th>Date</th><th>Meeting</th><th>Status</th></tr>' +
      skipped.map(function (m) { return "<tr><td>" + esc(m.d) + "</td><td>" + esc(m.title) + "</td><td>" + esc(m.status) + "</td></tr>"; }).join("") +
      '</table><p class="note">With the default held-meeting rule, a scheduled call with no Teams conversation that day is not counted, even if the calendar shows it.</p>';
    $("tasks").innerHTML = '<h2>Tasks and effort estimates</h2><table><tr><th>Task</th><th>Status</th><th class="n">Low h</th><th class="n">High h</th><th class="n">Likely h</th></tr>' +
      tasks.map(function (k) { return "<tr><td>" + esc(k.task) + "</td><td>" + esc(k.status) + '</td><td class="n">' + esc(k.low) + '</td><td class="n">' + esc(k.high) + '</td><td class="n">' + one(k.likely || 0) + "</td></tr>"; }).join("") +
      '</table><p class="note">Estimates from the scope described in the messages, not tracked time. Tasks are filtered by the date they were raised.</p>';
    monthSel.value = "";
    monthKeys.forEach(function (k) { var ds = days.filter(function (x) { return x.d.slice(0, 7) === k; }); if (ds[0].d === f && ds[ds.length - 1].d === t) { monthSel.value = k; } });
    try { var u = new URLSearchParams(); if (f !== first || t !== last) { u.set("from", f); u.set("to", t); } history.replaceState(null, "", location.pathname + (u.toString() ? "?" + u.toString() : "")); } catch (e) { /* file URLs may refuse */ }
  }
  fromSel.addEventListener("change", render);
  toSel.addEventListener("change", render);
  monthSel.addEventListener("change", function () {
    if (!monthSel.value) { fromSel.value = first; toSel.value = last; } else {
      var ds = days.filter(function (x) { return x.d.slice(0, 7) === monthSel.value; });
      fromSel.value = ds[0].d; toSel.value = ds[ds.length - 1].d;
    }
    render();
  });
  $("reset").addEventListener("click", function () { fromSel.value = first; toSel.value = last; render(); });
  render();
})();
"""


def collect(wb, cfg: dict) -> dict:
    """Pull the per-day figures, meetings and tasks out of the recalculated workbook (all values, no formulas)."""
    labels_cfg = cfg.get("meetings", {}).get("labels", {})
    not_counted = tuple(
        labels_cfg.get(k, d)
        for k, d in (("not_held", "Not held (no conversation that day)"), ("likely_not_held", "Likely not held (chat)"), ("cancelled", "Cancelled"))
    )
    dws = wb["Daily Time"]
    if not str(dws.cell(row=1, column=11).value or "").startswith("helper: meeting min (logged)"):
        raise SystemExit("Daily Time tab layout not recognised: rebuild the workbook with this version of make_workbook.py")
    daily: dict[str, dict] = {}
    for r in dws.iter_rows(min_row=2, values_only=True):
        if not isinstance(r[0], (dt.datetime, dt.date)):
            continue
        daily[day(r[0])] = {k: float(r[i] or 0) for k, i in DAILY.items()}
    for key in ("sessions", "emails", "staff", "related", "alerts", "tmsg", "tmine"):
        for v in daily.values():
            v[key] = 0
    for r in table(wb["Teams Sessions"]):
        if day(r.get("Date")) in daily:
            daily[day(r["Date"])]["sessions"] += 1
    if "Communications" in wb.sheetnames:
        for r in wb["Communications"].iter_rows(min_row=2, values_only=True):
            if not isinstance(r[0], (dt.datetime, dt.date)) or day(r[0]) not in daily:
                continue
            slot = daily[day(r[0])]
            if r[2] == "Teams":
                slot["tmsg"] += 1
                slot["tmine"] += 1 if r[7] == "Author" else 0
            elif r[2] == "Email":
                slot["emails"] += 1 if str(r[7]).startswith("Author") else 0
                if r[3] == "Staff":
                    slot["staff"] += 1
                elif r[3] == "Related":
                    slot["related"] += 1
                elif r[3] == "Alert / system":
                    slot["alerts"] += 1
    meetings = []
    for r in table(wb["Meetings"]):
        status = str(r.get("Status", ""))
        kind = "held" if status.startswith("Held") else "skipped" if status in not_counted else "other"
        if kind != "other":
            meetings.append({"d": day(r["Date"]), "title": r.get("Meeting"), "status": status, "kind": kind,
                             "card": r.get("Teams actual (min)"), "evidence": r.get("Evidence")})
    tasks = []
    for t in table(wb["Tasks"]):
        if isinstance(t.get("#"), (int, float)):
            tasks.append({"task": t.get("Task"), "status": t.get("Status"), "raised": day(t["Raised"]) if t.get("Raised") else "",
                          "low": t.get("Est. low (h)"), "high": t.get("Est. high (h)"), "likely": t.get("Est. likely (h)") or 0})
    ts = wb["Time Spent"]
    labels = {"teams": ts["A2"].value, "email": ts["A3"].value, "meet": ts["A4"].value, "topup": ts["A5"].value, "unc": ts["A6"].value}
    scale = (cfg.get("style", {}).get("calendar_color_scale") or {}).get("max_min", 240)
    return {"days": [{"d": k, **v} for k, v in sorted(daily.items())], "meetings": meetings, "tasks": tasks, "labels": labels, "scale": scale}


def build(cfg: dict, workbook: str, theme: str, config_dir: str) -> str:
    brand = brandlib.resolve(cfg.get("branding"), config_dir)
    accent_on = brandlib.on_colour(brand["accent"])
    wb = load_workbook(workbook, data_only=True)
    for name in ("Summary", "Time Spent", "Daily Time", "Meetings", "Tasks", "Teams Sessions"):
        if name not in wb.sheetnames:
            raise SystemExit(f"workbook has no '{name}' tab: build it with scripts/make_workbook.py")
    ts = wb["Time Spent"]
    if ts["D7"].value is None:
        raise SystemExit("workbook has no cached values: recalculate it first (scripts/recalc_excel.ps1 or LibreOffice)")

    data = collect(wb, cfg)
    org = cfg.get("org", {}).get("name", "")
    period = cfg.get("period", {})
    cfg_theme = theme or cfg.get("style", {}).get("theme", "auto")
    theme_attr = f' data-theme="{esc(cfg_theme)}"' if cfg_theme in ("light", "dark") else ""

    logo = ""
    if brand.get("logo_path"):
        mime = mimetypes.guess_type(brand["logo_path"])[0] or "image/png"
        with open(brand["logo_path"], "rb") as handle:
            blob = base64.b64encode(handle.read()).decode("ascii")
        logo = f'<img src="data:{mime};base64,{blob}" alt="{esc(brand.get("name") or "logo")} logo">'

    # JSON for a <script> block: neutralize anything that could close the tag or start an HTML comment
    payload = json.dumps(data, separators=(",", ":"), default=str).replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
    script = SCRIPT.replace("__DATA__", payload)
    digest = base64.b64encode(hashlib.sha256(script.encode("utf-8")).digest()).decode("ascii")
    csp = f"default-src 'none'; style-src 'unsafe-inline'; img-src data:; script-src 'sha256-{digest}'"

    footer = brand.get("footer") or ""
    contact = brand.get("contact") or ""
    total_logged = ts["D7"].value or 0
    out = [
        f'<!doctype html><html lang="en"{theme_attr}><head><meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width,initial-scale=1">',
        f'<meta http-equiv="Content-Security-Policy" content="{csp}">',
        f"<title>{esc(org)} time tracking report</title><style>{css(brand, accent_on)}</style></head><body><main>",
        '<header class="brand">' + logo + "<div>" + (f'<div class="eyebrow">{esc(brand["eyebrow"])}</div>' if brand.get("eyebrow") else "")
        + f'<h1>{esc(org)} time tracking</h1><p class="tagline">{esc(brand.get("tagline") or "")}'
        + f" {esc(period.get('start', ''))} to {esc(period.get('end', ''))}</p></div></header>",
        '<form class="filters" id="filters" hidden onsubmit="return false">'
        '<label>From<select id="from"></select></label><label>To<select id="to"></select></label>'
        '<label>Quick range<select id="month"></select></label><button type="button" id="reset">Reset</button>'
        '<span class="range" id="range" aria-live="polite"></span></form>',
        '<section class="tiles" id="tiles"></section><div id="billed"></div><div id="cal"></div><div id="held"></div><div id="skipped"></div><div id="tasks"></div>',
        f"<noscript><p class=\"note\">Turn on JavaScript to filter by date. Total billed time for the whole period (logged basis): {hm(total_logged)} h:mm.</p></noscript>",
        f"<footer>{esc(footer)}{' · ' if footer and contact and contact not in footer else ''}{esc(contact if contact not in footer else '')}"
        f"<br>Generated {dt.date.today().isoformat()} from the recalculated workbook. Review before sharing: this report contains client names and activity.</footer>",
        f"</main><script>{script}</script></body></html>",
    ]
    return "".join(out)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", required=True)
    ap.add_argument("--workbook", default="")
    ap.add_argument("--out", default="")
    ap.add_argument("--theme", default="", choices=["", "auto", "light", "dark"])
    a = ap.parse_args()
    config_dir = os.path.dirname(os.path.abspath(a.config))
    with open(a.config, encoding="utf-8") as handle:
        cfg = json.load(handle)

    def path(p: str) -> str:
        return p if os.path.isabs(p) else os.path.join(config_dir, p)

    workbook = path(a.workbook or cfg["output_xlsx"])
    out = path(a.out or os.path.splitext(cfg["output_xlsx"])[0] + "_report.html")
    document = build(cfg, workbook, a.theme, config_dir)
    with open(out, "w", encoding="utf-8") as handle:
        handle.write(document)
    print(f"saved {out} ({len(document)} bytes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
