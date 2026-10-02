#!/usr/bin/env python3
"""Build the org-comms-export tree from a flat text file.

Input line format (one record per line, UTF-8; blank lines and lines starting with # are ignored):
    YYYY-MM-DD|HH:MM|category|source|who|text
  category: staff | related | alerts | teams
  source:   for teams the chat name, for mail anything short (e.g. "mail")
  who:      sender / author display name
  text:     subject and/or message text; use "|" sparingly (extra pipes stay in text)
  HH:MM may be empty when unknown (mail list view has no time for older mail).

Usage:
  python build_tree.py --in lines.txt --out ./acme --start YYYY-MM-DD [--tz "America/New_York"]
         [--patterns extra_patterns.txt] [--placeholder "[REDACTED]"] [--categories staff,related,alerts,teams,meetings]
  --patterns     optional file, one regex per line (# comments allowed), applied in addition to the built-in credential patterns
  --placeholder  text that replaces a redacted token
  --categories   allowed category names (anything else is skipped with a warning)
"""
import argparse
import collections
import os
import re
import sys

# Tokens that look like passwords / secrets: mixed letters+digits with a symbol, long hex/GUID-like secrets
# are NOT redacted (client IDs are not secret), but Azure-style secret values (xxx~yyyy) and common key prefixes are.
PATTERNS = [
    re.compile(r"(?<!\S)(?=\S*\d)(?=\S*[A-Za-z])(?=\S*[$#%^&*!])\S{9,}"),
    re.compile(r"\b[A-Za-z0-9_-]{3,}~[A-Za-z0-9_.~-]{20,}"),
    re.compile(r"\b(?:sk|pk|ghp|gho|xox[abp]|AKIA)[-_A-Za-z0-9]{16,}"),
    re.compile(r"(?i)\b(password|passwd|pwd|secret|token|api[_ -]?key)\b\s*[:=]\s*\S+"),
]

def redact(t, patterns, placeholder):
    n = 0
    for p in patterns:
        t, k = p.subn(placeholder, t)
        n += k
    return t, n

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--in", dest="inp", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--start", required=True)
    ap.add_argument("--tz", default="")
    ap.add_argument("--patterns", default="")
    ap.add_argument("--placeholder", default="[REDACTED]")
    ap.add_argument("--categories", default="staff,related,alerts,teams,meetings")
    a = ap.parse_args()
    patterns = list(PATTERNS)
    if a.patterns:
        for ln in open(a.patterns, encoding="utf-8").read().splitlines():
            if ln.strip() and not ln.startswith("#"):
                patterns.append(re.compile(ln.strip()))
    allowed = {c.strip() for c in a.categories.split(",") if c.strip()}

    by = collections.defaultdict(lambda: collections.defaultdict(list))
    total = redacted = dropped = 0
    for raw in open(a.inp, encoding="utf-8").read().splitlines():
        if not raw.strip() or raw.startswith("#"):
            continue
        parts = raw.split("|", 5)
        if len(parts) < 6 or not re.match(r"\d{4}-\d\d-\d\d$", parts[0]):
            print("skip malformed:", raw[:80], file=sys.stderr)
            continue
        d, tm, cat, src, who, text = [p.strip() for p in parts]
        if d < a.start:
            dropped += 1
            continue
        if cat not in allowed:
            print("skip unknown category:", cat, file=sys.stderr)
            continue
        text, n = redact(text, patterns, a.placeholder)
        redacted += bool(n)
        key = (cat, d)
        by[key][src].append(f"- {tm + ' ' if tm else ''}**{who}**: {text}" if cat == "teams" or who else f"- {text}")
        total += 1

    counts = collections.Counter()
    for (cat, d), srcs in sorted(by.items()):
        p = os.path.join(a.out, cat, d[:7], d[8:] + ".md")
        os.makedirs(os.path.dirname(p), exist_ok=True)
        tz = f" (times {a.tz})" if a.tz and cat == "teams" else ""
        body = "".join(f"## {s}\n" + "\n".join(lines) + "\n\n" for s, lines in srcs.items())
        with open(p, "w", encoding="utf-8") as f:
            f.write(f"# {cat} - {d}{tz}\n\n{body}")
        counts[cat] += 1

    rp = os.path.join(a.out, "README.md")
    if not os.path.exists(rp):
        with open(rp, "w", encoding="utf-8") as f:
            f.write(f"# Export from {a.start} on\n\nFill in: source, pull date, caveats, excluded chats, redactions.\n")
    print(f"records written: {total}; dropped before {a.start}: {dropped}; records with redactions: {redacted}")
    print("day files per category:", dict(counts))

if __name__ == "__main__":
    main()
