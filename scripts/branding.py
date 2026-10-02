#!/usr/bin/env python3
"""Resolve branding and theme tokens for the workbook and the HTML report (standard library only).

Identity is neutral unless the operator selects an approved one: a preset (templates/branding.<preset>.json, e.g.
"completetech") and/or explicit values in the config's "branding" block. Explicit values win over the preset.
No logo is ever fetched or synthesized: "logo" is a local image file, resolved against the config folder and then
the repository root.
"""
from __future__ import annotations

import json
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HEX = re.compile(r"^#?[0-9a-fA-F]{6}$")
TOKENS = ("bg", "surface", "surface-2", "ink", "ink-2", "ink-3", "line", "line-2")


def _load(path: str) -> dict:
    with open(path, encoding="utf-8") as handle:
        data = json.load(handle)
    return {k: v for k, v in data.items() if not k.startswith("_")}


def _merge(base: dict, extra: dict) -> dict:
    out = dict(base)
    for key, value in extra.items():
        if key.startswith("_"):
            continue
        if key in ("light", "dark") and isinstance(value, dict):
            out[key] = {**out.get(key, {}), **value}
        elif value not in (None, ""):
            out[key] = value
    return out


def resolve(cfg_branding: dict | None, config_dir: str = ".") -> dict:
    """Return the effective brand: neutral defaults, then preset, then explicit config values."""
    cfg_branding = cfg_branding or {}
    brand = _load(os.path.join(ROOT, "templates", "branding.neutral.json"))
    preset = cfg_branding.get("preset", "neutral")
    if preset and preset != "neutral":
        path = preset if os.path.isfile(preset) else os.path.join(ROOT, "templates", f"branding.{preset}.json")
        if not os.path.isfile(path):
            raise SystemExit(f"unknown branding preset: {preset} (looked for {path})")
        brand = _merge(brand, _load(path))
    brand = _merge(brand, cfg_branding)
    brand["preset"] = preset or "neutral"
    if not HEX.match(str(brand["accent"])):
        raise SystemExit(f"branding.accent must be a 6-digit hex colour, got {brand['accent']!r}")
    brand["accent"] = "#" + str(brand["accent"]).lstrip("#").upper()
    for mode in ("light", "dark"):
        for token, value in brand[mode].items():
            if not HEX.match(str(value)):
                raise SystemExit(f"branding.{mode}.{token} must be a 6-digit hex colour, got {value!r}")
            brand[mode][token] = "#" + str(value).lstrip("#").upper()
    brand["logo_path"] = _logo_path(brand.get("logo", ""), config_dir)
    return brand


def _logo_path(logo: str, config_dir: str) -> str:
    if not logo:
        return ""
    if logo.startswith(("http://", "https://", "data:")):
        raise SystemExit("remote or data-URI logos are not accepted; point branding.logo at a local image file")
    for base in (config_dir, ROOT):
        candidate = logo if os.path.isabs(logo) else os.path.join(base, logo)
        if os.path.isfile(candidate):
            return os.path.abspath(candidate)
    raise SystemExit(f"branding.logo not found: {logo}")


def rgb(hex_colour: str) -> tuple[int, int, int]:
    h = hex_colour.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def luminance(hex_colour: str) -> float:
    def lin(c: int) -> float:
        c = c / 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = rgb(hex_colour)
    return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b)


def on_colour(background: str, light: str = "#FFFFFF", dark: str = "#0F172A") -> str:
    """Readable text colour on a background (WCAG-style contrast choice)."""
    lb = luminance(background)
    contrast = lambda other: (max(lb, luminance(other)) + 0.05) / (min(lb, luminance(other)) + 0.05)  # noqa: E731
    return light if contrast(light) >= contrast(dark) else dark


def mix(colour: str, other: str, amount: float) -> str:
    """Blend `colour` toward `other` by amount (0 = colour, 1 = other); returns #RRGGBB."""
    a, b = rgb(colour), rgb(other)
    return "#%02X%02X%02X" % tuple(round(a[i] + (b[i] - a[i]) * amount) for i in range(3))


def heat_scale(accent: str) -> tuple[str, str, str]:
    """Three-step brand scale for calendar heat: none, medium, high (white to accent tints)."""
    return "#FFFFFF", mix(accent, "#FFFFFF", 0.70), mix(accent, "#FFFFFF", 0.30)


def contrast(a: str, b: str) -> float:
    la, lb = luminance(a), luminance(b)
    return (max(la, lb) + 0.05) / (min(la, lb) + 0.05)


def readable_on(colour: str, background: str, ratio: float = 4.5) -> str:
    """Nudge `colour` toward white (on a dark background) or black (on a light one) until it reads on `background`."""
    toward = "#FFFFFF" if luminance(background) < 0.5 else "#000000"
    out = colour
    for step in range(0, 21):
        out = mix(colour, toward, step / 20)
        if contrast(out, background) >= ratio:
            return out
    return out


LOGO_SUFFIXES = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg"}
MAX_LOGO_BYTES = 2 * 1024 * 1024


def client_logo(org_cfg: dict | None, config_dir: str = ".") -> str:
    """Absolute path of the client's logo (config `org.logo`), or "" when none is set.

    The client's mark belongs to the client: it must be a local file the operator is allowed to use. Remote and
    data-URI logos are rejected, only common image types are accepted, and the file is capped at 2 MB.
    """
    logo = (org_cfg or {}).get("logo", "")
    path = _logo_path(logo, config_dir)
    if path:
        if os.path.splitext(path)[1].lower() not in LOGO_SUFFIXES:
            raise SystemExit(f"org.logo must be one of {sorted(LOGO_SUFFIXES)}, got {os.path.basename(path)}")
        if os.path.getsize(path) > MAX_LOGO_BYTES:
            raise SystemExit("org.logo is larger than 2 MB: use a smaller image")
    return path
