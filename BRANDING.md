# CompleteTech LLC Skills: branding and handoffs

[Start here](ONBOARDING.md) · [Asset rights](BRAND_ASSETS.md) · [Specialist instructions](SKILL.md)

## Resolve identity before rendering

An approved engagement brand takes precedence over a starter example. Preserve the operator's explicit settings; ask about unresolved identity instead of silently applying company branding. Output is **neutral by default**: no company name, no logo, no tagline. CompleteTech examples identify this library, not permission to represent another business.

Select an identity in `config.json`:

```json
"branding": { "preset": "completetech" }
```

or give explicit values (`name`, `eyebrow`, `tagline`, `contact`, `logo`, `accent`, `footer`, `font`, `light{}`, `dark{}`); explicit values override the preset. Presets live in `templates/branding.<preset>.json`. Use the existing local `assets/logo.png` for authorized CompleteTech output. Do not synthesize replacement marks, fetch remote artwork, redistribute fonts, or transfer private seals and signatures to public repositories. The MIT code license does not replace the asset policy. A logo must be a local file; remote and data-URI logos are rejected.

## Client logo

`org.logo` in `config.json` adds the client's own mark beside the brand mark: in the workbook banner and at the right of the report header (on a white tile, so a dark logo stays legible in dark mode). The client's mark belongs to the client. Use it only with the client's permission, supply a local file (remote and data-URI logos are rejected, files over 2 MB are refused), and keep it out of Git and out of the registry bundle: the repository's ignore files exclude `logos/`, `client-logo.*` and `*-client-logo.*`. It is separate from the brand identity, so a neutral report can still carry a client logo. SVG renders in the report but not in the workbook.

## Starter palette

These established library colors are a reference for the `completetech` preset:

| Token | Light | Dark |
|---|---|---|
| Accent | `#1E3A8A` | `#1E3A8A` |
| Background (`bg`) | `#F8FAFC` | `#0F172A` |
| Surface (`surface`) | `#FFFFFF` | `#1E293B` |
| Soft surface (`surface-2`) | `#EEF2FF` | `#273449` |
| Ink (`ink`, `ink-2`, `ink-3`) | `#0F172A`, `#1E293B`, `#64748B` | `#F1F5F9`, `#CBD5E1`, `#94A3B8` |
| Border (`line`, `line-2`) | `#E2E8F0`, `#CBD5E1` | `#334155`, `#475569` |

Theme: the HTML report (with its From / To date filter) carries both token sets and follows the reader's system (`style.theme` or `--theme` forces `light` or `dark`). Excel has no light/dark switch, so the workbook uses the light tokens: the accent colors table headers, tab colors and the calendar heat scale, and header text color is chosen for contrast. Yellow input cells stay yellow in every brand: it is a semantic convention (cells the operator may edit), not a theme color.

Keep specialist layouts: this skill's workbook and report differ from invoices, envelopes or certificates. Use the existing generator options and configuration documented in `SKILL.md` and `--help`; this document does not introduce a universal brand-import flag.

## Handoff to the next skill

Carry verified facts and their sources, artifact path and version, approved identity and local logo path, palette overrides, confidentiality, intended audience, approval owner and status, blockers, the billing parameters used, borderline cases awaiting a decision, and the next specialist (for example `agentic-invoice-skill`). Preserve these in the orchestrator's existing `project_state`; do not replace its workflow schema with an incompatible parallel state format.

Missing approval stays unknown or draft. Rendering, redaction, a clean package check or a reused brand does not authorize sending, publication, billing, signing or launch. Review the actual output for clipping, contrast, placeholders and factual accuracy before approved use.
