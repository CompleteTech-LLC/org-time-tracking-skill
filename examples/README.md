# Examples

Synthetic only. The fixture organization (Northwind Trading Co.) and its people are fictional; nothing here comes from a real mailbox, chat or calendar.

- `python tests/make_fixtures.py` builds the tree, the workbooks and the reports under `tests/out/` (git-ignored) and prints `ALL OK`.
- `../templates/*.example.json` are the starting points for your own `config.json`, `meetings.json` and `tasks.json`. Copy them; do not edit them in place for a real export.
- The committed demonstration files in `../assets/examples/` were rendered from the fixture with the CompleteTech preset and are excluded from the ClawHub bundle.
