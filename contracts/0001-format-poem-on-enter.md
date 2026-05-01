# 0001 — Format Poem on Enter

- **Date:** 2026-04-29
- **Status:** Approved
- **Suggested branch:** `feat/format-poem-on-enter`

## Scope contract

```
Goal:           First-step macro: on Enter, convert an Arabic poetry line containing `**`
                into a borderless 2-column RTL table (right cell = before `**`, left cell = after).
Files touched:  auto-poetry-word-formatting/macros/FormatPoem.bas        (new file)
                auto-poetry-word-formatting/contracts/0001-format-poem-on-enter.md  (this file)
Out of scope:   Custom font/style, padding/slack tuning, merging consecutive verses into
                one table, bulk-format-all, status bar, gridline hiding, persist helper,
                tests/CI.
Tests added:    None — VBA-only project, "How to run / test: manual testing" per AGENTS.md
                project context. Manual-test recipe documented under "Done when" below.
Done when:      In Word with the macro loaded and toggle ON, typing `صدر ** عجز` then
                pressing Enter produces a 1-row, 2-column, borderless RTL table; toggling
                OFF restores normal Enter behaviour.
Rollback:       Delete auto-poetry-word-formatting/macros/FormatPoem.bas; in Word run
                ToggleArabicPoetryTable to unbind Enter, then remove the module from
                Normal.dotm in the VBE.
Reverse Context: None — both files are new; no code became redundant or was deleted.
```

> Note: this contract was originally agreed under the earlier `AGENTS.md` (171-line
> version) which did not require `Tests added:` or `Reverse Context:` lines. The
> two fields above were backfilled to match the current contract template before
> opening the PR, with no change to the originally-agreed scope.

## Plan

1. Create `auto-poetry-word-formatting/macros/FormatPoem.bas` (module name `FormatPoem`), with a short header comment explaining what it does and how to install it.
2. Add `ToggleArabicPoetryTable` — flips a document variable and binds/unbinds Enter on `NormalTemplate`.
3. Add `AutoExec` / `AutoOpen` — re-bind Enter on document open if mode is ON, so the feature survives Word restarts.
4. Add `FormatArabicPoetryOnEnter` — let Enter happen normally, then inspect the just-finished paragraph for `**`; if found, call the converter; if inside a table, do nothing.
5. Add `ConvertLineToPoetryTable` — split on the first `**`, replace the paragraph's range with a 2×1 RTL table, disable all borders, set 50/50 column widths, drop each hemistich (trimmed) into its cell. No font, no padding tuning.
6. Add helpers: `BindEnter`, `UnbindEnter`, `GetMode`/`SetMode`, `StripTrailingCR`. No status bar, no merging, no measurement.

## Source / context

- Drawn from `Arabic Word Formatting - VBA script/ArabicPoetryTableOnly.bas`, with the
  font/style, padding/slack measurement, multi-verse merging, bulk formatter, status bar,
  gridline hiding, and persistence helpers intentionally removed.
