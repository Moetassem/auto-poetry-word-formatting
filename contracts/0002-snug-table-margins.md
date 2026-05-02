# 0002 — Snug Poetry Table Margins to the Widest Hemistich

- **Date:** 2026-05-01
- **Status:** Approved
- **Suggested branch:** `feat/snug-table-margins`

## Scope contract

```
Goal:           Manual macro that, given a poetry table at the cursor,
                shrinks the table's left/right cell padding (Word's
                "default cell margins") so the widest hemistich sits in
                its cell with a small ~0.2 cm horizontal halo. The same
                padding applies to every cell, so narrower-text cells
                gain visibly more breathing room.
Files touched:  auto-poetry-word-formatting/macros/FormatPoem.bas
                (one new Public Sub + two private helpers; existing
                routines untouched)
                auto-poetry-word-formatting/contracts/0002-snug-table-margins.md
                (this file, new)
Out of scope:   - Running this automatically on Enter — stays manual
                - Changing column widths or the 50/50 split
                - Top/bottom padding (existing 0.2 cm bottom stays)
                - Compensating for the paragraph-mark advance in
                  centered RTL (small visible asymmetry on the ruler
                  is expected and documented; left/right halo is
                  symmetric in advance-width terms)
                - Cells whose text already wraps at zero padding —
                  their wrap is taken as the baseline, not "fixed"
                - Tests / CI
Tests added:    None — VBA-only project, "How to run / test: manual
                testing" per AGENTS.md project context.
Done when:      With the cursor inside a poetry table built by
                FormatArabicPoetryOnEnter, running
                AdjustPoetryTableMargins shrinks the table's left/right
                padding so the widest hemistich sits ~0.1 cm from each
                side of its cell content area, and narrower-text cells
                in the same table show visibly more empty space around
                their text. No cell gets squished into a new line.
Rollback:       Undo (Ctrl+Z) inside Word for a single run, or
                `git revert <this-commit>` to remove the macro.
Reverse Context: None — purely additive; no existing routines, key
                bindings, state vars, or AutoOpen behavior were
                changed.
```

## Plan

1. Add `Public Sub AdjustPoetryTableMargins` to `FormatPoem.bas`. Validates the cursor is inside a table; saves the original `LeftPadding` / `RightPadding` and restores them on any error via an `On Error GoTo Cleanup` block.
2. Compute the bisection upper bound as half the narrowest cell's width (any larger padding would give that cell a non-positive content area).
3. With padding temporarily set to `0`, capture each cell's natural display-line layout via the helper `CellTextSpan` (vertical gap between the cursor at the cell text's logical start and its logical end). This is the per-cell baseline. Cells with no text are detected via `CellHasText` so the macro exits silently for empty tables.
4. Bisect the padding value within `[0, hiPts]` to a 0.01 cm tolerance (max 30 iterations). At each step, set the padding and ask Word's layout engine — via `CellTextSpan` again — whether any cell's text now spans more vertical distance than its baseline (with a 1 pt epsilon). The largest padding that still doesn't trigger a new wrap is `lo`.
5. Back off `lo` by 0.1 cm per side to leave a ~0.2 cm halo around the widest text. If `lo < 0.1 cm` (no room for the halo), drop the back-off and use `lo` as-is. Clamp to `>= 0`.
6. Apply the final value to `tbl.LeftPadding` and `tbl.RightPadding` (table-level "default cell margins" — propagates to every cell).
7. Add private helpers `TableExceedsBaseline`, `CellTextSpan`, and `CellHasText`. No new public API beyond `AdjustPoetryTableMargins`. No changes to `FormatArabicPoetryOnEnter`, `ConvertLineToPoetryTable`, `FillCell`, key bindings, document state, or `AutoExec` / `AutoOpen`.

## Source / context

- The earlier reference module (`Arabic Word Formatting - VBA script/ArabicPoetryTableOnly.bas`) included a padding/slack measurement step that was deliberately deferred when 0001 was scoped. This contract picks that thread up as a separate, manually-invoked macro.
- An earlier draft of this feature used a horizontal pixel-width measurement
  (`Range.Information(wdHorizontalPositionRelativeToPage)` on collapsed
  start/end ranges) to estimate each cell's text width, and computed the
  margin from `(cellWidth - widestText - 0.2 cm) / 2`. That approach was
  rejected after manual testing: it under-reported the visible width of
  centered RTL Arabic text, which produced margins that squished the text
  into a wrap. A second draft based on
  `Range.ComputeStatistics(wdStatisticLines)` was also rejected: on a
  per-cell `Range`, that statistic counts paragraph-delimited lines (always
  `1` for a single-paragraph hemistich) rather than display lines, so the
  bisection never observed a wrap and converged to the maximum possible
  padding. The shipped implementation uses vertical-position-based wrap
  detection, which is layout-driven and works correctly for centered RTL
  paragraphs.
