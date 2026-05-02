# 0003 — Merge Subsequent `**` Lines Into the Previous Poetry Table

- **Date:** 2026-05-01 (revised 2026-05-02)
- **Status:** Approved
- **Suggested branch:** `feat/merge-into-previous-poetry-table`

## Scope contract

```
Goal:           When Enter fires on a `**` line that sits below an existing
                2-column poetry table separated by AT MOST 4 empty paragraphs
                (0 to 4 blank lines), append the split hemistichs as a new row
                in that table instead of creating a separate one. The blank
                paragraphs between the table and the `**` line are consumed
                by the merge (they were spacing between two would-be tables;
                the merge collapses them out).
Files touched:  auto-poetry-word-formatting/macros/FormatPoem.bas
                  (modify FormatArabicPoetryOnEnter; add two private helpers:
                   AdjacentPoetryTable, AppendLineToPoetryTable. Existing
                   ConvertLineToPoetryTable / FillCell /
                   AdjustPoetryTableMargins stay untouched.)
                auto-poetry-word-formatting/contracts/0003-merge-into-previous-poetry-table.md
                  (this file)
Out of scope:   - Merging across MORE than 4 blank paragraphs (5+ blank lines
                  is treated as the user's "new section" signal — fresh
                  table).
                - Merging across a non-empty, non-table paragraph between the
                  table and the `**` line (e.g. a heading) — fresh table.
                - Merging into a table that isn't 2-column.
                - Re-running AdjustPoetryTableMargins automatically after the
                  append (still manual).
                - Any change to ConvertLineToPoetryTable, FillCell,
                  AdjustPoetryTableMargins, key bindings, AutoExec/AutoOpen,
                  state vars, or the toggle UI.
                - Tests / CI.
Tests added:    None — VBA-only project, "How to run / test: manual testing"
                per AGENTS.md project context. Manual-test recipe under
                "Done when".
Done when:      With the toggle ON:
                  (a) Zero-gap case: typing `صدر1 ** عجز1` + Enter then
                      `صدر2 ** عجز2` + Enter immediately below the resulting
                      table appends a SECOND row; no stray blank paragraph
                      between the table and the cursor; the new row is
                      visible in the table after the macro finishes (no
                      Ctrl+Z required to "reveal" it).
                  (b) Gap case: leaving 1 to 4 blank Enter-paragraphs between
                      the existing table and the new `**` line still appends
                      a new row, AND the blank paragraphs are removed.
                  (c) Past-limit case: leaving 5+ blank paragraphs creates a
                      fresh standalone table (existing 0002 behavior).
                  (d) Non-empty interruption: a non-empty, non-table paragraph
                      between the table and the `**` line creates a fresh
                      standalone table.
Rollback:       `git revert <this-commit>` removes the merge branch and
                restores 0002 behavior (every `**` line gets its own table).
Reverse Context: FormatArabicPoetryOnEnter's tail (the single
                ConvertLineToPoetryTable call) becomes a 2-branch dispatch
                (append vs. convert). No code is deleted; one call site is
                replaced with an If/Else around the same call plus a new
                AppendLineToPoetryTable call.
```

## Plan

1. Add `Private Function AdjacentPoetryTable(ByVal para As Paragraph) As Table` to `FormatPoem.bas`. Walk back from `para.Previous`, skipping at most 4 empty paragraphs (text length 0 after `StripTrailingCR`). On the first paragraph that reports `wdWithInTable`, return its `Tables(1)` if `Columns.Count = 2`. On the first non-empty non-table paragraph, or after exceeding the 4-blank cap, return `Nothing`. The cap is a single named constant (`Const MAX_GAP As Long = 4`) so it's easy to retune later.
2. Add `Private Sub AppendLineToPoetryTable(ByVal tbl As Table, ByVal para As Paragraph, ByVal lineText As String, ByVal sepPos As Long)`. Splits `lineText` on the first `**` (same `Trim$` / `Left$` / `Mid$` logic as `ConvertLineToPoetryTable`); captures `tbl.Rows(1).Range.ParagraphFormat.Alignment` so a previously snugged table keeps its `wdAlignParagraphDistribute` on the new row; calls `tbl.Rows.Add` to append a row; fills its two cells via the existing `FillCell` helper; reapplies the captured alignment to the new row's range; finally deletes the slab `Document.Range(Start:=tbl.Range.End, End:=para.Range.End)` to remove any 1-4 in-between blank paragraphs and the user's `**` paragraph in one Range.
3. **Bugfix.** The first cut captured `previousPara.Range` BEFORE `Rows.Add` and called `.Delete` on that captured Range. `Rows.Add` inserts content at exactly that Range's start, and Word auto-expands a Range when content is inserted at its start, so the captured Range silently grew to cover the new row + the original `**` paragraph; `.Delete` then removed both, and Ctrl+Z resurrected both as a single combined undo step. Computing the deletion slab AFTER `Rows.Add` from the LIVE `tbl.Range.End` and `para.Range.End` sidesteps the issue (and naturally doubles as the gap-eating mechanism for step 1).
4. Modify `FormatArabicPoetryOnEnter`'s tail. After computing `prevText` / `sepPos` (unchanged), call `AdjacentPoetryTable(previousPara)`; if it returns a table, dispatch to `AppendLineToPoetryTable adjacentTbl, previousPara, prevText, sepPos`; otherwise call the existing `ConvertLineToPoetryTable previousPara.Range, prevText, sepPos`. Keep the surrounding `Application.ScreenUpdating` False/True bracket and the early `wasInTable` exit unchanged.
5. Verify by re-reading the modified file (and `git diff`) that nothing in `ConvertLineToPoetryTable`, `FillCell`, `AdjustPoetryTableMargins`, `TableExceedsBaseline`, `CellTextSpan`, `CellHasText`, `BindEnter` / `UnbindEnter`, `AutoExec` / `AutoOpen`, `GetMode` / `SetMode`, `StripTrailingCR`, or the module header was touched.

## Source / context

- 0001 created the per-line table; 0002 added optional snug-margins. Neither addressed multi-verse merging — 0001 explicitly listed "merging consecutive verses into one table" as out of scope. This contract picks that thread up.
- The 4-blank cap matches the user's stated upper bound from manual testing: big enough to absorb typing-comfort spacing, small enough that an obvious section break (5+ blanks, or a heading) still produces a fresh table.
- The gap paragraphs are *eaten* rather than preserved because if the user wanted them as a semantic separator they'd expect a fresh table; the fact that the merge fires means the blanks were transient typing space. Preserving them would also leave dangling empty paragraphs *below* the now-bigger table, which is visually worse than the no-gap case.
- The append path reuses `FillCell` rather than duplicating its RTL / center / vertical-align logic, so a future change to cell formatting only has to be made in one place.
- Capturing the existing first-row paragraph alignment before adding the new row preserves a user's prior `AdjustPoetryTableMargins` distribute-justification across appended rows; without this, `FillCell`'s `wdAlignParagraphCenter` would silently downgrade the new row.
