# Avery 5155 - manual print acceptance test

Automated tests in `tests/docx-renderer/` and `tests/integration/` verify
that `renderFixedGridDocx()` produces a **structurally valid** `.docx`:
correct table/row/cell counts, fixed table layout, explicit row heights,
`w:cantSplit` on every row, expected SKU/description/price text, and no
stray content outside the label grid. **None of that proves the labels
print in the right physical positions on a real Avery 5155 sheet.**
Structural validity and print accuracy are different guarantees - CI can
only check the former. This document is the manual procedure for the
latter, and it has not been performed yet for this milestone.

## Prerequisite: rendering does not succeed against this template yet

`fixtures/label-templates/avery-5155/original.docx` **is** a valid Word
document (an earlier, invalid/theme-derived version of this file has been
replaced). However, `renderFixedGridDocx()` still fails against it: its
`<w:tblGrid>` has 7 columns (4 real label columns interleaved with 3
narrow, vertically-merged spacer columns for horizontal pitch - a common
Avery Word-template pattern), not the uniform 4 columns per row the
renderer currently requires. Run `pnpm exec tsx scripts/inspect-docx-template.ts
fixtures/label-templates/avery-5155/original.docx` to confirm this
yourself; see `packages/docx-renderer/src/README.md` for the full
explanation and `tests/docx-renderer/real-avery-5155.test.ts` for the
tests documenting this exact, current behavior. Before any of the steps
below are possible:

1. Extend `validateFixedGridTemplate()`/`renderFixedGridDocx()` (in
   `packages/docx-renderer/src/`) to recognize interleaved
   vertically-merged spacer columns, so it can correctly identify the 4
   real label columns among the 7 raw grid columns.
2. Re-run `pnpm exec tsx scripts/inspect-docx-template.ts` and confirm
   generation succeeds end-to-end against the real fixture.
3. Populate `packages/database/prisma/seed.ts`'s
   `avery5155ConfigJson.geometry.label`/`.pitch`/`.safeInsets` (currently
   explicit TODOs) from the real per-column measurements already captured
   under `geometry.rawTableGrid`, and finalize `provisionalLabelTextStyle`
   before trusting any print output.

## Manual acceptance procedure

Perform this **every time** the source template, `labelTextStyle` config,
or `renderFixedGridDocx()` itself changes in any way that could affect
physical positioning.

### 1. Generate a test label run

- Use `POST /label-runs` (or the equivalent worker job) with a small set of
  controlled test products - enough to fill at least one complete sheet
  (60 labels) plus a partial second sheet, e.g. 8 products x 8 copies (60
  on sheet 1, 4 on sheet 2).
- Download the generated artifact via `GET /label-runs/:id/download`.

### 2. Open in Microsoft Word

- Open the downloaded `.docx` in a real, licensed Microsoft Word
  installation - not Word Online, not LibreOffice, not Google Docs. Only
  Word's own rendering/printing pipeline is authoritative here.
- **Verify expected page count.** For an 8-products x 8-copies run: exactly
  2 pages.
- **Confirm 60 positions per complete sheet.** Page 1 should show all 60
  cells occupied edge-to-edge across 4 columns x 15 rows; page 2 should
  show only the first 4 cells (row 1) occupied, the remaining 56 blank.

### 3. Print at Actual Size

- Print Setup: **Actual Size / 100%** scaling. **Never** "Fit to Page" /
  "Shrink to Fit" / "Scale to Fit Paper" - any of these will silently
  resize every label and invalidate the whole test.
- Confirm US Letter paper size is selected (or whatever the source
  template's `<w:pgSz>` actually declares - see the inspection report).

### 4. First pass: print on plain paper

- Print the first page onto ordinary plain paper (not a real Avery sheet
  yet).
- **Overlay the plain-paper printout behind a real, unprinted Avery 5155
  sheet**, holding both up to a light source (or a well-lit window). Check
  that the printed label boundaries line up with the sheet's actual
  perforated/die-cut label boundaries.
- Inspect **five positions specifically**: top-left, top-right,
  bottom-left, bottom-right, and center. These four corners plus center are
  the positions most likely to reveal cumulative drift (a small per-column
  or per-row offset that's invisible on one label but obvious 14 rows
  down).

### 5. Confirm the blank source template aligns first

- Before attributing any misalignment to the _generated_ output, print the
  **unmodified source template** (`original.docx`, with its blank cells)
  the same way and check it against a real sheet the same way.
- If the blank template itself doesn't align, the problem is in the source
  template's geometry (or the printer/tray/scaling setup), not in
  `renderFixedGridDocx()` - fix that first, independently.

### 6. Record the result

For every test pass, record all of the following (a simple markdown table
appended below this file, or a linked doc, is fine):

| Field                                     | Value                          |
| ----------------------------------------- | ------------------------------ |
| Date                                      |                                |
| Printer model                             |                                |
| Tray used                                 |                                |
| Paper size selected                       |                                |
| Word version                              |                                |
| Scaling setting used                      | (must be "Actual Size / 100%") |
| X correction needed (if any)              |                                |
| Y correction needed (if any)              |                                |
| Blank template alignment                  | pass / fail                    |
| Generated output alignment (top-left)     | pass / fail                    |
| Generated output alignment (top-right)    | pass / fail                    |
| Generated output alignment (bottom-left)  | pass / fail                    |
| Generated output alignment (bottom-right) | pass / fail                    |
| Generated output alignment (center)       | pass / fail                    |
| Overall result                            | pass / fail                    |
| Notes                                     |                                |

## Bottom line

Rendering can be **structurally** validated in CI (and is - see
`tests/docx-renderer/` and `tests/integration/`). **Physical print
accuracy on a real Avery 5155 sheet requires this manual test**, performed
by a human with a real printer and real label stock. Do not treat a green
CI run as evidence that labels will print correctly.
