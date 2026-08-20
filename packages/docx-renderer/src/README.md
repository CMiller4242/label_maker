# docx-renderer

Turns a deterministic `PlacementPlan` (from `@label-maker/label-layout`) into
a printable artifact. DOCX is the canonical export format for this
application, because labels are printed from Microsoft Word.

**Current support status:**

| Template    | Rendering mode | DOCX export                                                                                                                                      |
| ----------- | -------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| Avery 5155  | `FIXED_GRID`   | **Supported** - real, valid `.docx` generation implemented and structurally tested, including the committed source fixture's interleaved-spacer-column grid (see below) |
| Avery 22802 | `FLOATING`     | **Not supported** - inspected and registered, but `renderFloatingDocx()` always throws `UnsupportedFloatingTemplateError`                        |

The optional debug-JSON renderer (`renderDebugJsonArtifact`) still exists
and can be used as a secondary, development-time output for either
template - it never claims to be a real DOCX.

## The committed Avery 5155 fixture: an interleaved-spacer-column grid

`fixtures/label-templates/avery-5155/original.docx` is a valid Word
document. Run `pnpm exec tsx scripts/inspect-docx-template.ts
fixtures/label-templates/avery-5155/original.docx` to see the current real
report: one table, 15 rows, fixed layout, exact 950-twip row heights, real
US Letter page geometry - and its `<w:tblGrid>` declares **7** raw columns
(widths in twips: `[2520, 432, 2520, 432, 2520, 432, 2520]`). Inspecting the
raw XML shows the narrow 432-twip columns (indices 1, 3, 5) carry
`<w:vMerge>` down the whole table: this is the common Avery Word-template
pattern of **4 real label columns interleaved with 3 narrow,
vertically-merged spacer/gutter columns** used to get the correct
horizontal pitch between label columns - 105 raw grid cells, of which 60
(raw indices 0, 2, 4, 6 in every row) are the actual writable label
positions.

`inspectDocxTemplate()`/`validateFixedGridTemplate()` recognize this shape
as the `INTERLEAVED_SPACER` fixed-grid pattern (see `FixedGridPattern` in
`types.ts`): the inspector reports `logicalLabelColumnIndexes: [0, 2, 4,
6]`, `spacerColumnIndexes: [1, 3, 5]`, and a `writableCellMap` giving every
logical label slot's exact physical `(row, cell)` location.
`renderFixedGridDocx()` uses that map as its only source of truth for where
to write: it fills only the 4 writable raw columns and never touches the 3
spacer columns (their XML, including `<w:vMerge>`, table grid widths, row
heights, and section settings, is cloned byte-for-byte from the source).
The package also still supports a simpler `SIMPLE` pattern (every raw grid
column is writable, no vertical merging) for the synthetic test fixture
used in `generation.test.ts`.

## Rendering strategy

### Fixed-grid labels (Avery 5155)

- `renderFixedGridDocx()` starts from a **byte-for-byte JSZip copy** of the
  source template and only ever rewrites `word/document.xml`. Every other
  package part (styles, theme, fonts, settings, relationships, media) is
  preserved untouched. The rewritten `word/document.xml` zip entry reuses
  the source template's own entry timestamp (via `date`) and disables
  JSZip's automatic parent-folder synthesis (`createFolders: false`), so
  two renders of the exact same template + placement plan produce
  byte-identical output - not a different `word/` directory entry stamped
  with whatever wall-clock time the render happened to run at. See
  `tests/docx-renderer/real-avery-5155.test.ts`'s reproducibility
  regression test.
- `buildXml()` strips any XML declaration node from the tree before
  serializing, then prepends exactly one canonical `<?xml ...?>` itself.
  Without this, rebuilding a tree parsed from a source document that
  already had a declaration (every real template does) produced **two**
  declarations in `word/document.xml` - a structurally valid zip that
  re-parses fine with the same lenient parser used to build it (so
  `unzip -t` and a naive re-parse-and-compare test both missed it), but
  invalid XML that Microsoft Word's strict parser rejects outright
  ("Word experienced an error trying to open the file"). This did not
  conflict with the reproducibility fix above - both are satisfied
  simultaneously; there was no tradeoff to make. `renderFixedGridDocx()`
  now also calls `assertGeneratedDocxPackageIsValid()` on the actual
  output bytes before ever returning them - it re-opens the generated
  buffer and checks every package part matches the source template's part
  set and that every `.xml`/`.rels` part is well-formed per
  fast-xml-parser's `XMLValidator` (a strict check, distinct from the
  lenient parser used elsewhere) - so this class of defect can never ship
  silently again. See `tests/docx-renderer/ooxml.test.ts` (unit-level,
  the exact root cause) and the well-formedness assertions added to
  `real-avery-5155.test.ts`/`generation.test.ts`/
  `tests/scripts/generate-sample-avery-5155.test.ts` (integration-level,
  every render path).
- `validateFixedGridTemplate()` is a strict, throwing gate: it requires
  exactly one normal in-flow, `<w:tblLayout w:type="fixed"/>` table whose
  recognized `FixedGridPattern` (`SIMPLE` or `INTERLEAVED_SPACER` - see
  `inspect-template.ts`'s `detectFixedGridPattern()`) resolves to the
  expected `columns x rows = labelsPerSheet` logical shape, with a unique,
  complete `writableCellMap` and no ambiguity between writable and
  spacer/gutter raw columns. `renderFixedGridDocx()` never runs against a
  template that fails this check.
- The single validated table is **cloned once per sheet** (via
  `structuredClone` on the parsed XML node tree - see `ooxml.ts`). Every
  cell without a placement for that sheet+slot is left exactly as cloned
  from the source (blank), per the placement plan's `unfilledSlots`.
- Every row's `<w:trPr>` is rewritten to guarantee `<w:cantSplit/>` is
  present, and to force `w:hRule="exact"` on `<w:trHeight>` **when a height
  is known from the source** - preserving the source's physical height
  number, never fabricating one. If a row has no height information at
  all, the renderer leaves it absent (see "no hardcoded physical
  measurements" below) rather than guessing.
- Exactly one explicit Word page break (`<w:br w:type="page"/>` inside its
  own paragraph) is inserted **between** completed sheet tables - never
  after the last one, and never once per product. This matches
  `@label-maker/label-layout`'s placement engine, which also never forces a
  break per product.
- The document's original `<w:sectPr>` (page size/margins/section
  properties) is preserved verbatim as the final body element - physical
  page geometry always comes from the source template, never from
  hardcoded values in this package.

### Cell content

Each filled cell gets exactly 3 paragraphs, replacing whatever was in the
cell before (no leftover blank/default paragraph is left before the label
content):

1. SKU
2. Description
3. `As Low As: $X.XX` (formatted from integer cents only - see
   `formatCentsAsDollars()` - never floating point)

Every stylistic property - font family, per-line font size, boldness,
color, horizontal/vertical alignment, line spacing, paragraph
spacing-before/after (zero by default), and optional safe insets
(cell-margin override) - comes from `LabelTemplate.configJson.labelTextStyle`
(validated by `labelTextStyleConfigSchema` in `types.ts`), never from magic
numbers scattered in renderer code. The initial values seeded for Avery
5155 are explicitly marked provisional (see the seed script and
`PRINT-TEST.md`) - they have not been confirmed against a physical
printout.

### Centering/geometry correctness

A real visual acceptance test in Microsoft Word on Windows found generated
Avery 5155 labels were not reliably centered: text slightly right-shifted,
and columns 2-4 wrapping severely on populated sheets. Structural
comparison against `fixtures/label-templates/avery-5155/original.docx`
(`unzip -t`/`-l`, `xmllint`, and direct `w:tblGrid`/`w:tcW`/`w:tblW`
inspection - never guesswork) found the source template's own `<w:tblPr>`
declares `<w:tblW w:w="0" w:type="auto"/>` - i.e. "auto" width - despite
also declaring `<w:tblLayout w:type="fixed"/>`. This combination is a
well-documented Word compatibility hazard: `tblLayout="fixed"` alone does
not reliably force Word to honor per-column widths when the table's
overall preferred width is left "auto"; Word can still recompute effective
column widths from available page width rather than the declared grid,
particularly once real content actually fills cells the template's blank
preview state never exercised. (The source's own table is also ~139 twips
wider than the page's writable area after margins - real hardware-pitch
geometry that this renderer must never shrink, but consistent with a
slight rightward overflow.)

Fixes in `renderFixedGridDocx()` (`fixed-grid-renderer.ts`), applied to
the source table once before cloning so every sheet inherits them:

- `normalizeTableWidth()` sets an explicit `<w:tblW w:type="dxa">` equal to
  the exact sum of the table's own already-declared `w:tblGrid` column
  widths - deriving an internally-consistent value from geometry the
  source already declares, never inventing a new physical dimension or
  resizing any column/row.
- Every generated paragraph now carries an explicit
  `<w:ind w:left="0" w:right="0" w:firstLine="0" w:hanging="0"/>` -
  defense in depth so no inherited paragraph indentation (from
  `styles.xml`'s `Normal` style, or a future template) can silently shift
  content off-center. Combined with the existing `<w:jc w:val="center"/>`
  and `<w:vAlign w:val="center"/>` (unchanged), centering now depends only
  on explicit, tested settings.
- A new `setTcPrChildInOrder()` helper inserts/replaces `w:tcPr` children
  (`w:vAlign`, `w:tcBorders`) at their CT_TcPrBase schema-correct position
  instead of blindly appending, since Word's strict parser can silently
  ignore/misplace an out-of-order element.

`tests/docx-renderer/real-avery-5155.test.ts` now directly compares the
generated output's `w:tblGrid`/`w:tcW`/`w:tblW`/`w:tblLayout`/paragraph
`w:jc`/cell `w:vAlign` against the **source fixture's own inspected
values** (via `inspectDocxTemplate()`), not hardcoded literals - a real
structural regression test, not a re-parse-and-hope check.

**Review-only cell outlines**: `renderFixedGridDocx()` accepts an optional
4th `options.reviewOutlines` parameter. When `true`, every cell (writable
*and* spacer/gutter, filled or blank) gets a thin, uniform `w:tcBorders`
so the real physical cell rectangles are visible on screen - useful for
visually confirming centering. **Never enable this for print**: unlike
Word's built-in non-printing "Table Gridlines" view (View > Gridlines,
already on by default for any borderless cell - true for every *writable*
cell here normally), a `w:tcBorders` border is real ink Word will print.
`pnpm docx:sample-5155` produces a dedicated
`storage/artifacts/sample-avery-5155-review-grid.docx` with this enabled;
the two standard sample artifacts are never affected. Note the source
template's own 3 spacer/gutter columns already carry their own
`w:tcBorders` (cloned untouched, as with all spacer-column XML) - `
reviewOutlines` only adds borders to the 60 writable cells that have none
in the source, so the *review* artifact shows all 105 cells outlined.

Neither this renderer nor its tests can confirm physical print alignment
on real label stock - only a human, a real printer, and the manual
plain-paper overlay procedure in `PRINT-TEST.md` can. What changed here is
the DOCX's own internal geometry declarations (self-consistent `tblW`,
zeroed paragraph indentation); it removes a known Word rendering hazard,
but final on-paper alignment still requires that manual test.

### No hardcoded physical measurements

This package never invents page/label/margin dimensions. Where the source
template has explicit values (row heights, column widths, cell margins,
page size), they are preserved. Where they're absent or the source is
unusable (as with the current Avery 5155 fixture), the renderer either
leaves that specific property unset (for non-critical row-height overrides)
or refuses to render at all (via `validateFixedGridTemplate()`), rather
than substituting an invented number that might silently conflict with the
real physical sheet.

### Floating/tag templates (Avery 22802)

- Avery 22802's real, valid source `.docx` was inspected: 8 separate
  tables, each carrying `<w:tblpPr>` (floating/anchored positioning), not
  one unified grid. This is registered as `renderingMode: "FLOATING"`.
- `renderFloatingDocx()` **always throws `UnsupportedFloatingTemplateError`**
  (or `DocxNotImplementedError` if called against a non-FLOATING template
  by mistake). It never produces partial or malformed output. Floating
  layout uses a fundamentally different geometry model (absolute
  page-anchored positions rather than table rows), so implementing it is
  future work, not a variant of the fixed-grid path.

## What exists in this package

- `ooxml.ts` - label-agnostic WordprocessingML/OOXML plumbing: JSZip
  package loading (`loadTemplateDocx`), XML parsing/serialization via
  `fast-xml-parser`'s order-preserving mode, node-tree accessor/builder
  helpers, and unit conversions (twips/points/half-points).
- `inspect-template.ts` - `inspectDocxTemplate()` (and the reusable
  `inspectLoadedTemplate()`): produces a full structural report - page
  geometry, every table's layout/grid/row/cell/font/paragraph properties,
  writable cell counts, and a best-effort classification
  (`AVERY_5155_LIKE_FIXED_GRID` / `AVERY_22802_LIKE_FLOATING` /
  `AMBIGUOUS`) with explicit warnings for missing/ambiguous metadata.
- `template-validation.ts` - `validateFixedGridTemplate()` (strict,
  throwing) and `validateTemplateLayout()` (non-throwing adapter matching
  the `LabelRenderer` interface).
- `fixed-grid-renderer.ts` - `renderFixedGridDocx()` (real DOCX
  generation) and `renderFloatingDocx()` (always throws).
- `debug-json-renderer.ts` - the pre-existing debug JSON artifact
  generator (`renderDebugJsonArtifact`) and a `DebugJsonRenderer` class
  implementing the `LabelRenderer` interface for the debug-only path.
- `types.ts` - all shared types/errors/zod schemas:
  `DocxInspectionReport` and its sub-types, `LabelRenderer`, `RenderResult`,
  `labelTextStyleConfigSchema`, and the typed errors
  (`InvalidFixedGridTemplateError`, `UnsupportedFloatingTemplateError`,
  `MissingLabelTextStyleConfigError`, `DocxNotImplementedError`).

## TODOs

- [ ] Populate the seed's `geometry.label`/`geometry.pitch`/
      `geometry.safeInsets` (currently left as explicit provisional zeros -
      see `packages/database/prisma/seed.ts`'s `logicalGrid`/`rawTableGrid`)
      with real physical measurements once manual Word/printer calibration
      (`PRINT-TEST.md`) confirms them - these are print-alignment/typography
      values, not structural ones, so they intentionally aren't inferred
      from raw OOXML geometry alone.
- [ ] Complete `fixtures/label-templates/avery-5155/PRINT-TEST.md`'s manual
      acceptance procedure by physically printing a generated sheet.
- [ ] Design and implement `renderFloatingDocx()` for Avery 22802 - a
      separate geometry model (page-anchored absolute positioning) from the
      fixed-grid path, tested by printing from Word.
- [ ] Consider extending `validateFixedGridTemplate()`/inspection to sample
      more than one cell per table (currently only the first row/cell is
      sampled for font/paragraph/margin properties), if real templates turn
      out to have inconsistent per-cell formatting worth catching.
