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
