# docx-renderer

Turns a deterministic `PlacementPlan` (from `@label-maker/label-layout`) into
a printable artifact. DOCX is the canonical export format for this
application, because labels are printed from Microsoft Word.

**Current support status:**

| Template    | Rendering mode | DOCX export                                                                                                                                      |
| ----------- | -------------- | ------------------------------------------------------------------------------------------------------------------------------------------------ |
| Avery 5155  | `FIXED_GRID`   | **Supported** - real, valid `.docx` generation implemented and structurally tested (see below for the caveat about the committed source fixture) |
| Avery 22802 | `FLOATING`     | **Not supported** - inspected and registered, but `renderFloatingDocx()` always throws `UnsupportedFloatingTemplateError`                        |

The optional debug-JSON renderer (`renderDebugJsonArtifact`) still exists
and can be used as a secondary, development-time output for either
template - it never claims to be a real DOCX.

## Important caveat: the committed Avery 5155 fixture uses an interleaved-spacer-column grid the validator doesn't recognize yet

`fixtures/label-templates/avery-5155/original.docx` **is** a valid Word
document (it was previously an invalid/theme-derived file; that has been
replaced). Run `pnpm exec tsx scripts/inspect-docx-template.ts
fixtures/label-templates/avery-5155/original.docx` to see the current real
report: one table, 15 rows, fixed layout, exact 950-twip row heights, real
US Letter page geometry - but its `<w:tblGrid>` declares **7** columns
(widths in twips: `[2520, 432, 2520, 432, 2520, 432, 2520]`), not the 4 this
validator currently expects. Inspecting the raw XML shows the narrow
432-twip columns (indices 1, 3, 5) carry `<w:vMerge>`, i.e. this is the
common Avery Word-template pattern of **4 real label columns interleaved
with 3 narrow, vertically-merged spacer/gutter columns** used to get the
correct horizontal pitch between label columns - 105 raw grid cells, of
which 60 are the actual writable label positions.

Practical effect: `renderFixedGridDocx()` throws
`InvalidFixedGridTemplateError` against the committed fixture today, because
`validateFixedGridTemplate()` requires every row to have exactly `columns`
(4) cells and does not yet understand interleaved spacer columns. **This is
a real, common template pattern, not a corrupt or wrong file** - handling
it correctly requires validator/renderer changes (recognizing which grid
columns are real label columns vs. merged spacers) that are intentionally
out of scope for the current milestone. This package's DOCX-generation
tests exercise the existing uniform-4-column renderer path against a small,
controlled, hand-authored fixed-grid `.docx` fixture built purely for
structural testing (never claimed to be the real Avery 5155 template, and
never derived from any product deck); `tests/docx-renderer/real-avery-5155.test.ts`
separately documents the real fixture's actual (currently-failing)
behavior, so this gap is tracked rather than silently left undiscovered.

## Rendering strategy

### Fixed-grid labels (Avery 5155)

- `renderFixedGridDocx()` starts from a **byte-for-byte JSZip copy** of the
  source template and only ever rewrites `word/document.xml`. Every other
  package part (styles, theme, fonts, settings, relationships, media) is
  preserved untouched.
- `validateFixedGridTemplate()` is a strict, throwing gate: it requires
  exactly one normal in-flow table with the expected `columns x rows =
labelsPerSheet` shape, `<w:tblLayout w:type="fixed"/>`, and an explicit
  `<w:tblGrid>` with one `<w:gridCol>` per column. `renderFixedGridDocx()`
  never runs against a template that fails this check.
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

- [ ] Extend `validateFixedGridTemplate()`/`renderFixedGridDocx()` to
      recognize interleaved vertically-merged spacer columns (see the
      caveat above) so the real, valid `fixtures/label-templates/avery-5155/original.docx`
      validates and renders - this is the actual blocker today, not an
      invalid/missing source file.
- [ ] Once that support exists, populate the seed's `geometry.label`/
      `geometry.pitch`/`geometry.safeInsets` (currently left as explicit
      TODOs - see `packages/database/prisma/seed.ts`) from the real
      per-column measurements already captured under `geometry.rawTableGrid`.
- [ ] Complete `fixtures/label-templates/avery-5155/PRINT-TEST.md`'s manual
      acceptance procedure once rendering succeeds against the real template.
- [ ] Design and implement `renderFloatingDocx()` for Avery 22802 - a
      separate geometry model (page-anchored absolute positioning) from the
      fixed-grid path, tested by printing from Word.
- [ ] Consider extending `validateFixedGridTemplate()`/inspection to sample
      more than one cell per table (currently only the first row/cell is
      sampled for font/paragraph/margin properties), if real templates turn
      out to have inconsistent per-cell formatting worth catching.
