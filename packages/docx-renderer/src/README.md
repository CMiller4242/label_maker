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

## Important caveat: the committed Avery 5155 fixture is not a valid Word document

`fixtures/label-templates/avery-5155/original.docx` currently only contains
`[Content_Types].xml`, `_rels/.rels`, and `theme/theme/*` parts - **no
`word/document.xml`**. Its `_rels/.rels` points at
`theme/theme/themeManager.xml`, which is the structure of an Office _theme_
package (`.thmx`), not a Word document. Run
`pnpm exec tsx scripts/inspect-docx-template.ts` to see this for yourself;
the inspector reports `tableCount: 0`, `classification: "AMBIGUOUS"`, and an
explicit warning naming the missing part.

Practical effect: `renderFixedGridDocx()` will throw
`InvalidFixedGridTemplateError` against the committed fixture today, because
`loadTemplateDocx()`/`validateFixedGridTemplate()` correctly refuse to treat
a non-Word-document as a usable template. **The renderer code itself is
complete and works against a real, valid Avery 5155 `.docx`** containing a
normal in-flow, fixed-layout table of exactly 4 columns × 15 rows (60
cells) - this package's DOCX-generation tests exercise that exact path
against a small, controlled, hand-authored fixed-grid `.docx` fixture built
purely for structural testing (never claimed to be the real Avery 5155
template, and never derived from any product deck). Until a real Avery 5155
source template is committed to `fixtures/label-templates/avery-5155/`,
end-to-end rendering against the "real" fixture will fail loudly and
correctly, rather than silently producing garbage.

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

- [ ] Replace `fixtures/label-templates/avery-5155/original.docx` with a
      real, valid Avery 5155 Word template (see the caveat above), then
      re-run `pnpm exec tsx scripts/inspect-docx-template.ts` and update the
      seed's `geometry`/`labelTextStyle` from actual measurements.
- [ ] Complete `fixtures/label-templates/avery-5155/PRINT-TEST.md`'s manual
      acceptance procedure once a real template exists.
- [ ] Design and implement `renderFloatingDocx()` for Avery 22802 - a
      separate geometry model (page-anchored absolute positioning) from the
      fixed-grid path, tested by printing from Word.
- [ ] Consider extending `validateFixedGridTemplate()`/inspection to sample
      more than one cell per table (currently only the first row/cell is
      sampled for font/paragraph/margin properties), if real templates turn
      out to have inconsistent per-cell formatting worth catching.
