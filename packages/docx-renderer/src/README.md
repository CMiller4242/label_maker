# docx-renderer

This package is the boundary between a deterministic `PlacementPlan` (from
`@label-maker/label-layout`) and a printed sheet of labels. **Today it does
not produce a real `.docx` file.** It defines the renderer interface every
future implementation must satisfy, a template loader/validator that works
against real `.docx` bytes once a verified template exists, and a working
stub renderer (`renderDebugJsonArtifact`) that writes a structured JSON
artifact describing exactly what a real renderer would need to draw.

`renderFixedGridDocx()` and `renderFloatingDocx()` both throw
`DocxNotImplementedError` right now. Do not treat their presence as a claim
that print output has been validated - it has not. Producing correct,
print-accurate `.docx` output is the next milestone, not part of this
scaffold.

## Why DOCX at all

Labels in this workflow are printed from Microsoft Word, not from a PDF
export pipeline. That makes `.docx` the canonical export format: whatever
this package eventually produces must open and print correctly in Word
without the user manually fixing table layout, fonts, or margins.

## Planned rendering strategy

### Standard (FIXED_GRID) labels - e.g. Avery 5155

- Standard label sheets are rendered by **cloning a fixed-grid Word table**
  out of a **verified source template** (a real `.docx`/`.dotx` file whose
  geometry has been measured against a physical printed sheet - see
  `fixtures/label-templates/README.md`). We do not generate table geometry
  from scratch; we clone rows/cells from a template that is already known to
  align with the physical label stock.
- Each table **row must use an exact height** (`w:trHeight` with
  `w:hRule="exact"`), not `"atLeast"`. Word's default "at least" behavior
  lets row heights grow with content, which silently shifts every label
  below it out of alignment with the physical sheet - unacceptable for a
  print-accuracy-critical workflow.
- Tables **must use fixed layout** (`<w:tblLayout w:type="fixed"/>`) and must
  **never auto-fit**. Autofit recalculates column widths from cell content,
  which is exactly what fixed-grid label printing cannot tolerate.
  `validateTemplateLayout()` will be extended to assert this before any
  render proceeds.
- Every cell must declare **explicit margins** (`w:tcMar`) and **explicit
  paragraph spacing** (`w:spacing` with `w:before`/`w:after` pinned, not
  inherited from a Normal style that could change). Implicit/inherited
  spacing is the single most common cause of "it looked right in the
  preview but printed one line off."
- The generator may insert a **page break only between complete sheet
  tables** - i.e. once every slot on a sheet has been filled (or left
  intentionally blank per `PlacementResult.unfilledSlots`), not between
  individual products or rows. This mirrors `build-placements.ts`, which
  deliberately does not force a page break per product.
- Implementation sketch (not yet built): `loadTemplateDocx()` opens the
  verified template via JSZip, `validateTemplateLayout()` confirms its
  `<w:tbl>` structure matches `LabelTemplate.rows`/`columns`, and
  `renderFixedGridDocx()` will clone the table's row/cell XML
  `labelsPerSheet` times per sheet, filling each cell's text runs from the
  corresponding `Placement`, then reassembling `word/document.xml` inside
  the zip and writing the modified package back out.

### Floating templates

- Some label stock (and some proof-sheet workflows) do not use a uniform
  grid - tags/labels are positioned freely on the page (floating text boxes
  or absolutely-positioned frames rather than table cells).
- `renderFloatingDocx()` is a **separate implementation path** from the
  fixed-grid renderer, not a generalization of it. Floating placement in
  OOXML uses drawing anchors (`<w:drawing>`/`wp:anchor`) with explicit
  x/y offsets rather than table rows, so the geometry model, validation
  rules, and template requirements are different enough that sharing code
  with the fixed-grid path would obscure both.
- Floating templates **must be authored and print-tested directly in
  Microsoft Word** before being wired up here - there is no physical-sheet
  fallback the way there is for a standard Avery grid, so an unverified
  floating template has no ground truth to check placement against.

## What exists today

- `types.ts` - the `LabelRenderer` interface (`loadTemplateDocx`,
  `validateTemplateLayout`, `renderFixedGridDocx`, `renderFloatingDocx`),
  `TemplatePackage`, and `DocxNotImplementedError`.
- `template-loader.ts` - opens a real `.docx` zip package with JSZip and
  extracts `word/document.xml`; `validateTemplateLayout()` currently only
  checks that a `<w:tbl>` element exists for `FIXED_GRID` templates. The
  deeper checks described above (exact row heights, fixed table layout,
  explicit cell margins) are TODOs.
- `debug-json-renderer.ts` - `renderDebugJsonArtifact()` is the renderer
  actually used by `apps/worker`'s `build-label-run` processor today. It
  writes every `Placement` (sheet, slot, row, column, product fields) plus
  the intended `labelTemplateId` as JSON. `DebugJsonRenderer` implements the
  full `LabelRenderer` interface, but its `renderFixedGridDocx`/
  `renderFloatingDocx` methods throw `DocxNotImplementedError` - they exist
  so calling code can already be written against the stable interface.

## TODOs for the DOCX milestone

- [ ] Acquire/measure a verified Avery 5155 physical template; replace the
      placeholder `configJson.geometry` values in the seed data.
- [ ] Implement `renderFixedGridDocx()`: clone table rows/cells from the
      verified template, fill text runs, reassemble `document.xml`.
- [ ] Extend `validateTemplateLayout()` with the fixed-layout/exact-row-
      height/explicit-margin checks described above.
- [ ] Implement a floating-template authoring workflow and
      `renderFloatingDocx()`, tested by printing from Word.
- [ ] Add print-accuracy regression tests (e.g. rendering a known plan and
      diffing the produced `document.xml` table geometry).
