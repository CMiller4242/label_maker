# Label template fixtures

```
label-templates/
  avery-5155/
    original.docx    Source template for the FIXED_GRID preset (see caveat below)
    PRINT-TEST.md     Manual physical-print acceptance procedure
  avery-22802/
    original.docx    Source template for the FLOATING preset (inspected, DOCX export not implemented)
```

Both `LabelTemplate` presets (`avery-5155`, `avery-22802`) point at the
`.docx` files here via `templateStorageKey`, seeded by
`packages/database/prisma/seed.ts` (which also runs
`inspectDocxTemplate()` against both at seed time and stores the result in
`configJson.sourceInspection`).

**Caveat:** `avery-5155/original.docx` is a valid Word document (an
earlier invalid/theme-derived version of this file has been replaced), but
DOCX generation still fails against it: its `<w:tblGrid>` has 7 columns (4
real label columns interleaved with 3 narrow, vertically-merged spacer
columns for horizontal pitch - a common Avery Word-template pattern), not
the uniform 4 columns per row `validateFixedGridTemplate()` currently
requires. Run `pnpm exec tsx scripts/inspect-docx-template.ts
fixtures/label-templates/avery-5155/original.docx` to confirm this
yourself. `POST /label-runs` against `avery-5155` fails with a typed
`TEMPLATE_UNAVAILABLE` error until the validator/renderer are extended to
recognize interleaved spacer columns. See
`packages/docx-renderer/src/README.md` and `avery-5155/PRINT-TEST.md` for
the full explanation and the manual print-validation procedure to run once
rendering succeeds.

`avery-22802/original.docx` is a genuine, valid Word document: 8 tables,
each using floating/anchored positioning (`<w:tblpPr>`), which is why it is
registered as `renderingMode: "FLOATING"`. DOCX export for floating
templates is intentionally not implemented yet (see
`packages/docx-renderer/src/README.md`).

Do not add real product deck PDFs (e.g. the Woodhull deck) or other
customer-supplied documents here or anywhere else in this repository - see
the root README's "Sensitive documents" section.
