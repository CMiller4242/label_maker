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

**Note:** `avery-5155/original.docx` is a valid Word document whose
`<w:tblGrid>` has 7 raw columns (4 real label columns interleaved with 3
narrow, vertically-merged spacer columns for horizontal pitch - a common
Avery Word-template pattern), not a uniform 4 columns per row. Run `pnpm
exec tsx scripts/inspect-docx-template.ts
fixtures/label-templates/avery-5155/original.docx` to confirm this
yourself - it reports classification `AVERY_5155_LIKE_FIXED_GRID` via the
`INTERLEAVED_SPACER` fixed-grid pattern. `POST /label-runs` against
`avery-5155` succeeds and produces a real `.docx`. See
`packages/docx-renderer/src/README.md` and `avery-5155/PRINT-TEST.md` for
the full explanation and the manual print-validation procedure (physical
print alignment has not yet been verified).

`avery-22802/original.docx` is a genuine, valid Word document: 8 tables,
each using floating/anchored positioning (`<w:tblpPr>`), which is why it is
registered as `renderingMode: "FLOATING"`. DOCX export for floating
templates is intentionally not implemented yet (see
`packages/docx-renderer/src/README.md`).

Do not add real product deck PDFs (e.g. the Woodhull deck) or other
customer-supplied documents here or anywhere else in this repository - see
the root README's "Sensitive documents" section.
