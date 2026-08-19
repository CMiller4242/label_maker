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

**Caveat:** `avery-5155/original.docx`, as currently committed, is **not a
valid Word document** - it has no `word/document.xml` part (it appears to
be an Office theme/`.thmx` package committed under the wrong name). Run
`pnpm exec tsx scripts/inspect-docx-template.ts` to confirm this yourself.
Until it is replaced with a real, valid Avery 5155 template (a normal
in-flow, fixed-layout table of exactly 4 columns x 15 rows), DOCX
generation against the `avery-5155` preset will fail with a typed
`TEMPLATE_UNAVAILABLE` error. See `packages/docx-renderer/src/README.md`
and `avery-5155/PRINT-TEST.md` for the full explanation and the manual
print-validation procedure to run once a real template is in place.

`avery-22802/original.docx` is a genuine, valid Word document: 8 tables,
each using floating/anchored positioning (`<w:tblpPr>`), which is why it is
registered as `renderingMode: "FLOATING"`. DOCX export for floating
templates is intentionally not implemented yet (see
`packages/docx-renderer/src/README.md`).

Do not add real product deck PDFs (e.g. the Woodhull deck) or other
customer-supplied documents here or anywhere else in this repository - see
the root README's "Sensitive documents" section.
