# PDF fixtures

No PDF fixtures are committed yet. `packages/ingestion`'s `parsePdf()` is
exercised today only by unit tests that construct minimal PDF byte streams
or that stub `pdfjs-dist`, since hand-authoring a realistic product-deck PDF
is out of scope for this backend scaffold.

TODO: once a real (or representative synthetic) product-deck PDF is
available, add it here and add an integration test asserting:

- native text extraction is attempted per page,
- pages with low text density are flagged `needsReview` with
  `suggestedFutureExtractionMethod: "OCR_REQUIRED"`,
- generic SKU/currency/"AS LOW AS" candidates are discovered but not paired
  into finished Product records (see the TODO in `parse-pdf.ts`).
