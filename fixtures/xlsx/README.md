# XLSX fixtures

These workbooks are generated, not hand-committed as binaries. Run:

```
pnpm exec tsx scripts/create-xlsx-fixtures.ts
```

This writes `products-two-sheets.xlsx` into this directory, containing:

- `Products` sheet: a valid SKU/Description/As Low As layout with
  unambiguous headers.
- `Archived` sheet: a second sheet with a similarly-shaped, viable-looking
  header row, so ingestion tests can verify that `parseWorkbook()` refuses
  to silently pick one and instead returns `needsReview: true`.

Regenerate this fixture any time `scripts/create-xlsx-fixtures.ts` changes.
