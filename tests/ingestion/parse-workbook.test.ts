import { readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";
import { parseWorkbook } from "@label-maker/ingestion";

function buildWorkbookBuffer(sheets: Record<string, (string | number)[][]>): Buffer {
  const workbook = XLSX.utils.book_new();
  for (const [name, rows] of Object.entries(sheets)) {
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), name);
  }
  return XLSX.write(workbook, { type: "buffer", bookType: "xlsx" }) as Buffer;
}

describe("parseWorkbook", () => {
  it("parses the single viable sheet automatically when only one sheet has unambiguous headers", () => {
    const buffer = buildWorkbookBuffer({
      Products: [
        ["Item Number", "Description", "As Low As"],
        ["XLS-1", "Widget A", "$10.00"],
        ["XLS-2", "Widget B", "$20.50"],
      ],
    });

    const result = parseWorkbook(buffer);

    expect(result.needsReview).toBe(false);
    expect(result.sheetNames).toEqual(["Products"]);
    expect(result.products).toHaveLength(2);
    expect(result.products[0]?.sku).toBe("XLS-1");
    expect(result.products[0]?.priceCents).toBe(1000);
    expect(result.products[1]?.priceCents).toBe(2050);
  });

  it("returns needsReview and no products when multiple sheets are viable", () => {
    const buffer = buildWorkbookBuffer({
      Products: [
        ["Item Number", "Description", "As Low As"],
        ["A-1", "Widget A", "$10.00"],
      ],
      Archived: [
        ["Item Number", "Description", "As Low As"],
        ["B-1", "Old Widget", "$5.00"],
      ],
    });

    const result = parseWorkbook(buffer);

    expect(result.needsReview).toBe(true);
    expect(result.products).toHaveLength(0);
    expect(result.sheetPreviews).toHaveLength(2);
    expect(result.sheetPreviews.every((s) => s.isViable)).toBe(true);
    expect(result.warnings.some((w) => w.includes("Multiple viable sheets"))).toBe(true);
  });

  it("parses a sheet using an explicit SpreadsheetMapping even when ambiguous", () => {
    const buffer = buildWorkbookBuffer({
      Products: [
        ["Item Number", "Description", "As Low As"],
        ["A-1", "Widget A", "$10.00"],
      ],
      Archived: [
        ["Item Number", "Description", "As Low As"],
        ["B-1", "Old Widget", "$5.00"],
      ],
    });

    const result = parseWorkbook(buffer, {
      sheetName: "Archived",
      headerRowNumber: 1,
      skuColumn: "Item Number",
      descriptionColumn: "Description",
      priceColumn: "As Low As",
    });

    expect(result.needsReview).toBe(false);
    expect(result.products).toHaveLength(1);
    expect(result.products[0]?.sku).toBe("B-1");
  });

  it("preserves prices as formatted strings, avoiding float rounding drift", () => {
    const buffer = buildWorkbookBuffer({
      Products: [
        ["Item Number", "Description", "As Low As"],
        ["A-1", "Widget A", 14.49],
      ],
    });

    const result = parseWorkbook(buffer);
    expect(result.products[0]?.priceCents).toBe(1449);
  });

  it("loads the committed two-sheet fixture and requires review", () => {
    const fixturePath = path.join(
      import.meta.dirname,
      "..",
      "..",
      "fixtures",
      "xlsx",
      "products-two-sheets.xlsx",
    );
    const buffer = readFileSync(fixturePath);
    const result = parseWorkbook(buffer);

    expect(result.needsReview).toBe(true);
    expect(result.sheetNames).toEqual(["Products", "Archived"]);
  });
});
