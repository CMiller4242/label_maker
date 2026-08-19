import { readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { parseCsv } from "@label-maker/ingestion";

const fixturesDir = path.join(import.meta.dirname, "..", "..", "fixtures", "csv");

function loadFixture(name: string): Buffer {
  return readFileSync(path.join(fixturesDir, name));
}

describe("parseCsv", () => {
  it("parses a standard CSV with unambiguous headers", () => {
    const result = parseCsv(loadFixture("products-standard.csv"));

    expect(result.documentType).toBe("CSV");
    expect(result.products).toHaveLength(5);
    expect(result.needsReview).toBe(false);

    const first = result.products[0];
    expect(first?.sku).toBe("SKU-1001");
    expect(first?.description).toBe("Stainless Steel Water Bottle");
    expect(first?.priceCents).toBe(1449);

    const cheapest = result.products[3];
    expect(cheapest?.sku).toBe("SKU-1004");
    expect(cheapest?.priceCents).toBe(450); // "4.5" -> $4.50

    const usdPriced = result.products[4];
    expect(usdPriced?.priceCents).toBe(1999); // "USD 19.99"
  });

  it("parses header-variant columns (SKU, Product Name, Low Price, Active)", () => {
    const result = parseCsv(loadFixture("products-header-variants.csv"));

    expect(result.products).toHaveLength(4);
    expect(result.needsReview).toBe(false);

    const first = result.products[0];
    expect(first?.sku).toBe("VAR-2001");
    expect(first?.description).toBe("Wireless Mouse");
    expect(first?.priceCents).toBe(1299);
    expect(first?.include).toBe(true);

    const thousands = result.products[2];
    expect(thousands?.priceCents).toBe(123450); // "$1,234.50"
    expect(thousands?.include).toBe(false); // Active = "No"

    const eachSuffix = result.products[3];
    expect(eachSuffix?.priceCents).toBe(2999); // "29.99 each"
  });

  it("flags malformed price and blank description as review issues, not crashes", () => {
    const result = parseCsv(loadFixture("products-invalid-price.csv"));

    expect(result.products).toHaveLength(3);
    expect(result.needsReview).toBe(true);

    const blankDescription = result.products[0];
    expect(blankDescription?.description).toBeNull();
    expect(blankDescription?.priceCents).toBe(1449);

    const malformedPrice = result.products[1];
    expect(malformedPrice?.priceCents).toBeNull();
    expect(result.warnings.some((w) => w.includes("malformed price"))).toBe(true);

    const valid = result.products[2];
    expect(valid?.sku).toBe("BAD-3003");
    expect(valid?.priceCents).toBe(825);
  });

  it("handles BOM, CRLF line endings, and quoted commas", () => {
    const csv = '﻿Item Number,Description,As Low As\r\nQ-1,"Widget, Deluxe",$5.00\r\n';
    const result = parseCsv(Buffer.from(csv, "utf-8"));

    expect(result.products).toHaveLength(1);
    expect(result.products[0]?.sku).toBe("Q-1");
    expect(result.products[0]?.description).toBe("Widget, Deluxe");
    expect(result.products[0]?.priceCents).toBe(500);
  });

  it("returns a warning and needsReview when multiple columns map to the same field", () => {
    const csv = "SKU,Item Number,Description,As Low As\nA-1,A-1-ALT,Widget,$1.00\n";
    const result = parseCsv(Buffer.from(csv, "utf-8"));

    expect(result.needsReview).toBe(true);
    expect(
      result.warnings.some((w) => w.includes("Multiple columns map to logical field SKU")),
    ).toBe(true);
  });
});
