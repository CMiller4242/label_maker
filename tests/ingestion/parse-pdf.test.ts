import { describe, expect, it } from "vitest";
import { parsePdf } from "@label-maker/ingestion";

/**
 * Builds a minimal single-page PDF with one text-showing stream, using raw
 * PDF syntax (no external PDF-writing library). Good enough to exercise
 * parsePdf()'s native text extraction path without depending on a real
 * product-deck fixture (see fixtures/pdf/README.md for that TODO).
 */
function buildMinimalPdf(pageText: string): Buffer {
  const contentStream = `BT /F1 12 Tf 10 50 Td (${pageText}) Tj ET`;
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 4 0 R >> >> /MediaBox [0 0 300 150] /Contents 5 0 R >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    `<< /Length ${contentStream.length} >>\nstream\n${contentStream}\nendstream`,
  ];

  let pdf = "%PDF-1.4\n";
  const offsets: number[] = [];
  objects.forEach((body, index) => {
    offsets.push(Buffer.byteLength(pdf, "latin1"));
    pdf += `${index + 1} 0 obj\n${body}\nendobj\n`;
  });

  const xrefOffset = Buffer.byteLength(pdf, "latin1");
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets) {
    pdf += `${offset.toString().padStart(10, "0")} 00000 n \n`;
  }
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;

  return Buffer.from(pdf, "latin1");
}

describe("parsePdf", () => {
  it("extracts native text and discovers generic SKU/price candidates", async () => {
    const pdf = buildMinimalPdf("SKU-1001 AS LOW AS $9.99");
    const result = await parsePdf(pdf);

    expect(result.documentType).toBe("PDF");
    expect(result.pageCount).toBe(1);
    expect(result.pages).toHaveLength(1);
    // Native extraction is not guaranteed reliable for arbitrary product
    // decks, so we never claim finished Product records here.
    expect(result.products).toHaveLength(0);
    expect(result.needsReview).toBe(true);

    const page = result.pages[0];
    expect(page?.rawText).toContain("SKU-1001");

    const skuCandidates = page?.candidates.filter((c) => c.field === "SKU") ?? [];
    expect(skuCandidates.some((c) => c.rawValue.includes("SKU-1001"))).toBe(true);

    const priceCandidates = page?.candidates.filter((c) => c.field === "PRICE") ?? [];
    expect(priceCandidates.length).toBeGreaterThan(0);
  });

  it("flags low text density pages as needing review with OCR_REQUIRED noted", async () => {
    const pdf = buildMinimalPdf("x");
    const result = await parsePdf(pdf);

    const page = result.pages[0];
    expect(page?.needsReview).toBe(true);
    expect(page?.suggestedFutureExtractionMethod).toBe("OCR_REQUIRED");
  });
});
