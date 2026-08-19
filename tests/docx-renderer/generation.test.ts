import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { buildPlacements } from "@label-maker/label-layout";
import type { LabelTemplate } from "@label-maker/shared";
import {
  InvalidFixedGridTemplateError,
  UnsupportedFloatingTemplateError,
  renderFixedGridDocx,
  renderFloatingDocx,
} from "@label-maker/docx-renderer";
import { buildSyntheticFixedGridDocx } from "./helpers/synthetic-fixed-grid-docx.js";

const GEOMETRY = { id: "test-fixed-grid", columns: 4, rows: 15, labelsPerSheet: 60 };

function testTemplate(overrides: Partial<LabelTemplate> = {}): LabelTemplate {
  return {
    id: "test-fixed-grid",
    displayName: "Test Fixed Grid",
    renderingMode: "FIXED_GRID",
    columns: 4,
    rows: 15,
    labelsPerSheet: 60,
    templateStorageKey: "test-fixed-grid/original.docx",
    templateVersion: "0.0.0-test",
    sourceTemplateSha256: null,
    configJson: {
      labelTextStyle: {
        fontFamily: "Calibri",
        skuFontSizeHalfPoints: 18,
        descriptionFontSizeHalfPoints: 16,
        priceFontSizeHalfPoints: 18,
        bold: false,
        colorHex: "000000",
        horizontalAlignment: "center",
        verticalAlignment: "center",
        lineSpacing: { lineRule: "auto", line: 240 },
        paragraphSpacingBeforeTwips: 0,
        paragraphSpacingAfterTwips: 0,
      },
    },
    isPreset: false,
    createdAt: new Date(),
    updatedAt: new Date(),
    ...overrides,
  };
}

function product(id: string, sku: string, description: string, priceCents: number) {
  return { id, sku, description, priceCents, include: true };
}

async function renderAndReopen(
  templateBuffer: Buffer,
  template: LabelTemplate,
  plan: Parameters<typeof renderFixedGridDocx>[2],
) {
  const result = await renderFixedGridDocx(templateBuffer, template, plan);
  const zip = await JSZip.loadAsync(result.buffer);
  const documentXml = await zip.file("word/document.xml")?.async("text");
  if (!documentXml) throw new Error("word/document.xml missing from generated docx");
  return { result, zip, documentXml };
}

describe("renderFixedGridDocx", () => {
  it("1 product x 8 copies: 1 sheet, slots 0-7 filled, 52 empty", async () => {
    const templateBuffer = await buildSyntheticFixedGridDocx({ columns: 4, rows: 15 });
    const template = testTemplate();
    const plan = buildPlacements([product("p1", "SKU-1", "Widget One", 999)], GEOMETRY, 8);

    const { documentXml, zip } = await renderAndReopen(templateBuffer, template, plan);

    // Required package entries preserved.
    expect(zip.file("[Content_Types].xml")).toBeTruthy();
    expect(zip.file("_rels/.rels")).toBeTruthy();
    expect(zip.file("word/document.xml")).toBeTruthy();

    // Well-formed XML.
    const parser = new XMLParser({ ignoreAttributes: false });
    expect(() => parser.parse(documentXml)).not.toThrow();

    // Exactly 1 sheet-grid table, no page break (only 1 sheet).
    expect((documentXml.match(/<w:tbl>/g) ?? []).length).toBe(1);
    expect(documentXml).not.toContain('w:type="page"');

    // Fixed layout, every row has cantSplit.
    expect(documentXml).toContain('<w:tblLayout w:type="fixed"');
    const rowCount = (documentXml.match(/<w:tr>/g) ?? []).length;
    const cantSplitCount = (documentXml.match(/<w:cantSplit\/?>/g) ?? []).length;
    expect(rowCount).toBe(15);
    expect(cantSplitCount).toBe(15);

    // Content present.
    expect(documentXml).toContain("SKU-1");
    expect(documentXml).toContain("Widget One");
    expect(documentXml).toContain("As Low As: $9.99");

    // No unexpected paragraphs: 8 filled cells x 3 label-line paragraphs
    // (24) + 52 untouched blank cells x 1 empty <w:p/> each (52) = 76, and
    // nothing else in the document body outside the table/sectPr.
    expect(plan.totalPlacements).toBe(8);
    expect(
      plan.placements.every((p) => p.sheetNumber === 1 && p.slotIndex >= 0 && p.slotIndex <= 7),
    ).toBe(true);
    const paragraphOpenCount = (documentXml.match(/<w:p(?:\/>|>)/g) ?? []).length;
    expect(paragraphOpenCount).toBe(8 * 3 + 52 * 1);
  });

  it("8 products x 8 copies: sheet 1 has 60 filled slots, sheet 2 has 4, no break after the first product", async () => {
    const templateBuffer = await buildSyntheticFixedGridDocx({ columns: 4, rows: 15 });
    const template = testTemplate();
    const products = Array.from({ length: 8 }, (_, i) =>
      product(`p${i}`, `SKU-${i}`, `Widget ${i}`, 100 * (i + 1)),
    );
    const plan = buildPlacements(products, GEOMETRY, 8);

    expect(plan.totalSheets).toBe(2);
    const sheet1 = plan.placements.filter((p) => p.sheetNumber === 1);
    const sheet2 = plan.placements.filter((p) => p.sheetNumber === 2);
    expect(sheet1).toHaveLength(60);
    expect(sheet2).toHaveLength(4);

    const { documentXml } = await renderAndReopen(templateBuffer, template, plan);

    // Exactly 2 sheet-grid tables.
    expect((documentXml.match(/<w:tbl>/g) ?? []).length).toBe(2);

    // Exactly 1 page break, and it appears BETWEEN the two tables (not
    // after the first product's 8 slots, and not after the last table).
    const pageBreakMatches = documentXml.match(/w:type="page"/g) ?? [];
    expect(pageBreakMatches).toHaveLength(1);

    const firstTableEnd = documentXml.indexOf("</w:tbl>");
    const secondTableStart = documentXml.indexOf("<w:tbl>", firstTableEnd);
    const pageBreakIndex = documentXml.indexOf('w:type="page"');
    expect(pageBreakIndex).toBeGreaterThan(firstTableEnd);
    expect(pageBreakIndex).toBeLessThan(secondTableStart);

    // No page break appears inside the first table's content (i.e. not
    // triggered per-product / per-8-slots).
    const firstTableXml = documentXml.slice(0, firstTableEnd);
    expect(firstTableXml).not.toContain('w:type="page"');

    // Every row across both tables (30 total) still has cantSplit and fixed layout.
    expect((documentXml.match(/<w:cantSplit\/?>/g) ?? []).length).toBe(30);
    expect((documentXml.match(/<w:tblLayout w:type="fixed"/g) ?? []).length).toBe(2);

    // Document ends with sectPr as the final body element (no trailing
    // empty paragraph appended after the last table).
    const bodyEnd = documentXml.indexOf("</w:body>");
    const lastSectPr = documentXml.lastIndexOf("<w:sectPr>");
    const lastTableEnd = documentXml.lastIndexOf("</w:tbl>");
    expect(lastSectPr).toBeGreaterThan(lastTableEnd);
    expect(bodyEnd).toBeGreaterThan(lastSectPr);
    // Nothing but the page-break paragraph and sectPr appears between the tables' end and body's end besides table content itself.
    const tailAfterLastTable = documentXml.slice(lastTableEnd + "</w:tbl>".length, bodyEnd);
    expect(tailAfterLastTable).toBe(documentXml.slice(lastSectPr, bodyEnd));

    // Content from both first and last product present.
    expect(documentXml).toContain("SKU-0");
    expect(documentXml).toContain("SKU-7");
  });

  it("rejects a template that fails fixed-grid validation", async () => {
    const templateBuffer = await buildSyntheticFixedGridDocx({ columns: 4, rows: 10 });
    const template = testTemplate();
    const plan = buildPlacements([product("p1", "SKU-1", "Widget", 100)], GEOMETRY, 1);

    await expect(renderFixedGridDocx(templateBuffer, template, plan)).rejects.toThrow(
      InvalidFixedGridTemplateError,
    );
  });
});

describe("renderFloatingDocx", () => {
  it("always throws UnsupportedFloatingTemplateError for a FLOATING template, producing no output", async () => {
    const floatingTemplate = testTemplate({
      id: "avery-22802",
      renderingMode: "FLOATING",
      columns: 1,
      rows: 8,
      labelsPerSheet: 8,
    });
    const plan = buildPlacements(
      [product("p1", "SKU-1", "Widget", 100)],
      { id: "avery-22802", columns: 1, rows: 8, labelsPerSheet: 8 },
      1,
    );

    await expect(renderFloatingDocx(Buffer.from(""), floatingTemplate, plan)).rejects.toThrow(
      UnsupportedFloatingTemplateError,
    );
  });
});
