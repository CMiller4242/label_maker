import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { buildPlacements } from "@label-maker/label-layout";
import type { LabelTemplate } from "@label-maker/shared";
import { renderFixedGridDocx } from "@label-maker/docx-renderer";
import { buildSyntheticFixedGridDocx } from "./helpers/synthetic-fixed-grid-docx.js";

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

describe("smoke: renderFixedGridDocx", () => {
  it("produces a re-openable docx with expected content", async () => {
    const templateBuffer = await buildSyntheticFixedGridDocx({ columns: 4, rows: 15 });
    const template = testTemplate();

    const products = [
      { id: "p1", sku: "SKU-1", description: "Widget One", priceCents: 1449, include: true },
    ];
    const plan = buildPlacements(
      products,
      { id: template.id, columns: 4, rows: 15, labelsPerSheet: 60 },
      8,
    );

    const result = await renderFixedGridDocx(templateBuffer, template, {
      labelTemplateId: template.id,
      copiesPerProduct: plan.copiesPerProduct,
      totalSheets: plan.totalSheets,
      totalPlacements: plan.totalPlacements,
      placements: plan.placements,
    });

    expect(result.mimeType).toBe(
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );

    const zip = await JSZip.loadAsync(result.buffer);
    const documentXml = await zip.file("word/document.xml")?.async("text");
    expect(documentXml).toBeTruthy();
    expect(documentXml).toContain("SKU-1");
    expect(documentXml).toContain("Widget One");
    expect(documentXml).toContain("As Low As: $14.49");
  });
});
