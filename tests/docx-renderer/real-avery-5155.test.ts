import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { buildPlacements } from "@label-maker/label-layout";
import type { LabelTemplate } from "@label-maker/shared";
import {
  InvalidFixedGridTemplateError,
  inspectDocxTemplate,
  renderFixedGridDocx,
} from "@label-maker/docx-renderer";

/**
 * Exercises the real, committed fixtures/label-templates/avery-5155/original.docx
 * (not a synthetic stand-in) through the actual fixed-grid generation path.
 *
 * As of this file's writing, the real template is a genuinely valid Word
 * document (word/document.xml present, one table, 15 rows, fixed layout,
 * exact row heights) but its <w:tblGrid> has 7 columns - 4 real label
 * columns interleaved with 3 narrow, vertically-merged spacer/gutter
 * columns (a common Avery Word-template pattern for getting correct
 * horizontal label pitch) - rather than the uniform 4 columns per row that
 * validateFixedGridTemplate()/renderFixedGridDocx() currently require. So
 * generation against this exact file currently fails validation. This test
 * documents that real, current behavior; it is not a synthetic-fixture
 * test (see generation.test.ts for those) and this file must NOT be
 * "fixed" by loosening assertions - if it starts passing generation for a
 * different reason, or the failure mode changes, that's a real signal
 * worth investigating, not noise to silence.
 */

const templatesDir = path.join(import.meta.dirname, "..", "..", "fixtures", "label-templates");
const artifactsDir = path.join(import.meta.dirname, "..", "..", "storage", "artifacts");

function realAveryTemplate(): LabelTemplate {
  return {
    id: "avery-5155",
    displayName: "Avery 5155",
    renderingMode: "FIXED_GRID",
    columns: 4,
    rows: 15,
    labelsPerSheet: 60,
    templateStorageKey: "avery-5155/original.docx",
    templateVersion: "0.1.0-inspected",
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
    isPreset: true,
    createdAt: new Date(),
    updatedAt: new Date(),
  };
}

function product(id: string, sku: string, description: string, priceCents: number) {
  return { id, sku, description, priceCents, include: true };
}

describe("real avery-5155/original.docx template", () => {
  it("is a valid Word document with the expected page/table geometry", async () => {
    const buffer = readFileSync(path.join(templatesDir, "avery-5155", "original.docx"));
    const report = await inspectDocxTemplate({ buffer, filePath: "avery-5155/original.docx" });

    expect(report.tableCount).toBe(1);
    const table = report.tables[0];
    expect(table).toBeTruthy();
    expect(table?.isFixedLayout).toBe(true);
    expect(table?.rowCount).toBe(15);
    expect(table?.positioning.isFloating).toBe(false);

    // Real measured page geometry (US Letter).
    expect(report.pageGeometry.widthTwips).toBe(12240);
    expect(report.pageGeometry.heightTwips).toBe(15840);

    // Real measured table shape: 7 raw grid columns (4 label + 3
    // vertically-merged spacer columns), 105 raw cells, NOT the uniform
    // 60-cell/4-column shape the current validator expects.
    expect(table?.gridColumnWidthsTwips).toEqual([2520, 432, 2520, 432, 2520, 432, 2520]);
    expect(table?.totalWritableCells).toBe(105);
    expect(table?.rows[0]?.heightTwips).toBe(950);
    expect(table?.rows[0]?.heightRule).toBe("exact");

    // Does not match the simple uniform-grid model -> AMBIGUOUS, not
    // AVERY_5155_LIKE_FIXED_GRID, under today's classifier.
    expect(report.classification).toBe("AMBIGUOUS");
  });

  it.each([
    ["1 product x 8 copies", [product("p1", "SKU-1", "Widget One", 999)]],
    [
      "8 products x 8 copies",
      Array.from({ length: 8 }, (_, i) =>
        product(`p${i}`, `SKU-${i}`, `Widget ${i}`, 100 * (i + 1)),
      ),
    ],
  ])(
    "%s: renderFixedGridDocx currently fails validation against this real file (7-column interleaved grid, not a uniform 4-column grid)",
    async (_label, products) => {
      const buffer = readFileSync(path.join(templatesDir, "avery-5155", "original.docx"));
      const plan = buildPlacements(
        products,
        { id: "avery-5155", columns: 4, rows: 15, labelsPerSheet: 60 },
        8,
      );

      await expect(renderFixedGridDocx(buffer, realAveryTemplate(), plan)).rejects.toThrow(
        InvalidFixedGridTemplateError,
      );

      try {
        await renderFixedGridDocx(buffer, realAveryTemplate(), plan);
        expect.unreachable();
      } catch (error) {
        const issues = (error as InvalidFixedGridTemplateError).issues;
        expect(issues.some((i) => i.includes("105 cells"))).toBe(true);
        expect(issues.some((i) => i.includes("Expected <w:tblGrid> to declare 4"))).toBe(true);

        // Record the exact current failure for manual/CI inspection - no
        // DOCX artifact is produced since generation never gets past
        // validation. Written to gitignored local storage.
        mkdirSync(artifactsDir, { recursive: true });
        const reportPath = path.join(
          artifactsDir,
          `real-avery-5155-validation-failure-${products.length}-products.json`,
        );
        writeFileSync(
          reportPath,
          JSON.stringify({ errorName: (error as Error).name, issues }, null, 2),
        );
      }
    },
  );
});
