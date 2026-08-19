import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { buildPlacements } from "@label-maker/label-layout";
import type { LabelTemplate, Placement, PlacementPlan } from "@label-maker/shared";
import {
  childrenOf,
  extractText,
  findAllByTag,
  findDirectChildren,
  findFirstByTag,
  getAttr,
  inspectDocxTemplate,
  parseXml,
  renderFixedGridDocx,
  type XmlNode,
} from "@label-maker/docx-renderer";

/**
 * Exercises the real, committed fixtures/label-templates/avery-5155/original.docx
 * (not a synthetic stand-in) through the actual fixed-grid generation path.
 *
 * The real template is a genuinely valid Word document: one table, 15 rows,
 * fixed layout, exact row heights - and a <w:tblGrid> with 7 raw columns (4
 * real label columns interleaved with 3 narrow, vertically-merged
 * spacer/gutter columns, a common Avery Word-template pattern for correct
 * horizontal label pitch). validateFixedGridTemplate()/renderFixedGridDocx()
 * recognize this as the INTERLEAVED_SPACER fixed-grid pattern (4 logical
 * label columns x 15 logical rows = 60 writable cells) and generation
 * succeeds against this exact file. This file is not a synthetic-fixture
 * test (see generation.test.ts for those).
 */

const templatesDir = path.join(import.meta.dirname, "..", "..", "fixtures", "label-templates");
const artifactsDir = path.join(import.meta.dirname, "..", "..", "storage", "artifacts");

const WRITABLE_RAW_COLUMNS = [0, 2, 4, 6];
const SPACER_RAW_COLUMNS = [1, 3, 5];

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

async function renderAndReopen(templateBuffer: Buffer, template: LabelTemplate, plan: PlacementPlan) {
  const result = await renderFixedGridDocx(templateBuffer, template, plan);
  const zip = await JSZip.loadAsync(result.buffer);

  // Package validity: required parts still present.
  expect(zip.file("[Content_Types].xml")).toBeTruthy();
  expect(zip.file("_rels/.rels")).toBeTruthy();
  expect(zip.file("word/document.xml")).toBeTruthy();

  const documentXml = await zip.file("word/document.xml")?.async("text");
  if (!documentXml) throw new Error("word/document.xml missing from generated docx");

  const parser = new XMLParser({ ignoreAttributes: false });
  expect(() => parser.parse(documentXml)).not.toThrow();

  const tree = parseXml(documentXml);
  return { result, zip, documentXml, tree };
}

function getTables(tree: XmlNode[]): XmlNode[] {
  return findAllByTag(tree, "w:tbl");
}
function getRows(table: XmlNode): XmlNode[] {
  return findDirectChildren(childrenOf(table), "w:tr");
}
function getCells(row: XmlNode): XmlNode[] {
  return findDirectChildren(childrenOf(row), "w:tc");
}
function cellHasVMerge(cell: XmlNode): boolean {
  const tcPr = findFirstByTag(childrenOf(cell), "w:tcPr");
  if (!tcPr) return false;
  return findFirstByTag(childrenOf(tcPr), "w:vMerge") !== undefined;
}
function gridWidths(table: XmlNode): number[] {
  const tblGrid = findFirstByTag(childrenOf(table), "w:tblGrid");
  if (!tblGrid) return [];
  return findDirectChildren(childrenOf(tblGrid), "w:gridCol")
    .map((c) => getAttr(c, "w:w"))
    .filter((v): v is string => typeof v === "string")
    .map(Number);
}

/**
 * Asserts the full interleaved-grid contract against one generated sheet
 * table: writable cells (raw columns 0,2,4,6) hold exactly the expected
 * placement text or are blank, spacer/gutter cells (raw columns 1,3,5) are
 * never written to and keep their <w:vMerge>, and the raw 7-column grid
 * shape/fixed layout is preserved unchanged.
 */
function assertSheetTable(
  table: XmlNode,
  sheetNumber: number,
  sheetPlacements: Map<number, Placement>,
  allSkus: string[],
): void {
  expect(gridWidths(table)).toEqual([2520, 432, 2520, 432, 2520, 432, 2520]);

  const tblPr = findFirstByTag(childrenOf(table), "w:tblPr");
  const tblLayout = tblPr && findFirstByTag(childrenOf(tblPr), "w:tblLayout");
  expect(tblLayout?.[":@"]?.["@_w:type"]).toBe("fixed");

  const rows = getRows(table);
  expect(rows).toHaveLength(15);

  rows.forEach((row, rowIndex) => {
    const trPr = findFirstByTag(childrenOf(row), "w:trPr");
    expect(trPr).toBeTruthy();
    expect(findFirstByTag(childrenOf(trPr as XmlNode), "w:cantSplit")).toBeTruthy();

    const cells = getCells(row);
    expect(cells).toHaveLength(7);

    WRITABLE_RAW_COLUMNS.forEach((rawCol, logicalCol) => {
      const slotIndex = rowIndex * 4 + logicalCol;
      const cell = cells[rawCol];
      if (!cell) throw new Error(`Missing writable cell at row ${rowIndex} raw col ${rawCol}`);
      const text = extractText(cell);
      const placement = sheetPlacements.get(slotIndex);
      if (placement) {
        expect(text).toContain(placement.sku ?? "");
      } else {
        expect(text.trim()).toBe("");
      }
    });

    SPACER_RAW_COLUMNS.forEach((rawCol) => {
      const cell = cells[rawCol];
      if (!cell) throw new Error(`Missing spacer cell at row ${rowIndex} raw col ${rawCol}`);
      expect(cellHasVMerge(cell)).toBe(true);
      const text = extractText(cell);
      expect(text.trim()).toBe("");
      for (const sku of allSkus) {
        expect(text).not.toContain(sku);
      }
    });
  });

  void sheetNumber;
}

describe("real avery-5155/original.docx template", () => {
  it("is a valid Word document with the expected page/table geometry, recognized as an interleaved-spacer-column fixed grid", async () => {
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
    // vertically-merged spacer columns), 105 raw cells.
    expect(table?.gridColumnWidthsTwips).toEqual([2520, 432, 2520, 432, 2520, 432, 2520]);
    expect(table?.totalWritableCells).toBe(105);
    expect(table?.rows[0]?.heightTwips).toBe(950);
    expect(table?.rows[0]?.heightRule).toBe("exact");

    // Recognized as the INTERLEAVED_SPACER fixed-grid pattern -> classifies
    // as AVERY_5155_LIKE_FIXED_GRID, not AMBIGUOUS.
    expect(report.classification).toBe("AVERY_5155_LIKE_FIXED_GRID");
    const pattern = table?.fixedGridPattern;
    expect(pattern?.patternType).toBe("INTERLEAVED_SPACER");
    expect(pattern?.logicalLabelColumnIndexes).toEqual([0, 2, 4, 6]);
    expect(pattern?.spacerColumnIndexes).toEqual([1, 3, 5]);
    expect(pattern?.logicalColumns).toBe(4);
    expect(pattern?.logicalRows).toBe(15);
    expect(pattern?.writableCellMap).toHaveLength(60);
  });

  it("1 product x 8 copies: generates one real sheet, fills 8 writable labels, leaves 52 blank, never writes into spacer cells", async () => {
    const buffer = readFileSync(path.join(templatesDir, "avery-5155", "original.docx"));
    const products = [product("p1", "SKU-1", "Widget One", 999)];
    const plan = buildPlacements(products, { id: "avery-5155", columns: 4, rows: 15, labelsPerSheet: 60 }, 8);

    expect(plan.totalSheets).toBe(1);
    expect(plan.totalPlacements).toBe(8);

    const { documentXml, tree } = await renderAndReopen(buffer, realAveryTemplate(), plan);

    const tables = getTables(tree);
    expect(tables).toHaveLength(1);
    expect(documentXml).not.toContain('w:type="page"');

    const sheetPlacements = new Map(plan.placements.map((p) => [p.slotIndex, p]));
    assertSheetTable(tables[0] as XmlNode, 1, sheetPlacements, ["SKU-1"]);

    expect(documentXml).toContain("SKU-1");
    expect(documentXml).toContain("Widget One");
    expect(documentXml).toContain("As Low As: $9.99");

    // 8 populated, 52 blank writable label cells.
    expect(sheetPlacements.size).toBe(8);

    mkdirSync(artifactsDir, { recursive: true });
    writeFileSync(
      path.join(artifactsDir, "real-avery-5155-1-product.docx"),
      (await renderFixedGridDocx(buffer, realAveryTemplate(), plan)).buffer,
    );
  });

  it("8 products x 8 copies: generates two real sheets, sheet 1 fills 60 writable labels, sheet 2 fills 4, page break only between sheets", async () => {
    const buffer = readFileSync(path.join(templatesDir, "avery-5155", "original.docx"));
    const products = Array.from({ length: 8 }, (_, i) =>
      product(`p${i}`, `SKU-${i}`, `Widget ${i}`, 100 * (i + 1)),
    );
    const plan = buildPlacements(products, { id: "avery-5155", columns: 4, rows: 15, labelsPerSheet: 60 }, 8);

    expect(plan.totalSheets).toBe(2);
    const sheet1Placements = plan.placements.filter((p) => p.sheetNumber === 1);
    const sheet2Placements = plan.placements.filter((p) => p.sheetNumber === 2);
    expect(sheet1Placements).toHaveLength(60);
    expect(sheet2Placements).toHaveLength(4);

    const { result, documentXml, tree } = await renderAndReopen(buffer, realAveryTemplate(), plan);

    const tables = getTables(tree);
    expect(tables).toHaveLength(2);

    const allSkus = products.map((p) => p.sku);
    assertSheetTable(
      tables[0] as XmlNode,
      1,
      new Map(sheet1Placements.map((p) => [p.slotIndex, p])),
      allSkus,
    );
    assertSheetTable(
      tables[1] as XmlNode,
      2,
      new Map(sheet2Placements.map((p) => [p.slotIndex, p])),
      allSkus,
    );

    // Explicit page break exists exactly once, and only between the two
    // sheet tables (not inside either one).
    const pageBreakMatches = documentXml.match(/w:type="page"/g) ?? [];
    expect(pageBreakMatches).toHaveLength(1);
    const firstTableEnd = documentXml.indexOf("</w:tbl>");
    const secondTableStart = documentXml.indexOf("<w:tbl>", firstTableEnd);
    const pageBreakIndex = documentXml.indexOf('w:type="page"');
    expect(pageBreakIndex).toBeGreaterThan(firstTableEnd);
    expect(pageBreakIndex).toBeLessThan(secondTableStart);
    expect(documentXml.slice(0, firstTableEnd)).not.toContain('w:type="page"');

    // Every row across both tables (30 total) retains cantSplit; both
    // tables remain fixed layout.
    expect((documentXml.match(/<w:cantSplit\/?>/g) ?? []).length).toBe(30);
    expect((documentXml.match(/<w:tblLayout w:type="fixed"/g) ?? []).length).toBe(2);

    expect(documentXml).toContain("SKU-0");
    expect(documentXml).toContain("SKU-7");

    mkdirSync(artifactsDir, { recursive: true });
    writeFileSync(path.join(artifactsDir, "real-avery-5155-8-products.docx"), result.buffer);
  });
});
