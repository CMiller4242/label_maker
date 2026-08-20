import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { createHash } from "node:crypto";
import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import { XMLParser, XMLValidator } from "fast-xml-parser";
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

  // Every XML/rels part in the package must be *strictly* well-formed XML -
  // not just re-parseable by the same lenient parser used to build it.
  // This is the exact class of check that catches a real-Word-rejects-the-
  // file defect (e.g. more than one XML declaration in word/document.xml)
  // that renderFixedGridDocx() itself also now asserts internally before
  // ever returning a buffer (see assertGeneratedDocxPackageIsValid() in
  // ooxml.ts) - checked again here, independently, at the test level.
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir || !(name.endsWith(".xml") || name.endsWith(".rels"))) continue;
    const text = await entry.async("text");
    const validation = XMLValidator.validate(text);
    if (validation !== true) {
      throw new Error(`"${name}" is not well-formed XML: ${validation.err.msg} (line ${validation.err.line}).`);
    }
    expect((text.match(/<\?xml/g) ?? []).length).toBe(1);
  }

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
function cellTcW(cell: XmlNode): number | null {
  const tcPr = findFirstByTag(childrenOf(cell), "w:tcPr");
  const tcW = tcPr && findFirstByTag(childrenOf(tcPr), "w:tcW");
  const val = tcW && getAttr(tcW, "w:w");
  return val !== undefined && val !== null ? Number(val) : null;
}
function firstParagraphJc(cell: XmlNode): string | null {
  const p = findFirstByTag(childrenOf(cell), "w:p");
  const pPr = p && findFirstByTag(childrenOf(p), "w:pPr");
  const jc = pPr && findFirstByTag(childrenOf(pPr), "w:jc");
  return jc ? (getAttr(jc, "w:val") ?? null) : null;
}
function cellVAlign(cell: XmlNode): string | null {
  const tcPr = findFirstByTag(childrenOf(cell), "w:tcPr");
  const vAlign = tcPr && findFirstByTag(childrenOf(tcPr), "w:vAlign");
  return vAlign ? (getAttr(vAlign, "w:val") ?? null) : null;
}

/**
 * Asserts the full interleaved-grid contract against one generated sheet
 * table: writable cells (raw columns 0,2,4,6) hold exactly the expected
 * placement text or are blank, spacer/gutter cells (raw columns 1,3,5) are
 * never written to and keep their <w:vMerge>, and the raw 7-column grid
 * shape/fixed layout is preserved unchanged.
 */
const SOURCE_GRID_WIDTHS = [2520, 432, 2520, 432, 2520, 432, 2520];
const SOURCE_TABLE_WIDTH_TWIPS = SOURCE_GRID_WIDTHS.reduce((sum, w) => sum + w, 0);

function assertSheetTable(
  table: XmlNode,
  sheetNumber: number,
  sheetPlacements: Map<number, Placement>,
  allSkus: string[],
): void {
  expect(gridWidths(table)).toEqual(SOURCE_GRID_WIDTHS);

  const tblPr = findFirstByTag(childrenOf(table), "w:tblPr");
  const tblLayout = tblPr && findFirstByTag(childrenOf(tblPr), "w:tblLayout");
  expect(tblLayout?.[":@"]?.["@_w:type"]).toBe("fixed");

  // The table's overall width must be an explicit dxa value equal to the
  // sum of its own gridCol widths - never "auto" (a well-documented Word
  // compatibility hazard for "fixed" layout tables: see
  // normalizeTableWidth() in fixed-grid-renderer.ts).
  const tblW = tblPr && findFirstByTag(childrenOf(tblPr), "w:tblW");
  expect(tblW && getAttr(tblW, "w:type")).toBe("dxa");
  expect(Number(tblW && getAttr(tblW, "w:w"))).toBe(SOURCE_TABLE_WIDTH_TWIPS);

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

      // Cell width unchanged from the source grid, and centering settings
      // present regardless of whether this specific cell got filled.
      expect(cellTcW(cell)).toBe(SOURCE_GRID_WIDTHS[rawCol]);
      expect(cellVAlign(cell)).toBe("center");

      const text = extractText(cell);
      const placement = sheetPlacements.get(slotIndex);
      if (placement) {
        expect(text).toContain(placement.sku ?? "");
        expect(firstParagraphJc(cell)).toBe("center");
      } else {
        expect(text.trim()).toBe("");
      }
    });

    SPACER_RAW_COLUMNS.forEach((rawCol) => {
      const cell = cells[rawCol];
      if (!cell) throw new Error(`Missing spacer cell at row ${rowIndex} raw col ${rawCol}`);
      expect(cellHasVMerge(cell)).toBe(true);
      expect(cellTcW(cell)).toBe(SOURCE_GRID_WIDTHS[rawCol]);
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

    const { result, documentXml, tree } = await renderAndReopen(buffer, realAveryTemplate(), plan);

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

    // Persist the exact buffer that was just validated above - not a
    // separate re-render - so the artifact on disk is guaranteed to match
    // what this test actually checked.
    mkdirSync(artifactsDir, { recursive: true });
    writeFileSync(path.join(artifactsDir, "real-avery-5155-1-product.docx"), result.buffer);
  });

  it("renders byte-identical output for identical inputs (reproducible artifacts, not wall-clock-dependent)", async () => {
    // Regression test: renderFixedGridDocx() used to stamp the current
    // wall-clock time into the rewritten word/document.xml zip entry (and
    // JSZip would silently synthesize an extra "word/" directory entry,
    // also wall-clock-stamped) - so two renders of the exact same template
    // + placement plan produced different bytes/sha256 whenever they
    // landed in different (2-second-granularity) ticks, even though the
    // actual document content was identical. Persisted/cached artifacts
    // for the same input must be reproducible, not flaky by run timing.
    const buffer = readFileSync(path.join(templatesDir, "avery-5155", "original.docx"));
    const products = [product("p1", "SKU-1", "Widget One", 999)];
    const plan = buildPlacements(products, { id: "avery-5155", columns: 4, rows: 15, labelsPerSheet: 60 }, 8);

    const first = await renderFixedGridDocx(buffer, realAveryTemplate(), plan);
    await new Promise((resolve) => setTimeout(resolve, 2200)); // cross a DOS-date 2s boundary
    const second = await renderFixedGridDocx(buffer, realAveryTemplate(), plan);

    expect(createHash("sha256").update(second.buffer).digest("hex")).toBe(
      createHash("sha256").update(first.buffer).digest("hex"),
    );

    // No entries were silently added beyond the source template's own.
    const sourceZip = await JSZip.loadAsync(buffer);
    const outputZip = await JSZip.loadAsync(first.buffer);
    expect(Object.keys(outputZip.files).sort()).toEqual(Object.keys(sourceZip.files).sort());
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

  it("preserves the source fixture's own table/cell geometry exactly (structural comparison, not hardcoded literals)", async () => {
    // Directly compares the generated output's geometry against the real
    // source fixture's own inspected values - table count, cell count,
    // tblGrid/tcW widths, and fixed-layout settings must be identical to
    // what the source template itself declares, not merely equal to a
    // value this test happens to hardcode.
    const buffer = readFileSync(path.join(templatesDir, "avery-5155", "original.docx"));
    const sourceReport = await inspectDocxTemplate({ buffer, filePath: "avery-5155/original.docx" });
    const sourceTable = sourceReport.tables[0];
    if (!sourceTable) throw new Error("Expected the source fixture to have a table.");

    const products = [product("p1", "SKU-1", "Widget One", 999)];
    const plan = buildPlacements(products, { id: "avery-5155", columns: 4, rows: 15, labelsPerSheet: 60 }, 8);
    const { tree } = await renderAndReopen(buffer, realAveryTemplate(), plan);

    const tables = getTables(tree);
    expect(tables).toHaveLength(sourceReport.tableCount);

    const table = tables[0] as XmlNode;
    const rows = getRows(table);
    expect(rows).toHaveLength(sourceTable.rowCount);
    const totalCells = rows.reduce((sum, row) => sum + getCells(row).length, 0);
    expect(totalCells).toBe(sourceTable.totalWritableCells); // "writable" here means raw cell count, incl. spacers

    expect(gridWidths(table)).toEqual(sourceTable.gridColumnWidthsTwips);

    const tblPr = findFirstByTag(childrenOf(table), "w:tblPr");
    const tblLayout = tblPr && findFirstByTag(childrenOf(tblPr), "w:tblLayout");
    expect(tblLayout && getAttr(tblLayout, "w:type")).toBe(sourceTable.layoutType);
    expect(sourceTable.isFixedLayout).toBe(true); // sanity: the source really does declare fixed layout

    // Every generated cell's own tcW matches the source's per-row-index
    // tcW exactly (source rows are already confirmed uniform elsewhere).
    const sourceRow0 = sourceTable.rows[0];
    if (!sourceRow0) throw new Error("Expected the source fixture's table to have at least one row.");
    rows.forEach((row) => {
      const cells = getCells(row);
      cells.forEach((cell, cellIndex) => {
        const sourceCellWidth = sourceRow0.cells[cellIndex]?.widthTwips;
        expect(cellTcW(cell)).toBe(sourceCellWidth);
      });
    });
  });
});
