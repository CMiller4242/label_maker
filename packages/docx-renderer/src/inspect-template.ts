import {
  InvalidDocxPackageError,
  childrenOf,
  extractText,
  findAllByTag,
  findDirectChildren,
  findFirstByTag,
  getAttr,
  loadTemplateDocx,
  sha256Hex,
  type XmlNode,
} from "./ooxml.js";
import type {
  CellInspection,
  CellMarginInspection,
  DocxInspectionReport,
  FixedGridPattern,
  FontSample,
  PageGeometry,
  ParagraphSample,
  RowInspection,
  TableInspection,
  TableClassificationInput,
  TablePositioning,
  TemplateClassification,
  WritableCellMapping,
} from "./types.js";

function parseIntAttr(node: XmlNode | undefined, attr: string): number | null {
  if (!node) return null;
  const raw = getAttr(node, attr);
  if (raw === undefined) return null;
  const parsed = Number.parseInt(raw, 10);
  return Number.isFinite(parsed) ? parsed : null;
}

function stringAttr(node: XmlNode | undefined, attr: string): string | null {
  if (!node) return null;
  return getAttr(node, attr) ?? null;
}

function firstDirectChild(nodes: XmlNode[], tag: string): XmlNode | undefined {
  return findDirectChildren(nodes, tag)[0];
}

function inspectPageGeometry(documentTree: XmlNode[], warnings: string[]): PageGeometry {
  const sectPr = findFirstByTag(documentTree, "w:sectPr");
  if (!sectPr) {
    warnings.push("No <w:sectPr> found: page size/margins are unknown.");
    return {
      widthTwips: null,
      heightTwips: null,
      marginTopTwips: null,
      marginBottomTwips: null,
      marginLeftTwips: null,
      marginRightTwips: null,
    };
  }

  const pgSz = firstDirectChild(childrenOf(sectPr), "w:pgSz");
  const pgMar = firstDirectChild(childrenOf(sectPr), "w:pgMar");
  if (!pgSz) warnings.push("<w:sectPr> has no <w:pgSz>: page width/height unknown.");
  if (!pgMar) warnings.push("<w:sectPr> has no <w:pgMar>: page margins unknown.");

  return {
    widthTwips: parseIntAttr(pgSz, "w:w"),
    heightTwips: parseIntAttr(pgSz, "w:h"),
    marginTopTwips: parseIntAttr(pgMar, "w:top"),
    marginBottomTwips: parseIntAttr(pgMar, "w:bottom"),
    marginLeftTwips: parseIntAttr(pgMar, "w:left"),
    marginRightTwips: parseIntAttr(pgMar, "w:right"),
  };
}

/** OOXML `<w:vMerge>` semantics: presence with no w:val (or val != "restart") means "continue" a merge started above. */
function inspectCellVMerge(tcPr: XmlNode | undefined): CellInspection["vMerge"] {
  if (!tcPr) return null;
  const vMerge = firstDirectChild(childrenOf(tcPr), "w:vMerge");
  if (!vMerge) return null;
  return stringAttr(vMerge, "w:val") === "restart" ? "restart" : "continue";
}

function inspectCells(row: XmlNode): CellInspection[] {
  const cellNodes = findDirectChildren(childrenOf(row), "w:tc");
  return cellNodes.map((tc, cellIndex) => {
    const tcPr = firstDirectChild(childrenOf(tc), "w:tcPr");
    const tcW = tcPr ? firstDirectChild(childrenOf(tcPr), "w:tcW") : undefined;
    return {
      cellIndex,
      widthTwips: parseIntAttr(tcW, "w:w"),
      vMerge: inspectCellVMerge(tcPr),
    };
  });
}

function inspectRows(
  tableChildren: XmlNode[],
  tableIndex: number,
  warnings: string[],
): RowInspection[] {
  const rowNodes = findDirectChildren(tableChildren, "w:tr");
  return rowNodes.map((row, rowIndex) => {
    const cells = inspectCells(row);
    const cellCount = cells.length;
    const trPr = firstDirectChild(childrenOf(row), "w:trPr");
    const trHeight = trPr ? firstDirectChild(childrenOf(trPr), "w:trHeight") : undefined;
    const cantSplitNode = trPr ? firstDirectChild(childrenOf(trPr), "w:cantSplit") : undefined;

    if (!trHeight) {
      warnings.push(
        `Table ${tableIndex} row ${rowIndex}: no explicit <w:trHeight> (row height rule unknown).`,
      );
    }
    if (!cantSplitNode) {
      warnings.push(`Table ${tableIndex} row ${rowIndex}: no <w:cantSplit> present.`);
    }

    const cantSplitVal = cantSplitNode ? stringAttr(cantSplitNode, "w:val") : null;
    const cantSplit =
      cantSplitNode !== undefined && cantSplitVal !== "0" && cantSplitVal !== "false";

    return {
      rowIndex,
      cellCount,
      heightTwips: parseIntAttr(trHeight, "w:val"),
      heightRule: stringAttr(trHeight, "w:hRule"),
      cantSplit,
      cells,
    };
  });
}

function inspectCellMargins(
  tblPr: XmlNode | undefined,
  sampleCellTcPr: XmlNode | undefined,
  tableIndex: number,
  warnings: string[],
): CellMarginInspection {
  const source =
    (sampleCellTcPr && firstDirectChild(childrenOf(sampleCellTcPr), "w:tcMar")) ||
    (tblPr && firstDirectChild(childrenOf(tblPr), "w:tblCellMar"));

  if (!source) {
    warnings.push(
      `Table ${tableIndex}: no cell margins found (neither per-cell w:tcMar nor table-level w:tblCellMar).`,
    );
    return { topTwips: null, bottomTwips: null, leftTwips: null, rightTwips: null };
  }

  const readSide = (side: string): number | null => {
    const el = firstDirectChild(childrenOf(source), side);
    return parseIntAttr(el, "w:w");
  };

  return {
    topTwips: readSide("w:top"),
    bottomTwips: readSide("w:bottom"),
    leftTwips: readSide("w:left"),
    rightTwips: readSide("w:right"),
  };
}

function inspectParagraphSample(
  sampleCell: XmlNode | undefined,
  tableIndex: number,
  warnings: string[],
): ParagraphSample | null {
  const paragraph = sampleCell && firstDirectChild(childrenOf(sampleCell), "w:p");
  const pPr = paragraph && firstDirectChild(childrenOf(paragraph), "w:pPr");
  if (!pPr) {
    warnings.push(
      `Table ${tableIndex}: sample cell has no <w:pPr> (paragraph spacing/alignment unknown).`,
    );
    return null;
  }
  const spacing = firstDirectChild(childrenOf(pPr), "w:spacing");
  const jc = firstDirectChild(childrenOf(pPr), "w:jc");
  if (!spacing) warnings.push(`Table ${tableIndex}: sample paragraph has no <w:spacing>.`);

  return {
    spacingBeforeTwips: parseIntAttr(spacing, "w:before"),
    spacingAfterTwips: parseIntAttr(spacing, "w:after"),
    lineSpacing: stringAttr(spacing, "w:line"),
    lineSpacingRule: stringAttr(spacing, "w:lineRule"),
    alignment: stringAttr(jc, "w:val"),
  };
}

function inspectFontSample(
  sampleCell: XmlNode | undefined,
  tableIndex: number,
  warnings: string[],
): FontSample | null {
  const paragraph = sampleCell && firstDirectChild(childrenOf(sampleCell), "w:p");
  const run = paragraph && firstDirectChild(childrenOf(paragraph), "w:r");
  const rPr = run && firstDirectChild(childrenOf(run), "w:rPr");
  if (!rPr) {
    warnings.push(
      `Table ${tableIndex}: sample cell has no run properties (font family/size/bold/color unknown).`,
    );
    return null;
  }
  const rFonts = firstDirectChild(childrenOf(rPr), "w:rFonts");
  const sz = firstDirectChild(childrenOf(rPr), "w:sz");
  const bold = firstDirectChild(childrenOf(rPr), "w:b");
  const color = firstDirectChild(childrenOf(rPr), "w:color");

  const boldVal = bold ? stringAttr(bold, "w:val") : null;

  return {
    fontFamily: stringAttr(rFonts, "w:ascii"),
    sizeHalfPoints: parseIntAttr(sz, "w:val"),
    bold: bold !== undefined && boldVal !== "0" && boldVal !== "false",
    colorHex: stringAttr(color, "w:val"),
  };
}

function inspectPositioning(tblPr: XmlNode | undefined): TablePositioning {
  const tblpPr = tblPr && firstDirectChild(childrenOf(tblPr), "w:tblpPr");
  return {
    isFloating: tblpPr !== undefined,
    anchorVertical: stringAttr(tblpPr, "w:vertAnchor"),
    anchorHorizontal: stringAttr(tblpPr, "w:horzAnchor"),
    positionXTwips: parseIntAttr(tblpPr, "w:tblpX"),
    positionYTwips: parseIntAttr(tblpPr, "w:tblpY"),
  };
}

/**
 * Detects which fixed-grid pattern (if any) a table's rows/columns match:
 *
 * - SIMPLE: every raw grid column is a writable label column (no cell uses
 *   `<w:vMerge>`).
 * - INTERLEAVED_SPACER: an odd number of raw grid columns, alternating
 *   wide (label) / narrow (spacer), where every spacer-position cell in
 *   every row carries `<w:vMerge>` (down the table for horizontal pitch)
 *   and no label-position cell does.
 *
 * Never guesses: any condition that fails is recorded in `diagnostics` and
 * the function returns `pattern: null` rather than accepting a near-miss.
 */
function detectFixedGridPattern(
  rows: RowInspection[],
  gridColumnWidthsTwips: number[],
): { pattern: FixedGridPattern | null; diagnostics: string[] } {
  const diagnostics: string[] = [];

  if (rows.length === 0) {
    diagnostics.push("Table has no rows.");
    return { pattern: null, diagnostics };
  }

  const rawColumnCounts = new Set(rows.map((r) => r.cellCount));
  if (rawColumnCounts.size !== 1) {
    diagnostics.push(
      `Rows do not have a uniform cell count (found: ${[...rawColumnCounts].join(", ")}); cannot determine a consistent column layout.`,
    );
    return { pattern: null, diagnostics };
  }
  const rawColumns = rows[0]?.cellCount ?? 0;
  if (rawColumns === 0) {
    diagnostics.push("Rows have zero cells.");
    return { pattern: null, diagnostics };
  }

  const anyVMerge = rows.some((r) => r.cells.some((c) => c.vMerge !== null));

  // --- SIMPLE: no vertical merging anywhere - every raw column is writable. ---
  if (!anyVMerge) {
    if (gridColumnWidthsTwips.length !== rawColumns) {
      diagnostics.push(
        `<w:tblGrid> declares ${gridColumnWidthsTwips.length} column(s) but rows have ${rawColumns} cell(s) each.`,
      );
      return { pattern: null, diagnostics };
    }
    const writableCellMap: WritableCellMapping[] = [];
    rows.forEach((row, rowIndex) => {
      for (let col = 0; col < rawColumns; col++) {
        writableCellMap.push({
          logicalSlotIndex: rowIndex * rawColumns + col,
          physicalRowIndex: rowIndex,
          physicalCellIndex: col,
        });
      }
    });
    return {
      pattern: {
        patternType: "SIMPLE",
        rawGridColumnWidthsTwips: gridColumnWidthsTwips,
        logicalLabelColumnIndexes: Array.from({ length: rawColumns }, (_, i) => i),
        spacerColumnIndexes: [],
        logicalColumns: rawColumns,
        logicalRows: rows.length,
        writableCellMap,
      },
      diagnostics: [],
    };
  }

  // --- INTERLEAVED_SPACER: odd raw column count, alternating wide/narrow,
  // narrow (odd-position) columns vertically merged, wide (even-position)
  // columns not. ---
  if (rawColumns % 2 !== 1 || rawColumns < 3) {
    diagnostics.push(
      `Table uses vertical cell merging but has ${rawColumns} raw column(s) (even, or fewer than 3) - ` +
        `the interleaved-spacer-column pattern requires an odd count >= 3 (label, spacer, label, ..., label).`,
    );
    return { pattern: null, diagnostics };
  }

  if (gridColumnWidthsTwips.length !== rawColumns) {
    diagnostics.push(
      `<w:tblGrid> declares ${gridColumnWidthsTwips.length} column(s) but rows have ${rawColumns} cell(s) each - cannot verify the wide/narrow pattern.`,
    );
    return { pattern: null, diagnostics };
  }

  const evenIndexes: number[] = [];
  const oddIndexes: number[] = [];
  for (let i = 0; i < rawColumns; i++) {
    (i % 2 === 0 ? evenIndexes : oddIndexes).push(i);
  }

  const evenWidths = new Set(evenIndexes.map((i) => gridColumnWidthsTwips[i]));
  const oddWidths = new Set(oddIndexes.map((i) => gridColumnWidthsTwips[i]));
  if (evenWidths.size !== 1) {
    diagnostics.push(
      `Label (even-position) column widths are not uniform: ${[...evenWidths].join(", ")} twips.`,
    );
  }
  if (oddWidths.size !== 1) {
    diagnostics.push(
      `Spacer (odd-position) column widths are not uniform: ${[...oddWidths].join(", ")} twips.`,
    );
  }
  if (evenWidths.size === 1 && oddWidths.size === 1) {
    const [evenWidth] = evenWidths;
    const [oddWidth] = oddWidths;
    if (evenWidth === undefined || oddWidth === undefined || !(oddWidth < evenWidth)) {
      diagnostics.push(
        `Expected narrower spacer columns (${oddWidth} twips) than label columns (${evenWidth} twips), but they are not narrower.`,
      );
    }
  }

  const mergeIssues: string[] = [];
  rows.forEach((row, rowIndex) => {
    for (const i of oddIndexes) {
      const cell = row.cells[i];
      if (!cell || cell.vMerge === null) {
        mergeIssues.push(`row ${rowIndex} spacer column ${i} has no <w:vMerge>`);
      }
    }
    for (const i of evenIndexes) {
      const cell = row.cells[i];
      if (cell && cell.vMerge !== null) {
        mergeIssues.push(`row ${rowIndex} label column ${i} unexpectedly has <w:vMerge>`);
      }
    }
  });
  if (mergeIssues.length > 0) {
    diagnostics.push(
      `Vertical-merge pattern is inconsistent (expected spacer columns merged, label columns not): ` +
        `${mergeIssues.slice(0, 5).join("; ")}${mergeIssues.length > 5 ? `; and ${mergeIssues.length - 5} more` : ""}.`,
    );
  }

  if (diagnostics.length > 0) {
    return { pattern: null, diagnostics };
  }

  const writableCellMap: WritableCellMapping[] = [];
  rows.forEach((row, rowIndex) => {
    evenIndexes.forEach((rawCol, logicalCol) => {
      writableCellMap.push({
        logicalSlotIndex: rowIndex * evenIndexes.length + logicalCol,
        physicalRowIndex: rowIndex,
        physicalCellIndex: rawCol,
      });
    });
  });

  return {
    pattern: {
      patternType: "INTERLEAVED_SPACER",
      rawGridColumnWidthsTwips: gridColumnWidthsTwips,
      logicalLabelColumnIndexes: evenIndexes,
      spacerColumnIndexes: oddIndexes,
      logicalColumns: evenIndexes.length,
      logicalRows: rows.length,
      writableCellMap,
    },
    diagnostics: [],
  };
}

function inspectTable(table: XmlNode, tableIndex: number, warnings: string[]): TableInspection {
  const tableChildren = childrenOf(table);
  const tblPr = firstDirectChild(tableChildren, "w:tblPr");
  if (!tblPr) warnings.push(`Table ${tableIndex}: no <w:tblPr> found.`);

  const tblLayout = tblPr && firstDirectChild(childrenOf(tblPr), "w:tblLayout");
  const layoutType = stringAttr(tblLayout, "w:type");
  if (!tblLayout) {
    warnings.push(
      `Table ${tableIndex}: no <w:tblLayout> found - Word defaults to "autofit", which is NOT what a fixed-grid label sheet needs.`,
    );
  }

  const tblW = tblPr && firstDirectChild(childrenOf(tblPr), "w:tblW");

  const tblGrid = firstDirectChild(tableChildren, "w:tblGrid");
  const gridColumnWidthsTwips = tblGrid
    ? findDirectChildren(childrenOf(tblGrid), "w:gridCol")
        .map((col) => parseIntAttr(col, "w:w"))
        .filter((w): w is number => w !== null)
    : [];
  if (!tblGrid) warnings.push(`Table ${tableIndex}: no <w:tblGrid> found (column widths unknown).`);

  const rows = inspectRows(tableChildren, tableIndex, warnings);
  const totalWritableCells = rows.reduce((sum, row) => sum + row.cellCount, 0);
  const { pattern: fixedGridPattern, diagnostics: fixedGridPatternDiagnostics } =
    detectFixedGridPattern(rows, gridColumnWidthsTwips);

  const firstRow = findDirectChildren(tableChildren, "w:tr")[0];
  const sampleCell = firstRow ? findDirectChildren(childrenOf(firstRow), "w:tc")[0] : undefined;
  const sampleCellTcPr = sampleCell
    ? firstDirectChild(childrenOf(sampleCell), "w:tcPr")
    : undefined;
  const vAlign = sampleCellTcPr
    ? firstDirectChild(childrenOf(sampleCellTcPr), "w:vAlign")
    : undefined;

  return {
    tableIndex,
    layoutType,
    isFixedLayout: layoutType === "fixed",
    tableWidthTwips: parseIntAttr(tblW, "w:w"),
    tableWidthType: stringAttr(tblW, "w:type"),
    gridColumnWidthsTwips,
    rowCount: rows.length,
    rows,
    totalWritableCells,
    cellMargins: inspectCellMargins(tblPr, sampleCellTcPr, tableIndex, warnings),
    cellVerticalAlignment: stringAttr(vAlign, "w:val"),
    paragraphSample: inspectParagraphSample(sampleCell, tableIndex, warnings),
    fontSample: inspectFontSample(sampleCell, tableIndex, warnings),
    positioning: inspectPositioning(tblPr),
    fixedGridPattern,
    fixedGridPatternDiagnostics,
  };
}

/**
 * Classifies a set of inspected tables as looking like a fixed 4x15/60-cell
 * grid (Avery 5155-style), a floating/tag layout (Avery 22802-style), or
 * ambiguous. This is a heuristic classification meant to catch obvious
 * mismatches early and inform a human reviewer - it is NOT a substitute for
 * validateFixedGridTemplate()'s strict, throwing validation used before
 * actually rendering.
 */
export function classifyTables(input: TableClassificationInput): {
  classification: TemplateClassification;
  confidence: number;
  reasons: string[];
} {
  const { tables } = input;
  const reasons: string[] = [];

  if (tables.length === 0) {
    return {
      classification: "AMBIGUOUS",
      confidence: 0,
      reasons: ["No tables found in document."],
    };
  }

  const inFlowTables = tables.filter((t) => !t.positioning.isFloating);
  const floatingTables = tables.filter((t) => t.positioning.isFloating);

  if (tables.length === 1 && inFlowTables.length === 1) {
    const table = inFlowTables[0];
    if (!table) return { classification: "AMBIGUOUS", confidence: 0, reasons: ["Internal error."] };
    const pattern = table.fixedGridPattern;

    if (
      pattern &&
      pattern.logicalRows === 15 &&
      pattern.logicalColumns === 4 &&
      pattern.writableCellMap.length === 60
    ) {
      let confidence = 0.7;
      if (table.isFixedLayout) confidence += 0.2;
      else reasons.push("Table layout is not explicitly fixed - Word may auto-fit column widths.");
      if (table.fixedGridPatternDiagnostics.length === 0) confidence += 0.1;
      else reasons.push("Fixed-grid pattern detection reported diagnostics despite matching.");
      return {
        classification: "AVERY_5155_LIKE_FIXED_GRID",
        confidence: Math.min(confidence, 1),
        reasons,
      };
    }

    if (pattern) {
      reasons.push(
        `Single in-flow table has ${pattern.logicalRows} logical rows / ${pattern.logicalColumns} logical ` +
          `columns / ${pattern.writableCellMap.length} writable cells (pattern: ${pattern.patternType}) - ` +
          `does not match the expected 15x4/60-cell fixed grid.`,
      );
    } else {
      reasons.push(
        `Single in-flow table with ${table.rowCount} rows does not match a recognized fixed-grid pattern: ` +
          `${table.fixedGridPatternDiagnostics.join(" ")}`,
      );
    }
  }

  if (floatingTables.length > 0 && floatingTables.length === tables.length) {
    reasons.push(`All ${tables.length} table(s) use floating/anchored positioning (<w:tblpPr>).`);
    return { classification: "AVERY_22802_LIKE_FLOATING", confidence: 0.85, reasons };
  }

  if (floatingTables.length > 0) {
    reasons.push(
      `${floatingTables.length} of ${tables.length} table(s) float and ${inFlowTables.length} do not - mixed layout.`,
    );
  }

  return { classification: "AMBIGUOUS", confidence: 0, reasons };
}

/**
 * Inspects an already-loaded template package (skips re-parsing document.xml
 * when a caller - e.g. template-validation.ts - already has one). Exported
 * so validateFixedGridTemplate() can reuse the exact same table inspection
 * logic without a redundant load/parse pass.
 */
export function inspectLoadedTemplate(
  templatePackage: { documentTree: XmlNode[] },
  meta: { filePath: string | null; sha256: string; byteSize: number },
): DocxInspectionReport {
  const warnings: string[] = [];
  const pageGeometry = inspectPageGeometry(templatePackage.documentTree, warnings);
  const tableNodes = findAllByTag(templatePackage.documentTree, "w:tbl");
  const tables = tableNodes.map((table, index) => inspectTable(table, index, warnings));

  const { classification, confidence, reasons } = classifyTables({ tables });
  warnings.push(...reasons);

  if (classification === "AMBIGUOUS") {
    warnings.push(
      "Classification is AMBIGUOUS: this template did not clearly match either the " +
        "Avery 5155 fixed-grid shape or the Avery 22802 floating-tag shape.",
    );
  }

  return {
    filePath: meta.filePath,
    sha256: meta.sha256,
    byteSize: meta.byteSize,
    pageGeometry,
    tableCount: tables.length,
    tables,
    classification,
    classificationConfidence: confidence,
    warnings,
  };
}

/**
 * Inspects a .docx package directly (JSZip + XML parsing only - no Word,
 * no LibreOffice, no conversion service) and produces a structured report:
 * page geometry, every table's layout/grid/row/cell/font/paragraph
 * properties, and a best-effort classification against the two known
 * preset shapes (Avery 5155 fixed-grid, Avery 22802 floating).
 */
export async function inspectDocxTemplate(input: {
  buffer: Buffer;
  filePath?: string;
}): Promise<DocxInspectionReport> {
  const sha256 = sha256Hex(input.buffer);
  const byteSize = input.buffer.byteLength;

  let templatePackage;
  try {
    templatePackage = await loadTemplateDocx(input.buffer, input.filePath ?? "(in-memory buffer)");
  } catch (error) {
    if (error instanceof InvalidDocxPackageError) {
      return {
        filePath: input.filePath ?? null,
        sha256,
        byteSize,
        pageGeometry: {
          widthTwips: null,
          heightTwips: null,
          marginTopTwips: null,
          marginBottomTwips: null,
          marginLeftTwips: null,
          marginRightTwips: null,
        },
        tableCount: 0,
        tables: [],
        classification: "AMBIGUOUS",
        classificationConfidence: 0,
        warnings: [error.message],
      };
    }
    throw error;
  }

  return inspectLoadedTemplate(templatePackage, {
    filePath: input.filePath ?? null,
    sha256,
    byteSize,
  });
}

/** Debug helper: dumps the raw text of every w:t in a table, for manual sanity checks. */
export function debugDumpTableText(table: XmlNode): string {
  return extractText(table);
}
