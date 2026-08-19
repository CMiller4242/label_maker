import type { LabelTemplate } from "@label-maker/shared";
import { findAllByTag, type TemplatePackage, type XmlNode } from "./ooxml.js";
import { inspectLoadedTemplate } from "./inspect-template.js";
import {
  InvalidFixedGridTemplateError,
  type ExpectedFixedGridGeometry,
  type TableInspection,
  type TemplateLayoutValidationResult,
  type ValidatedFixedGridTemplate,
} from "./types.js";

/**
 * Locates the actual `w:tbl` XML node for a table index produced by
 * inspection. Uses the identical traversal order as inspectLoadedTemplate()
 * (findAllByTag over the same document tree), so `tableIndex` always
 * addresses the same node.
 */
export function findTableNode(documentTree: XmlNode[], tableIndex: number): XmlNode {
  const tables = findAllByTag(documentTree, "w:tbl");
  const table = tables[tableIndex];
  if (!table) {
    throw new InvalidFixedGridTemplateError("(unknown)", [
      `Table index ${tableIndex} not found (document has ${tables.length} table(s)).`,
    ]);
  }
  return table;
}

function describeTable(table: TableInspection): string {
  const columnCounts = new Set(table.rows.map((r) => r.cellCount));
  const columns = columnCounts.size === 1 ? String(table.rows[0]?.cellCount ?? "?") : "non-uniform";
  return `${table.rowCount} rows x ${columns} columns (${table.totalWritableCells} cells)`;
}

/**
 * Strictly validates that a loaded template's document actually contains a
 * single, normal in-flow, fixed-layout table matching `expected`'s
 * columns/rows/labelsPerSheet exactly. Throws InvalidFixedGridTemplateError
 * with a full list of concrete, actionable issues if it does not - this is
 * the gate renderFixedGridDocx() must pass before it will touch a template.
 *
 * On success, returns a descriptor identifying exactly which table to clone
 * (by index, using the same document-order traversal as inspection), so the
 * renderer never has to re-guess which table is "the" label grid.
 */
export function validateFixedGridTemplate(
  templatePackage: TemplatePackage,
  expected: ExpectedFixedGridGeometry,
): ValidatedFixedGridTemplate {
  const report = inspectLoadedTemplate(templatePackage, {
    filePath: templatePackage.templateStorageKey,
    sha256: "",
    byteSize: 0,
  });

  const issues: string[] = [];

  if (report.tableCount === 0) {
    issues.push("Document contains no tables at all.");
    throw new InvalidFixedGridTemplateError(templatePackage.templateStorageKey, issues);
  }

  const inFlowTables = report.tables.filter((t) => !t.positioning.isFloating);
  if (inFlowTables.length === 0) {
    issues.push(
      `Document has ${report.tableCount} table(s), but all of them use floating/anchored ` +
        `positioning (<w:tblpPr>) - a fixed grid needs a single normal in-flow table.`,
    );
    throw new InvalidFixedGridTemplateError(templatePackage.templateStorageKey, issues);
  }
  if (inFlowTables.length > 1) {
    issues.push(
      `Found ${inFlowTables.length} in-flow tables (indices ${inFlowTables
        .map((t) => t.tableIndex)
        .join(", ")}); expected exactly 1 unambiguous label grid table.`,
    );
    throw new InvalidFixedGridTemplateError(templatePackage.templateStorageKey, issues);
  }

  const table = inFlowTables[0];
  if (!table) {
    throw new InvalidFixedGridTemplateError(templatePackage.templateStorageKey, [
      "Internal error selecting table.",
    ]);
  }

  if (!table.isFixedLayout) {
    issues.push(
      `Table layout is "${table.layoutType ?? "(unspecified, defaults to autofit)"}", expected "fixed". ` +
        `Autofit layout lets Word resize columns based on content, which breaks label alignment.`,
    );
  }
  if (table.positioning.isFloating) {
    issues.push(
      "Table uses floating/anchored positioning (<w:tblpPr>); expected normal in-flow placement.",
    );
  }

  const pattern = table.fixedGridPattern;
  if (!pattern) {
    issues.push(
      `No recognized fixed-grid pattern (neither a simple uniform grid nor an interleaved-spacer-column ` +
        `grid): ${table.fixedGridPatternDiagnostics.join(" ") || "no further diagnostics available."}`,
    );
    issues.unshift(`Actual table shape: ${describeTable(table)}.`);
    throw new InvalidFixedGridTemplateError(templatePackage.templateStorageKey, issues);
  }

  if (pattern.logicalRows !== expected.rows) {
    issues.push(`Expected ${expected.rows} rows, found ${pattern.logicalRows}.`);
  }
  if (pattern.logicalColumns !== expected.columns) {
    issues.push(`Expected ${expected.columns} columns, found ${pattern.logicalColumns}.`);
  }
  if (pattern.writableCellMap.length !== expected.labelsPerSheet) {
    issues.push(
      `Expected ${expected.labelsPerSheet} total writable cells, found ${pattern.writableCellMap.length}.`,
    );
  }

  const overlap = pattern.logicalLabelColumnIndexes.filter((i) =>
    pattern.spacerColumnIndexes.includes(i),
  );
  if (overlap.length > 0) {
    issues.push(
      `Internal inconsistency: raw column(s) ${overlap.join(", ")} are classified as both writable and spacer.`,
    );
  }

  const slotCounts = new Map<number, number>();
  const cellKeyCounts = new Map<string, number>();
  for (const mapping of pattern.writableCellMap) {
    slotCounts.set(mapping.logicalSlotIndex, (slotCounts.get(mapping.logicalSlotIndex) ?? 0) + 1);
    const cellKey = `${mapping.physicalRowIndex}:${mapping.physicalCellIndex}`;
    cellKeyCounts.set(cellKey, (cellKeyCounts.get(cellKey) ?? 0) + 1);
  }
  const duplicateSlots = [...slotCounts.entries()].filter(([, count]) => count > 1);
  if (duplicateSlots.length > 0) {
    issues.push(
      `writableCellMap has duplicate logical slot index(es): ${duplicateSlots.map(([slot]) => slot).join(", ")}.`,
    );
  }
  const duplicateCells = [...cellKeyCounts.entries()].filter(([, count]) => count > 1);
  if (duplicateCells.length > 0) {
    issues.push(
      `writableCellMap maps more than one logical slot onto the same physical cell: ${duplicateCells
        .map(([key]) => key)
        .join(", ")}.`,
    );
  }
  if (pattern.logicalRows === expected.rows && pattern.logicalColumns === expected.columns) {
    const expectedSlotCount = expected.rows * expected.columns;
    const missingSlots: number[] = [];
    for (let slot = 0; slot < expectedSlotCount; slot++) {
      if (!slotCounts.has(slot)) missingSlots.push(slot);
    }
    if (missingSlots.length > 0) {
      issues.push(
        `writableCellMap is missing logical slot(s): ${missingSlots.slice(0, 5).join(", ")}` +
          `${missingSlots.length > 5 ? `, and ${missingSlots.length - 5} more` : ""}.`,
      );
    }
  }

  const declaredRowHeights = new Set(
    table.rows.map((r) => r.heightTwips).filter((h): h is number => h !== null),
  );
  if (declaredRowHeights.size > 1) {
    issues.push(
      `Rows declare inconsistent explicit heights: ${[...declaredRowHeights].join(", ")} twips - ` +
        `expected a single uniform row height rule.`,
    );
  }

  if (issues.length > 0) {
    issues.unshift(
      `Actual table shape: ${describeTable(table)} (pattern: ${pattern.patternType}, ` +
        `${pattern.logicalColumns} logical columns x ${pattern.logicalRows} logical rows).`,
    );
    throw new InvalidFixedGridTemplateError(templatePackage.templateStorageKey, issues);
  }

  return {
    templateStorageKey: templatePackage.templateStorageKey,
    tableIndex: table.tableIndex,
    columns: expected.columns,
    rows: expected.rows,
    labelsPerSheet: expected.labelsPerSheet,
    inspection: table,
    pattern,
  };
}

/**
 * Non-throwing adapter over validateFixedGridTemplate(), matching the
 * LabelRenderer interface's validateTemplateLayout() shape. FLOATING
 * templates are only checked for having at least one table; they are never
 * validated as a fixed grid (that would always fail, by design).
 */
export function validateTemplateLayout(
  templatePackage: TemplatePackage,
  template: LabelTemplate,
): TemplateLayoutValidationResult {
  if (template.renderingMode === "FLOATING") {
    const report = inspectLoadedTemplate(templatePackage, {
      filePath: templatePackage.templateStorageKey,
      sha256: "",
      byteSize: 0,
    });
    if (report.tableCount === 0) {
      return {
        valid: false,
        issues: [{ code: "NO_TABLES", message: "FLOATING template contains no tables." }],
      };
    }
    return { valid: true, issues: [] };
  }

  try {
    validateFixedGridTemplate(templatePackage, {
      columns: template.columns,
      rows: template.rows,
      labelsPerSheet: template.labelsPerSheet,
    });
    return { valid: true, issues: [] };
  } catch (error) {
    if (error instanceof InvalidFixedGridTemplateError) {
      return {
        valid: false,
        issues: error.issues.map((message) => ({ code: "FIXED_GRID_VALIDATION", message })),
      };
    }
    throw error;
  }
}
