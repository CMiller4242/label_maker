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

  if (table.rowCount !== expected.rows) {
    issues.push(`Expected ${expected.rows} rows, found ${table.rowCount}.`);
  }
  const nonMatchingRows = table.rows.filter((r) => r.cellCount !== expected.columns);
  if (nonMatchingRows.length > 0) {
    issues.push(
      `Expected every row to have exactly ${expected.columns} cells; ` +
        `${nonMatchingRows.length} row(s) do not (e.g. row ${nonMatchingRows[0]?.rowIndex} has ${nonMatchingRows[0]?.cellCount}).`,
    );
  }
  if (table.totalWritableCells !== expected.labelsPerSheet) {
    issues.push(
      `Expected ${expected.labelsPerSheet} total writable cells, found ${table.totalWritableCells}.`,
    );
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
  if (table.gridColumnWidthsTwips.length !== expected.columns) {
    issues.push(
      `Expected <w:tblGrid> to declare ${expected.columns} explicit column width(s), found ${table.gridColumnWidthsTwips.length}.`,
    );
  }

  if (issues.length > 0) {
    issues.unshift(`Actual table shape: ${describeTable(table)}.`);
    throw new InvalidFixedGridTemplateError(templatePackage.templateStorageKey, issues);
  }

  return {
    templateStorageKey: templatePackage.templateStorageKey,
    tableIndex: table.tableIndex,
    columns: expected.columns,
    rows: expected.rows,
    labelsPerSheet: expected.labelsPerSheet,
    inspection: table,
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
