import { readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import {
  InvalidFixedGridTemplateError,
  loadTemplateDocx,
  validateFixedGridTemplate,
} from "@label-maker/docx-renderer";
import { buildSyntheticFixedGridDocx } from "./helpers/synthetic-fixed-grid-docx.js";

const templatesDir = path.join(import.meta.dirname, "..", "..", "fixtures", "label-templates");
const EXPECTED = { columns: 4, rows: 15, labelsPerSheet: 60 };

describe("validateFixedGridTemplate", () => {
  it("recognizes a valid 4x15/60-cell fixed grid", async () => {
    const buffer = await buildSyntheticFixedGridDocx({ columns: 4, rows: 15 });
    const templatePackage = await loadTemplateDocx(buffer, "synthetic-valid");

    const result = validateFixedGridTemplate(templatePackage, EXPECTED);

    expect(result.columns).toBe(4);
    expect(result.rows).toBe(15);
    expect(result.labelsPerSheet).toBe(60);
    expect(result.inspection.isFixedLayout).toBe(true);
    expect(result.inspection.totalWritableCells).toBe(60);
    expect(result.inspection.positioning.isFloating).toBe(false);
  });

  it("rejects a grid with the wrong row count", async () => {
    const buffer = await buildSyntheticFixedGridDocx({ columns: 4, rows: 14 });
    const templatePackage = await loadTemplateDocx(buffer, "synthetic-wrong-rows");

    expect(() => validateFixedGridTemplate(templatePackage, EXPECTED)).toThrow(
      InvalidFixedGridTemplateError,
    );
    try {
      validateFixedGridTemplate(templatePackage, EXPECTED);
      expect.unreachable();
    } catch (error) {
      expect(error).toBeInstanceOf(InvalidFixedGridTemplateError);
      const issues = (error as InvalidFixedGridTemplateError).issues;
      expect(issues.some((i) => i.includes("Expected 15 rows, found 14"))).toBe(true);
    }
  });

  it("rejects a grid with the wrong column count", async () => {
    const buffer = await buildSyntheticFixedGridDocx({ columns: 5, rows: 15 });
    const templatePackage = await loadTemplateDocx(buffer, "synthetic-wrong-columns");

    expect(() => validateFixedGridTemplate(templatePackage, EXPECTED)).toThrow(
      InvalidFixedGridTemplateError,
    );
  });

  it("loads the real avery-5155 fixture and accepts it as a supported interleaved-spacer-column fixed grid", async () => {
    // The real fixture is a valid Word document (see real-avery-5155.test.ts
    // for the full inspection assertions): 4 real label columns interleaved
    // with 3 vertically-merged spacer columns for horizontal pitch, per a
    // common Avery template pattern. validateFixedGridTemplate() recognizes
    // this INTERLEAVED_SPACER shape and returns logical 4x15 placement
    // metadata rather than rejecting it.
    const buffer = readFileSync(path.join(templatesDir, "avery-5155", "original.docx"));
    const templatePackage = await loadTemplateDocx(buffer, "avery-5155");

    const result = validateFixedGridTemplate(templatePackage, EXPECTED);

    expect(result.columns).toBe(4);
    expect(result.rows).toBe(15);
    expect(result.labelsPerSheet).toBe(60);
    expect(result.inspection.isFixedLayout).toBe(true);
    expect(result.inspection.positioning.isFloating).toBe(false);
    expect(result.pattern.patternType).toBe("INTERLEAVED_SPACER");
    expect(result.pattern.rawGridColumnWidthsTwips).toEqual([2520, 432, 2520, 432, 2520, 432, 2520]);
    expect(result.pattern.logicalLabelColumnIndexes).toEqual([0, 2, 4, 6]);
    expect(result.pattern.spacerColumnIndexes).toEqual([1, 3, 5]);
    expect(result.pattern.writableCellMap).toHaveLength(60);

    // Explicit logical slot -> physical (row, raw writable column) mapping,
    // per the interleaved layout's expected pitch.
    const bySlot = new Map(result.pattern.writableCellMap.map((m) => [m.logicalSlotIndex, m]));
    expect(bySlot.get(0)).toMatchObject({ physicalRowIndex: 0, physicalCellIndex: 0 });
    expect(bySlot.get(1)).toMatchObject({ physicalRowIndex: 0, physicalCellIndex: 2 });
    expect(bySlot.get(2)).toMatchObject({ physicalRowIndex: 0, physicalCellIndex: 4 });
    expect(bySlot.get(3)).toMatchObject({ physicalRowIndex: 0, physicalCellIndex: 6 });
    expect(bySlot.get(4)).toMatchObject({ physicalRowIndex: 1, physicalCellIndex: 0 });
  });

  it("rejects the real avery-22802 fixture (floating, not a fixed grid)", async () => {
    const buffer = readFileSync(path.join(templatesDir, "avery-22802", "original.docx"));
    const templatePackage = await loadTemplateDocx(buffer, "avery-22802");

    expect(() => validateFixedGridTemplate(templatePackage, EXPECTED)).toThrow(
      InvalidFixedGridTemplateError,
    );
    try {
      validateFixedGridTemplate(templatePackage, EXPECTED);
      expect.unreachable();
    } catch (error) {
      const issues = (error as InvalidFixedGridTemplateError).issues;
      expect(issues.some((i) => i.toLowerCase().includes("floating"))).toBe(true);
    }
  });
});
