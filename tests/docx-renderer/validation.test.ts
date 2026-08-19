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

  it("loads the real avery-5155 fixture but rejects it as a 7-column interleaved grid, not a uniform 4-column grid", async () => {
    // The real fixture IS now a valid Word document (see real-avery-5155.test.ts
    // for the full inspection assertions) - it just doesn't match the
    // uniform-4-columns-per-row shape this validator currently requires
    // (it uses 4 real label columns interleaved with 3 vertically-merged
    // spacer columns for horizontal pitch, per a common Avery template
    // pattern).
    const buffer = readFileSync(path.join(templatesDir, "avery-5155", "original.docx"));
    const templatePackage = await loadTemplateDocx(buffer, "avery-5155");

    expect(() => validateFixedGridTemplate(templatePackage, EXPECTED)).toThrow(
      InvalidFixedGridTemplateError,
    );
    try {
      validateFixedGridTemplate(templatePackage, EXPECTED);
      expect.unreachable();
    } catch (error) {
      const issues = (error as InvalidFixedGridTemplateError).issues;
      expect(issues.some((i) => i.includes("7 columns"))).toBe(true);
    }
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
