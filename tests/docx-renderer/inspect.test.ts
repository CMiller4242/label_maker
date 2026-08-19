import { readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { inspectDocxTemplate } from "@label-maker/docx-renderer";

const templatesDir = path.join(import.meta.dirname, "..", "..", "fixtures", "label-templates");

describe("inspectDocxTemplate against the two committed real fixtures", () => {
  it("avery-5155/original.docx: is a valid Word document, but a 7-column interleaved grid (not a uniform 4-column grid)", async () => {
    const buffer = readFileSync(path.join(templatesDir, "avery-5155", "original.docx"));
    const report = await inspectDocxTemplate({ buffer, filePath: "avery-5155/original.docx" });

    // See tests/docx-renderer/real-avery-5155.test.ts for the full
    // structural assertions. Real finding: the fixture is now a valid
    // Word document with one fixed-layout table, 15 rows - but its
    // <w:tblGrid> has 7 columns (4 real label columns interleaved with 3
    // vertically-merged spacer columns for horizontal pitch), not the
    // uniform 4 the current classifier/validator expect.
    expect(report.tableCount).toBe(1);
    expect(report.classification).toBe("AMBIGUOUS");
    expect(report.tables[0]?.totalWritableCells).toBe(105);
    expect(report.sha256).toMatch(/^[0-9a-f]{64}$/);
  });

  it("avery-22802/original.docx: identifies floating/tag-style layout", async () => {
    const buffer = readFileSync(path.join(templatesDir, "avery-22802", "original.docx"));
    const report = await inspectDocxTemplate({ buffer, filePath: "avery-22802/original.docx" });

    expect(report.tableCount).toBe(8);
    expect(report.tables.every((t) => t.positioning.isFloating)).toBe(true);
    expect(report.classification).toBe("AVERY_22802_LIKE_FLOATING");
    expect(report.classificationConfidence).toBeGreaterThan(0.5);

    // Real, measured page geometry (US Letter).
    expect(report.pageGeometry.widthTwips).toBe(12240);
    expect(report.pageGeometry.heightTwips).toBe(15840);
    expect(report.sha256).toMatch(/^[0-9a-f]{64}$/);
  });
});
