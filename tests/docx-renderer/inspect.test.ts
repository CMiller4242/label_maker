import { readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { inspectDocxTemplate } from "@label-maker/docx-renderer";

const templatesDir = path.join(import.meta.dirname, "..", "..", "fixtures", "label-templates");

describe("inspectDocxTemplate against the two committed real fixtures", () => {
  it("avery-5155/original.docx: reports it is NOT a usable Word document", async () => {
    const buffer = readFileSync(path.join(templatesDir, "avery-5155", "original.docx"));
    const report = await inspectDocxTemplate({ buffer, filePath: "avery-5155/original.docx" });

    // This is the actual, honest finding for the currently-committed
    // fixture: it is an Office theme package (.thmx-shaped zip), not a
    // Word document - it has no word/document.xml part at all.
    expect(report.tableCount).toBe(0);
    expect(report.classification).toBe("AMBIGUOUS");
    expect(report.classificationConfidence).toBe(0);
    expect(report.warnings.some((w) => w.includes("word/document.xml"))).toBe(true);
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
