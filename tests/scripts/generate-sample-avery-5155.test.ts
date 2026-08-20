import { mkdtempSync, readFileSync, rmSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { createHash } from "node:crypto";
import { afterEach, describe, expect, it } from "vitest";
import JSZip from "jszip";
import { XMLValidator } from "fast-xml-parser";
import {
  SAMPLE_SCENARIOS,
  generateSampleArtifact,
  sampleAvery5155Template,
} from "../../scripts/generate-sample-avery-5155.js";

/**
 * Exercises generateSampleArtifact() - the reusable core behind
 * `pnpm docx:sample-5155` - against the real committed Avery 5155 source
 * template. Every file this test produces is written to a fresh temporary
 * directory (never storage/artifacts) and removed in afterEach, so running
 * this suite never leaves stray files behind.
 */

const templatesDir = path.join(import.meta.dirname, "..", "..", "fixtures", "label-templates");
const templateBuffer = readFileSync(path.join(templatesDir, "avery-5155", "original.docx"));

let tmpDir: string;

afterEach(() => {
  if (tmpDir) rmSync(tmpDir, { recursive: true, force: true });
});

describe("generateSampleArtifact", () => {
  it("renders the real 1-product scenario to a real .docx + metadata JSON on disk", async () => {
    tmpDir = mkdtempSync(path.join(tmpdir(), "avery-5155-sample-"));
    const scenario = SAMPLE_SCENARIOS.find((s) => s.fileBaseName === "sample-avery-5155-1-product");
    if (!scenario) throw new Error("Expected scenario not found in SAMPLE_SCENARIOS.");

    const artifact = await generateSampleArtifact(templateBuffer, scenario, tmpDir);

    expect(artifact.docxPath).toBe(path.join(tmpDir, "sample-avery-5155-1-product.docx"));
    expect(artifact.metadataPath).toBe(path.join(tmpDir, "sample-avery-5155-1-product.metadata.json"));
    expect(artifact.sheetCount).toBe(1);
    expect(artifact.filledSlotCount).toBe(8);
    expect(artifact.blankSlotCount).toBe(52);

    // The file genuinely exists on disk with the reported size/hash - not
    // just an in-memory result.
    const onDiskBytes = readFileSync(artifact.docxPath);
    expect(statSync(artifact.docxPath).size).toBe(artifact.byteSize);
    expect(createHash("sha256").update(onDiskBytes).digest("hex")).toBe(artifact.sha256);

    const metadata = JSON.parse(readFileSync(artifact.metadataPath, "utf8")) as Record<string, unknown>;
    expect(metadata.sha256).toBe(artifact.sha256);
    expect(metadata.byteSize).toBe(artifact.byteSize);
    expect(metadata.sheetCount).toBe(1);
    expect(metadata.filledSlotCount).toBe(8);
    expect(metadata.blankSlotCount).toBe(52);
  });

  it("renders the real 8-products scenario spanning two sheets", async () => {
    tmpDir = mkdtempSync(path.join(tmpdir(), "avery-5155-sample-"));
    const scenario = SAMPLE_SCENARIOS.find((s) => s.fileBaseName === "sample-avery-5155-8-products");
    if (!scenario) throw new Error("Expected scenario not found in SAMPLE_SCENARIOS.");
    expect(scenario.products).toHaveLength(8);

    const artifact = await generateSampleArtifact(templateBuffer, scenario, tmpDir);

    expect(artifact.sheetCount).toBe(2);
    expect(artifact.filledSlotCount).toBe(64); // 8 products x 8 copies
    expect(artifact.blankSlotCount).toBe(56); // 2 sheets x 60 - 64

    const onDiskBytes = readFileSync(artifact.docxPath);
    expect(onDiskBytes.byteLength).toBe(artifact.byteSize);
  });

  it("overwrites deterministically: re-running the same scenario produces byte-identical output", async () => {
    tmpDir = mkdtempSync(path.join(tmpdir(), "avery-5155-sample-"));
    const scenario = SAMPLE_SCENARIOS[0];
    if (!scenario) throw new Error("Expected at least one sample scenario.");

    const first = await generateSampleArtifact(templateBuffer, scenario, tmpDir);
    const second = await generateSampleArtifact(templateBuffer, scenario, tmpDir);

    expect(second.docxPath).toBe(first.docxPath);
    expect(second.sha256).toBe(first.sha256);
    expect(readFileSync(second.docxPath)).toEqual(readFileSync(first.docxPath));
  });

  it("generates a .docx openable by Microsoft Word - valid zip, every XML/rels part well-formed with exactly one declaration", async () => {
    // Regression test tied to the precise root cause of a real bug: a
    // developer opened storage/artifacts/sample-avery-5155-1-product.docx
    // in desktop Word on Windows and got "Word experienced an error trying
    // to open the file." The package was a structurally valid zip (unzip
    // -t passed) whose word/document.xml re-parsed fine with the same
    // lenient parser used to build it - but it actually contained TWO
    // <?xml ...?> declarations (buildXml() failed to strip the
    // declaration node parseXml() captures from the source template),
    // which is invalid XML that only a strict validator (or Word itself)
    // catches. This test opens the exact bytes this command persists to
    // disk and validates them the same strict way Word would reject them.
    tmpDir = mkdtempSync(path.join(tmpdir(), "avery-5155-sample-"));
    const scenario = SAMPLE_SCENARIOS[0];
    if (!scenario) throw new Error("Expected at least one sample scenario.");

    const artifact = await generateSampleArtifact(templateBuffer, scenario, tmpDir);
    const onDiskBytes = readFileSync(artifact.docxPath);

    const zip = await JSZip.loadAsync(onDiskBytes); // throws if not a valid zip
    const requiredParts = ["[Content_Types].xml", "_rels/.rels", "word/document.xml"];
    for (const part of requiredParts) {
      expect(zip.file(part)).toBeTruthy();
    }

    for (const [name, entry] of Object.entries(zip.files)) {
      if (entry.dir || !(name.endsWith(".xml") || name.endsWith(".rels"))) continue;
      const text = await entry.async("text");
      const validation = XMLValidator.validate(text);
      expect.soft(validation, `"${name}" should be well-formed XML`).toBe(true);
      expect((text.match(/<\?xml/g) ?? []).length, `"${name}" should have exactly one XML declaration`).toBe(1);
    }
  });

  it("review-grid scenario borders every cell (60 writable cells were border-free in the source); standard artifacts keep only the source's own spacer-column borders", async () => {
    // The real Avery 5155 source template's 3 spacer/gutter columns
    // already carry their own w:tcBorders (45 = 3 spacer cols x 15 rows) -
    // that's pre-existing source geometry standard artifacts must preserve
    // untouched (cloned once per sheet), not something this renderer adds.
    // reviewOutlines additionally borders the 60 writable cells that have
    // no border in the source, so every cell (105 = 15 rows x 7 cols) ends
    // up outlined for review.
    tmpDir = mkdtempSync(path.join(tmpdir(), "avery-5155-sample-"));
    const sourceXml = await (await JSZip.loadAsync(templateBuffer)).file("word/document.xml")?.async("text");
    if (!sourceXml) throw new Error("word/document.xml missing from the source fixture");
    const sourceBordersPerTable = (sourceXml.match(/<w:tcBorders>/g) ?? []).length;
    expect(sourceBordersPerTable).toBe(45); // 3 spacer columns x 15 rows

    const reviewScenario = SAMPLE_SCENARIOS.find((s) => s.fileBaseName === "sample-avery-5155-review-grid");
    if (!reviewScenario) throw new Error("Expected a review-grid scenario in SAMPLE_SCENARIOS.");
    expect(reviewScenario.reviewOutlines).toBe(true);

    const reviewArtifact = await generateSampleArtifact(templateBuffer, reviewScenario, tmpDir);
    const reviewZip = await JSZip.loadAsync(readFileSync(reviewArtifact.docxPath));
    const reviewXml = await reviewZip.file("word/document.xml")?.async("text");
    if (!reviewXml) throw new Error("word/document.xml missing from the review-grid artifact");
    expect(XMLValidator.validate(reviewXml)).toBe(true);
    // 2 sheets x 15 rows x 7 cells (writable + spacer alike) all bordered.
    expect((reviewXml.match(/<w:tcBorders>/g) ?? []).length).toBe(2 * 15 * 7);

    for (const [fileBaseName, expectedSheets] of [
      ["sample-avery-5155-1-product", 1],
      ["sample-avery-5155-8-products", 2],
    ] as const) {
      const standardScenario = SAMPLE_SCENARIOS.find((s) => s.fileBaseName === fileBaseName);
      if (!standardScenario) throw new Error(`Expected scenario "${fileBaseName}" in SAMPLE_SCENARIOS.`);
      expect(standardScenario.reviewOutlines).toBeFalsy();

      const artifact = await generateSampleArtifact(templateBuffer, standardScenario, tmpDir);
      const zip = await JSZip.loadAsync(readFileSync(artifact.docxPath));
      const xml = await zip.file("word/document.xml")?.async("text");
      if (!xml) throw new Error(`word/document.xml missing from ${fileBaseName}`);
      expect(artifact.sheetCount).toBe(expectedSheets);
      expect((xml.match(/<w:tcBorders>/g) ?? []).length).toBe(sourceBordersPerTable * expectedSheets);
    }
  });

  it("never writes real product-deck content - only the controlled sample data", () => {
    for (const scenario of SAMPLE_SCENARIOS) {
      for (const product of scenario.products) {
        expect(product.sku).toMatch(/^SAMPLE-SKU-/);
        expect(product.description).toMatch(/^Sample /);
      }
    }
  });

  it("sampleAvery5155Template() matches the real Avery 5155 geometry (4 columns x 15 rows = 60 labels/sheet)", () => {
    const template = sampleAvery5155Template();
    expect(template.id).toBe("avery-5155");
    expect(template.renderingMode).toBe("FIXED_GRID");
    expect(template.columns).toBe(4);
    expect(template.rows).toBe(15);
    expect(template.labelsPerSheet).toBe(60);
    expect(template.templateStorageKey).toBe("avery-5155/original.docx");
  });
});
