import { mkdtempSync, readFileSync, rmSync, statSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { createHash } from "node:crypto";
import { afterEach, describe, expect, it } from "vitest";
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
