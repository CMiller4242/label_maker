#!/usr/bin/env -S pnpm exec tsx
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { createHash } from "node:crypto";
import {
  AVERY_5155_GEOMETRY,
  buildPlacements,
  validatePlacements,
  type PlacementInputProduct,
} from "@label-maker/label-layout";
import { renderFixedGridDocx } from "@label-maker/docx-renderer";
import type { LabelTemplate } from "@label-maker/shared";

/**
 * Generates real, persistent sample Avery 5155 DOCX artifacts on disk (not
 * temporary test files) so a developer can open them directly in Microsoft
 * Word. Uses the exact same production code paths as a real label run
 * (`@label-maker/label-layout`'s `buildPlacements()`/`validatePlacements()`
 * and `@label-maker/docx-renderer`'s `renderFixedGridDocx()`) against the
 * real committed `fixtures/label-templates/avery-5155/original.docx`
 * source template - nothing here reimplements or alters rendering/placement
 * behavior. See tests/scripts/generate-sample-avery-5155.test.ts for
 * automated coverage of `generateSampleArtifact()` (written to a temporary
 * directory, cleaned up after itself - it never touches storage/artifacts).
 *
 * Usage:
 *   pnpm docx:sample-5155
 */

/**
 * Matches the provisional label text style seeded for the real "avery-5155"
 * LabelTemplate preset (packages/database/prisma/seed.ts's
 * `provisionalLabelTextStyle`), duplicated here so this script has no
 * Prisma/database dependency and can run standalone.
 */
export const SAMPLE_LABEL_TEXT_STYLE = {
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
} as const;

/** A LabelTemplate row shaped exactly like the real seeded "avery-5155" preset, without a database round trip. */
export function sampleAvery5155Template(): LabelTemplate {
  return {
    id: AVERY_5155_GEOMETRY.id,
    displayName: "Avery 5155",
    renderingMode: "FIXED_GRID",
    columns: AVERY_5155_GEOMETRY.columns,
    rows: AVERY_5155_GEOMETRY.rows,
    labelsPerSheet: AVERY_5155_GEOMETRY.labelsPerSheet,
    templateStorageKey: "avery-5155/original.docx",
    templateVersion: "0.1.0-sample-script",
    sourceTemplateSha256: null,
    configJson: { labelTextStyle: SAMPLE_LABEL_TEXT_STYLE },
    isPreset: true,
    createdAt: new Date(),
    updatedAt: new Date(),
  };
}

interface SampleProductSpec {
  id: string;
  sku: string;
  description: string;
  priceCents: number;
}

interface SampleScenario {
  fileBaseName: string;
  copiesPerProduct: number;
  products: SampleProductSpec[];
}

/** Controlled, non-sensitive sample data only - never real customer/product-deck content. */
export const SAMPLE_SCENARIOS: SampleScenario[] = [
  {
    fileBaseName: "sample-avery-5155-1-product",
    copiesPerProduct: 8,
    products: [
      { id: "sample-1", sku: "SAMPLE-SKU-100", description: "Sample Widget - Small", priceCents: 999 },
    ],
  },
  {
    fileBaseName: "sample-avery-5155-8-products",
    copiesPerProduct: 8,
    products: Array.from({ length: 8 }, (_, i) => ({
      id: `sample-${i + 1}`,
      sku: `SAMPLE-SKU-${100 + i}`,
      description: `Sample Widget ${String.fromCharCode(65 + i)}`,
      priceCents: 250 * (i + 1),
    })),
  },
];

export interface GeneratedSampleArtifact {
  fileBaseName: string;
  docxPath: string;
  metadataPath: string;
  byteSize: number;
  sha256: string;
  sheetCount: number;
  filledSlotCount: number;
  blankSlotCount: number;
}

/**
 * Renders one sample scenario against the real Avery 5155 source template
 * (via the production `renderFixedGridDocx()`) and writes it - deterministic
 * filename, overwritten on every run - plus a matching metadata JSON
 * sidecar into `outputDir`. Pure with respect to its inputs: the caller
 * decides `outputDir`, so tests can point it at a temporary directory
 * instead of the real `storage/artifacts`.
 */
export async function generateSampleArtifact(
  templateBuffer: Buffer,
  scenario: SampleScenario,
  outputDir: string,
): Promise<GeneratedSampleArtifact> {
  const placementInputs: PlacementInputProduct[] = scenario.products.map((p) => ({
    id: p.id,
    sku: p.sku,
    description: p.description,
    priceCents: p.priceCents,
    include: true,
  }));

  const plan = buildPlacements(placementInputs, AVERY_5155_GEOMETRY, scenario.copiesPerProduct);
  validatePlacements(plan.placements, AVERY_5155_GEOMETRY);

  const result = await renderFixedGridDocx(templateBuffer, sampleAvery5155Template(), plan);

  mkdirSync(outputDir, { recursive: true });
  const docxPath = path.join(outputDir, `${scenario.fileBaseName}.docx`);
  writeFileSync(docxPath, result.buffer);

  const sha256 = createHash("sha256").update(result.buffer).digest("hex");
  const byteSize = result.buffer.byteLength;
  const sheetCount = plan.totalSheets;
  const filledSlotCount = plan.totalPlacements;
  const blankSlotCount = sheetCount * AVERY_5155_GEOMETRY.labelsPerSheet - filledSlotCount;

  const metadataPath = path.join(outputDir, `${scenario.fileBaseName}.metadata.json`);
  writeFileSync(
    metadataPath,
    JSON.stringify(
      {
        generatedAt: new Date().toISOString(),
        sourceTemplate: "fixtures/label-templates/avery-5155/original.docx",
        scenario: scenario.fileBaseName,
        copiesPerProduct: scenario.copiesPerProduct,
        products: scenario.products,
        docxFile: path.basename(docxPath),
        byteSize,
        sha256,
        sheetCount,
        filledSlotCount,
        blankSlotCount,
      },
      null,
      2,
    ),
  );

  return { fileBaseName: scenario.fileBaseName, docxPath, metadataPath, byteSize, sha256, sheetCount, filledSlotCount, blankSlotCount };
}

/**
 * Best-effort UNC path a Windows host can open directly against a WSL
 * filesystem path (e.g. `\\wsl$\Ubuntu\home\user\...`). Only meaningful
 * under WSL (detected via WSL_DISTRO_NAME); on any other platform this is
 * intentionally identical to the absolute path, since the absolute path is
 * already Windows-native (native Windows/Git Bash) or not Windows-relevant
 * (macOS/Linux).
 */
function toWindowsAccessiblePath(absolutePath: string): string {
  const distro = process.env.WSL_DISTRO_NAME;
  if (!distro) return absolutePath;
  return `\\\\wsl$\\${distro}${absolutePath.replace(/\//g, "\\")}`;
}

async function main(): Promise<void> {
  const __dirname = path.dirname(fileURLToPath(import.meta.url));
  const repoRoot = path.join(__dirname, "..");
  const templatePath = path.join(repoRoot, "fixtures", "label-templates", "avery-5155", "original.docx");
  const outputDir = path.join(repoRoot, "storage", "artifacts");

  let templateBuffer: Buffer;
  try {
    templateBuffer = readFileSync(templatePath);
  } catch (error) {
    console.error(
      `Failed to read the Avery 5155 source template at "${templatePath}": ` +
        `${error instanceof Error ? error.message : String(error)}\n` +
        "Run this command from the repository root (or check the fixture hasn't been moved/deleted).",
    );
    process.exitCode = 1;
    return;
  }

  const generated: GeneratedSampleArtifact[] = [];
  for (const scenario of SAMPLE_SCENARIOS) {
    try {
      generated.push(await generateSampleArtifact(templateBuffer, scenario, outputDir));
    } catch (error) {
      console.error(
        `Failed to generate sample artifact "${scenario.fileBaseName}.docx": ` +
          `${error instanceof Error ? `${error.name}: ${error.message}` : String(error)}`,
      );
      process.exitCode = 1;
      return;
    }
  }

  console.log(`Generated ${generated.length} real Avery 5155 sample DOCX artifact(s):\n`);
  for (const artifact of generated) {
    console.log(`[${artifact.fileBaseName}]`);
    console.log(`  Absolute path:            ${artifact.docxPath}`);
    console.log(`  Windows-accessible path:  ${toWindowsAccessiblePath(artifact.docxPath)}`);
    console.log(`  Metadata JSON:            ${artifact.metadataPath}`);
    console.log(`  Byte size:                ${artifact.byteSize}`);
    console.log(`  SHA-256:                  ${artifact.sha256}`);
    console.log(`  Sheets:                   ${artifact.sheetCount}`);
    console.log(`  Filled slots:             ${artifact.filledSlotCount}`);
    console.log(`  Blank slots:              ${artifact.blankSlotCount}`);
    console.log("");
  }
  console.log(
    "These files live under storage/artifacts/ (git-ignored - see .gitignore) and are " +
      "overwritten deterministically every time this command runs.",
  );
}

// Only run the CLI when this file is executed directly (`tsx
// scripts/generate-sample-avery-5155.ts` / `pnpm docx:sample-5155`) - never
// as a side effect of another module importing generateSampleArtifact()
// etc. for testing. Without this guard, importing this file for its
// exports would also regenerate the real storage/artifacts/ files, which
// tests must never do.
const isMainModule = process.argv[1] !== undefined && import.meta.url === pathToFileURL(process.argv[1]).href;
if (isMainModule) {
  main().catch((error: unknown) => {
    console.error("Sample DOCX generation failed:", error);
    process.exitCode = 1;
  });
}
