import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { PrismaClient } from "@prisma/client";
import { inspectDocxTemplate } from "@label-maker/docx-renderer";

const prisma = new PrismaClient();

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const templatesRoot = path.join(__dirname, "..", "..", "..", "fixtures", "label-templates");

const AVERY_5155_TEMPLATE_STORAGE_KEY = "avery-5155/original.docx";
const AVERY_22802_TEMPLATE_STORAGE_KEY = "avery-22802/original.docx";

const TEMPLATE_VERSION = "0.1.0-inspected";

/**
 * Explicit, configurable style for the 3 label lines (SKU / description /
 * price). Provisional - not yet confirmed against a physically printed
 * Avery 5155 sheet (see fixtures/label-templates/avery-5155/PRINT-TEST.md) -
 * but deliberately explicit and centralized rather than scattered magic
 * numbers in renderer code. sz values are in half-points (OOXML w:sz).
 */
const provisionalLabelTextStyle = {
  fontFamily: "Calibri",
  skuFontSizeHalfPoints: 18, // 9pt
  descriptionFontSizeHalfPoints: 16, // 8pt
  priceFontSizeHalfPoints: 18, // 9pt
  bold: false,
  colorHex: "000000",
  horizontalAlignment: "center",
  verticalAlignment: "center",
  lineSpacing: { lineRule: "auto", line: 240 }, // single line spacing
  paragraphSpacingBeforeTwips: 0,
  paragraphSpacingAfterTwips: 0,
} as const;

async function main(): Promise<void> {
  // --- Avery 5155 (FIXED_GRID) --------------------------------------------
  const avery5155Buffer = readFileSync(path.join(templatesRoot, "avery-5155", "original.docx"));
  const avery5155Sha256 = createSha256(avery5155Buffer);
  const avery5155Inspection = await inspectDocxTemplate({
    buffer: avery5155Buffer,
    filePath: "fixtures/label-templates/avery-5155/original.docx",
  });

  const avery5155Table = avery5155Inspection.tables[0];

  const avery5155ConfigJson = {
    templateVersion: TEMPLATE_VERSION,
    sourceInspection: {
      classification: avery5155Inspection.classification,
      classificationConfidence: avery5155Inspection.classificationConfidence,
      tableCount: avery5155Inspection.tableCount,
      warnings: avery5155Inspection.warnings,
    },
    // fixtures/label-templates/avery-5155/original.docx (re-inspected here at
    // seed time) IS now a valid Word document with one table, 15 rows,
    // fixed layout, and exact row heights - see sourceInspection above.
    // HOWEVER: its <w:tblGrid> has 7 columns
    // (widths in twips: [2520, 432, 2520, 432, 2520, 432, 2520]), not the
    // uniform 4 columns per row that validateFixedGridTemplate()/
    // renderFixedGridDocx() currently require. This is a real, common Avery
    // Word-template pattern: 4 real label columns interleaved with 3
    // narrow, vertically-merged spacer/gutter columns (confirmed via
    // <w:vMerge> on the narrow columns) used to get the correct horizontal
    // pitch between label columns. inspectDocxTemplate() therefore reports
    // classification "AMBIGUOUS" (105 raw cells, not the expected 60), and
    // renderFixedGridDocx() will fail InvalidFixedGridTemplateError against
    // this exact file until the validator/renderer are extended to
    // recognize interleaved spacer columns - that extension is out of scope
    // for this pass (verification/fixture-activation only, no renderer
    // changes). Page size and margins below ARE real measurements from this
    // file; label/pitch/safeInsets are left as TODO rather than guessed,
    // since correctly deriving them requires that renderer-side
    // interleaved-column support which does not exist yet.
    geometry: {
      page: {
        widthPt: (avery5155Inspection.pageGeometry.widthTwips ?? 0) / 20,
        heightPt: (avery5155Inspection.pageGeometry.heightTwips ?? 0) / 20,
      },
      margins: {
        topPt: (avery5155Inspection.pageGeometry.marginTopTwips ?? 0) / 20,
        bottomPt: (avery5155Inspection.pageGeometry.marginBottomTwips ?? 0) / 20,
        leftPt: (avery5155Inspection.pageGeometry.marginLeftTwips ?? 0) / 20,
        rightPt: (avery5155Inspection.pageGeometry.marginRightTwips ?? 0) / 20,
      },
      // TODO: requires interleaved-spacer-column support in the renderer/
      // validator to derive correctly (see note above) - do not infer here.
      label: { widthPt: 0, heightPt: 0 },
      pitch: { horizontalPt: 0, verticalPt: 0 },
      safeInsets: { topPt: 0, bottomPt: 0, leftPt: 0, rightPt: 0 },
      // Raw measurements actually available from this file's <w:tblGrid> /
      // row height, for whenever interleaved-column support is added.
      rawTableGrid: {
        gridColumnWidthsTwips: avery5155Table?.gridColumnWidthsTwips ?? [],
        rowHeightTwips: avery5155Table?.rows[0]?.heightTwips ?? null,
        rowHeightRule: avery5155Table?.rows[0]?.heightRule ?? null,
      },
    },
    fixedTableLayout: {
      tableLayoutMode: "fixed",
      exactRowHeights: true,
      autoFit: false,
    },
    labelTextStyle: provisionalLabelTextStyle,
  };

  await prisma.labelTemplate.upsert({
    where: { id: "avery-5155" },
    update: {
      templateStorageKey: AVERY_5155_TEMPLATE_STORAGE_KEY,
      templateVersion: TEMPLATE_VERSION,
      sourceTemplateSha256: avery5155Sha256,
      configJson: avery5155ConfigJson,
    },
    create: {
      id: "avery-5155",
      displayName: "Avery 5155",
      renderingMode: "FIXED_GRID",
      columns: 4,
      rows: 15,
      labelsPerSheet: 60,
      templateStorageKey: AVERY_5155_TEMPLATE_STORAGE_KEY,
      templateVersion: TEMPLATE_VERSION,
      sourceTemplateSha256: avery5155Sha256,
      configJson: avery5155ConfigJson,
      isPreset: true,
    },
  });

  if (avery5155Inspection.classification !== "AVERY_5155_LIKE_FIXED_GRID") {
    console.warn(
      `WARNING: fixtures/label-templates/avery-5155/original.docx inspected as ` +
        `"${avery5155Inspection.classification}", not "AVERY_5155_LIKE_FIXED_GRID". ` +
        `renderFixedGridDocx() will fail InvalidFixedGridTemplateError against this file. See ` +
        `packages/database/prisma/seed.ts's avery5155ConfigJson comment and PRINT-TEST.md.`,
    );
  }

  // --- Avery 22802 (FLOATING) ----------------------------------------------
  const avery22802Buffer = readFileSync(path.join(templatesRoot, "avery-22802", "original.docx"));
  const avery22802Sha256 = createSha256(avery22802Buffer);
  const avery22802Inspection = await inspectDocxTemplate({
    buffer: avery22802Buffer,
    filePath: "fixtures/label-templates/avery-22802/original.docx",
  });

  const firstTable = avery22802Inspection.tables[0];
  const avery22802ConfigJson = {
    templateVersion: TEMPLATE_VERSION,
    renderingSupport: "NOT_IMPLEMENTED",
    sourceInspection: {
      classification: avery22802Inspection.classification,
      classificationConfidence: avery22802Inspection.classificationConfidence,
      tableCount: avery22802Inspection.tableCount,
      warnings: avery22802Inspection.warnings,
    },
    inspectedLayout: {
      pageWidthTwips: avery22802Inspection.pageGeometry.widthTwips,
      pageHeightTwips: avery22802Inspection.pageGeometry.heightTwips,
      marginTopTwips: avery22802Inspection.pageGeometry.marginTopTwips,
      marginBottomTwips: avery22802Inspection.pageGeometry.marginBottomTwips,
      marginLeftTwips: avery22802Inspection.pageGeometry.marginLeftTwips,
      marginRightTwips: avery22802Inspection.pageGeometry.marginRightTwips,
      tableCount: avery22802Inspection.tableCount,
      labelWidthTwips: firstTable?.gridColumnWidthsTwips[0] ?? null,
      labelHeightTwips: firstTable?.rows[0]?.heightTwips ?? null,
      floating: firstTable?.positioning.isFloating ?? null,
    },
  };

  const labelsPerSheet = avery22802Inspection.tableCount;

  await prisma.labelTemplate.upsert({
    where: { id: "avery-22802" },
    update: {
      templateStorageKey: AVERY_22802_TEMPLATE_STORAGE_KEY,
      templateVersion: TEMPLATE_VERSION,
      sourceTemplateSha256: avery22802Sha256,
      configJson: avery22802ConfigJson,
      labelsPerSheet,
    },
    create: {
      id: "avery-22802",
      displayName: "Avery 22802",
      renderingMode: "FLOATING",
      // FLOATING templates aren't a uniform columns x rows grid; these are
      // structural placeholders (not used by any grid math) until floating
      // rendering is implemented.
      columns: 1,
      rows: labelsPerSheet,
      labelsPerSheet,
      templateStorageKey: AVERY_22802_TEMPLATE_STORAGE_KEY,
      templateVersion: TEMPLATE_VERSION,
      sourceTemplateSha256: avery22802Sha256,
      configJson: avery22802ConfigJson,
      isPreset: true,
    },
  });

  console.log("Seed complete: avery-5155 and avery-22802 label template presets upserted.");
}

function createSha256(buffer: Buffer): string {
  return createHash("sha256").update(buffer).digest("hex");
}

main()
  .catch((error: unknown) => {
    console.error("Seed failed:", error);
    process.exitCode = 1;
  })
  .finally(() => {
    void prisma.$disconnect();
  });
