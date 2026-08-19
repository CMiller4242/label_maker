import { PrismaClient } from "@prisma/client";

const prisma = new PrismaClient();

/**
 * Avery 5155 preset.
 *
 * IMPORTANT: the geometry values in `configJson` below are structural
 * placeholders only (they establish the shape future code will read), NOT
 * verified measurements. Avery 5155 is a 4-column x 15-row, 60-label sheet,
 * so columns/rows/labelsPerSheet are correct, but every dimension inside
 * `configJson.geometry` MUST be replaced with values measured from a
 * verified physical Avery 5155 Word template before this preset is used to
 * generate anything that gets printed.
 */
const avery5155ConfigJson = {
  templateVersion: "0.0.0-unverified",
  geometry: {
    // All linear measurements are in points (1/72 inch). Placeholder only.
    page: {
      widthPt: 612, // TODO: confirm against verified US Letter template
      heightPt: 792, // TODO: confirm against verified US Letter template
    },
    margins: {
      topPt: 0, // TODO: replace with measured value
      bottomPt: 0, // TODO: replace with measured value
      leftPt: 0, // TODO: replace with measured value
      rightPt: 0, // TODO: replace with measured value
    },
    label: {
      widthPt: 0, // TODO: replace with measured value
      heightPt: 0, // TODO: replace with measured value
    },
    pitch: {
      horizontalPt: 0, // TODO: distance between label origins, left to right
      verticalPt: 0, // TODO: distance between label origins, top to bottom
    },
    safeInsets: {
      topPt: 0, // TODO: printable-safe inset from label edge
      bottomPt: 0,
      leftPt: 0,
      rightPt: 0,
    },
  },
  font: {
    // TODO: confirm final font family/size once a physical template is verified.
    family: "Calibri",
    sizePt: 10,
    lineHeightPt: 12,
  },
  fixedTableLayout: {
    // DOCX rendering requirement: Word tables must use fixed layout (not
    // autofit) with explicit row heights so label positions are deterministic.
    tableLayoutMode: "fixed",
    exactRowHeights: true,
    autoFit: false,
  },
} as const;

async function main(): Promise<void> {
  await prisma.labelTemplate.upsert({
    where: { id: "avery-5155" },
    update: {},
    create: {
      id: "avery-5155",
      displayName: "Avery 5155",
      renderingMode: "FIXED_GRID",
      columns: 4,
      rows: 15,
      labelsPerSheet: 60,
      templateStorageKey: null,
      configJson: avery5155ConfigJson,
      isPreset: true,
    },
  });

  console.log("Seed complete: avery-5155 label template preset upserted.");
}

main()
  .catch((error: unknown) => {
    console.error("Seed failed:", error);
    process.exitCode = 1;
  })
  .finally(() => {
    void prisma.$disconnect();
  });
