import type { LabelTemplateGeometry } from "./types.js";

/**
 * Geometry for known presets, mirroring the rows/columns/labelsPerSheet
 * seeded in packages/database/prisma/seed.ts. This lets the placement
 * engine and its tests run without a database round trip; the API/worker
 * still read the authoritative LabelTemplate row (including configJson
 * geometry for eventual DOCX rendering) from Postgres.
 */
export const AVERY_5155_GEOMETRY: LabelTemplateGeometry = {
  id: "avery-5155",
  columns: 4,
  rows: 15,
  labelsPerSheet: 60,
};

export const KNOWN_TEMPLATE_GEOMETRIES: Record<string, LabelTemplateGeometry> = {
  "avery-5155": AVERY_5155_GEOMETRY,
};
