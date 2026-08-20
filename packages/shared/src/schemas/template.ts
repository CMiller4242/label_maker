import { z } from "zod";
import { LABEL_TEMPLATE_RENDERING_MODES } from "../constants/index.js";

export const labelTemplateRenderingModeSchema = z.enum(LABEL_TEMPLATE_RENDERING_MODES);
export type LabelTemplateRenderingMode = z.infer<typeof labelTemplateRenderingModeSchema>;

export const labelTemplateSchema = z.object({
  id: z.string(),
  displayName: z.string(),
  renderingMode: labelTemplateRenderingModeSchema,
  columns: z.number().int().positive(),
  rows: z.number().int().positive(),
  labelsPerSheet: z.number().int().positive(),
  templateStorageKey: z.string().nullable(),
  templateVersion: z.string().nullable(),
  sourceTemplateSha256: z.string().nullable(),
  configJson: z.record(z.string(), z.unknown()),
  isPreset: z.boolean(),
  createdAt: z.coerce.date(),
  updatedAt: z.coerce.date(),
});
export type LabelTemplate = z.infer<typeof labelTemplateSchema>;

export const labelTemplateListResponseSchema = z.object({
  templates: z.array(labelTemplateSchema),
});
export type LabelTemplateListResponse = z.infer<typeof labelTemplateListResponseSchema>;

/**
 * Shape of `LabelTemplate.configJson.logicalGrid` for a FIXED_GRID template
 * whose source `.docx` uses the interleaved-spacer-column pattern (e.g.
 * Avery 5155) - see `packages/database/prisma/seed.ts` (which writes this,
 * derived from `inspectDocxTemplate()`, never hand-authored) and
 * `FixedGridPattern` in `packages/docx-renderer/src/types.ts` (which this
 * mirrors structurally, without importing that package's Node-oriented
 * runtime code into browser bundles). A UI rendering a real preview of the
 * physical label grid must read column widths/spacer positions from here,
 * never assume a uniform `columns`-wide grid.
 */
export const fixedGridLogicalGridSchema = z.object({
  patternType: z.enum(["SIMPLE", "INTERLEAVED_SPACER"]),
  logicalColumns: z.number().int().positive(),
  logicalRows: z.number().int().positive(),
  /** Raw `<w:tblGrid>` column indexes (0-based) that hold real, writable label content. */
  writableRawColumnIndexes: z.array(z.number().int().nonnegative()),
  /** Raw column indexes that are spacer/gutter columns - never write label content here. */
  spacerRawColumnIndexes: z.array(z.number().int().nonnegative()),
  /** Every raw column's width in twips, left to right, including spacer columns. */
  rawGridColumnWidthsTwips: z.array(z.number().int().positive()),
  rowHeightTwips: z.number().int().positive().nullable(),
});
export type FixedGridLogicalGrid = z.infer<typeof fixedGridLogicalGridSchema>;

/**
 * The subset of a FIXED_GRID template's `configJson` a UI needs to render
 * an accurate preview and know what's still provisional. Parse with
 * `.safeParse()` - a template with no `logicalGrid` (not yet inspected,
 * or a floating template) is a normal, expected case, not an error.
 */
export const fixedGridTemplateConfigSchema = z.object({
  logicalGrid: fixedGridLogicalGridSchema.nullable().optional(),
  geometry: z
    .object({
      page: z.object({ widthPt: z.number(), heightPt: z.number() }).optional(),
    })
    .partial()
    .optional(),
});
export type FixedGridTemplateConfig = z.infer<typeof fixedGridTemplateConfigSchema>;

/**
 * User-provided mapping of spreadsheet columns to logical product fields.
 * Accepted by the ingestion parser now; the API endpoint for submitting one
 * interactively is a TODO (see apps/api routes).
 */
export const spreadsheetMappingSchema = z.object({
  sheetName: z.string().min(1),
  headerRowNumber: z.number().int().positive(),
  skuColumn: z.string().min(1),
  descriptionColumn: z.string().min(1),
  priceColumn: z.string().min(1),
  includeColumn: z.string().min(1).optional(),
});
export type SpreadsheetMapping = z.infer<typeof spreadsheetMappingSchema>;
