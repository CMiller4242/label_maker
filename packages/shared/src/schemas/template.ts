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
