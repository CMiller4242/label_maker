import { z } from "zod";
import { LABEL_RUN_STATUSES } from "../constants/index.js";

export const labelRunStatusSchema = z.enum(LABEL_RUN_STATUSES);
export type LabelRunStatus = z.infer<typeof labelRunStatusSchema>;

/** A single occupied slot on a printed sheet, bound to one product copy. */
export const placementSchema = z.object({
  sheetNumber: z.number().int().positive(),
  slotIndex: z.number().int().nonnegative(),
  row: z.number().int().nonnegative(),
  column: z.number().int().nonnegative(),
  productId: z.string().uuid(),
  sku: z.string().nullable(),
  description: z.string().nullable(),
  priceCents: z.number().int().nonnegative().nullable(),
  copyIndex: z.number().int().nonnegative(),
});
export type Placement = z.infer<typeof placementSchema>;

export const placementPlanSchema = z.object({
  labelTemplateId: z.string(),
  copiesPerProduct: z.number().int().positive(),
  totalSheets: z.number().int().nonnegative(),
  totalPlacements: z.number().int().nonnegative(),
  placements: z.array(placementSchema),
});
export type PlacementPlan = z.infer<typeof placementPlanSchema>;

/** Request body for POST /label-runs */
export const createLabelRunRequestSchema = z.object({
  sourceDocumentId: z.string().uuid(),
  labelTemplateId: z.string().min(1),
  copiesPerProduct: z.number().int().positive(),
  /** Also generate the optional debug JSON artifact alongside the primary DOCX. Defaults to false. */
  includeDebugArtifact: z.boolean().default(false),
});
export type CreateLabelRunRequest = z.infer<typeof createLabelRunRequestSchema>;

export const labelRunSchema = z.object({
  id: z.string().uuid(),
  sourceDocumentId: z.string().uuid(),
  labelTemplateId: z.string(),
  copiesPerProduct: z.number().int().positive(),
  status: labelRunStatusSchema,
  placementPlanJson: placementPlanSchema,
  generatedArtifactStorageKey: z.string().nullable(),
  generatedArtifactSha256: z.string().nullable(),
  templateVersion: z.string().nullable(),
  templateSha256: z.string().nullable(),
  sheetCount: z.number().int().nonnegative().nullable(),
  filledSlotCount: z.number().int().nonnegative().nullable(),
  emptySlotCount: z.number().int().nonnegative().nullable(),
  metadataJson: z.unknown().nullable(),
  createdAt: z.coerce.date(),
  updatedAt: z.coerce.date(),
});
export type LabelRun = z.infer<typeof labelRunSchema>;

const generatedArtifactInfoSchema = z.object({
  storageKey: z.string(),
  byteSize: z.number().int().nonnegative(),
  sha256: z.string(),
  mimeType: z.string(),
});

/**
 * Response for POST /label-runs. `artifact` is the primary generated
 * output (a real DOCX for a validated FIXED_GRID template); `debugArtifact`
 * is an optional secondary development-time JSON artifact, only present
 * when explicitly requested.
 */
export const createLabelRunResponseSchema = z.object({
  labelRun: labelRunSchema,
  artifact: generatedArtifactInfoSchema,
  debugArtifact: generatedArtifactInfoSchema.omit({ mimeType: true }).optional(),
});
export type CreateLabelRunResponse = z.infer<typeof createLabelRunResponseSchema>;
