import { z } from "zod";
import { EXTRACTION_FIELDS, PRODUCT_STATUSES } from "../constants/index.js";

export const extractionFieldSchema = z.enum(EXTRACTION_FIELDS);
export type ExtractionField = z.infer<typeof extractionFieldSchema>;

export const productStatusSchema = z.enum(PRODUCT_STATUSES);
export type ProductStatus = z.infer<typeof productStatusSchema>;

/**
 * A single normalized product row as submitted/edited by a user, prior to
 * being persisted. priceCents is always an integer (cents), never a float.
 */
export const productInputSchema = z.object({
  sku: z.string().min(1),
  description: z.string().min(1),
  priceCents: z.number().int().nonnegative(),
  include: z.boolean().default(true),
});
export type ProductInput = z.infer<typeof productInputSchema>;

export const boundingBoxSchema = z.object({
  x: z.number(),
  y: z.number(),
  width: z.number(),
  height: z.number(),
});
export type BoundingBox = z.infer<typeof boundingBoxSchema>;

/** Raw extraction evidence backing a Product field. Kept separate from approved Product data. */
export const extractionCandidateSchema = z.object({
  id: z.string().uuid(),
  sourcePageId: z.string().uuid().nullable(),
  rowNumber: z.number().int().nonnegative().nullable(),
  field: extractionFieldSchema,
  rawValue: z.string(),
  normalizedValue: z.string().nullable(),
  confidence: z.number().min(0).max(1),
  boundingBoxJson: boundingBoxSchema.nullable(),
  ruleName: z.string().nullable(),
  createdAt: z.coerce.date(),
});
export type ExtractionCandidate = z.infer<typeof extractionCandidateSchema>;

export const productSchema = z.object({
  id: z.string().uuid(),
  sourceDocumentId: z.string().uuid(),
  sourcePageId: z.string().uuid().nullable(),
  sourceRowNumber: z.number().int().nonnegative().nullable(),
  sku: z.string().nullable(),
  description: z.string().nullable(),
  priceCents: z.number().int().nonnegative().nullable(),
  include: z.boolean(),
  status: productStatusSchema,
  confidence: z.number().min(0).max(1),
  extractionNotesJson: z.record(z.string(), z.unknown()).nullable(),
  createdAt: z.coerce.date(),
  updatedAt: z.coerce.date(),
});
export type Product = z.infer<typeof productSchema>;

/** Response for GET /documents/:documentId/products */
export const productListResponseSchema = z.object({
  documentId: z.string().uuid(),
  products: z.array(productSchema),
});
export type ProductListResponse = z.infer<typeof productListResponseSchema>;

/**
 * Request body for PATCH /products/:productId. Every field is optional -
 * only the fields present are updated - but at least one must be given.
 * Editing sku/description/priceCents on a NEEDS_REVIEW row does not by
 * itself make it eligible for a label run; the caller must also send
 * status: "APPROVED" (a human has now reviewed it) for it to be included,
 * matching createLabelRun's existing include=true AND status IN
 * (APPROVED, AUTO_ACCEPTED) filter - this endpoint does not change that
 * filter, only lets a client legitimately reach the states it already
 * understands.
 */
export const updateProductRequestSchema = z
  .object({
    sku: z.string().min(1).nullable(),
    description: z.string().min(1).nullable(),
    priceCents: z.number().int().nonnegative().nullable(),
    include: z.boolean(),
    status: productStatusSchema,
  })
  .partial()
  .refine((body) => Object.keys(body).length > 0, {
    message: "At least one field must be provided.",
  });
export type UpdateProductRequest = z.infer<typeof updateProductRequestSchema>;

/**
 * Request body for POST /documents/:documentId/products - adds a single
 * manually-entered product row (no source page/row, no extraction
 * candidate behind it). Defaults to status "APPROVED" and confidence 1,
 * since a human is directly authoring the row rather than an extractor
 * proposing one for review.
 */
export const createProductRequestSchema = z.object({
  sku: z.string().min(1),
  description: z.string().min(1),
  priceCents: z.number().int().nonnegative(),
  include: z.boolean().default(true),
});
export type CreateProductRequest = z.infer<typeof createProductRequestSchema>;
