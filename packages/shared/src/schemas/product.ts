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
