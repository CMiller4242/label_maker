import { z } from "zod";
import { SOURCE_DOCUMENT_TYPES } from "../constants/index.js";

/** Supported uploaded source document types. */
export const sourceDocumentTypeSchema = z.enum(SOURCE_DOCUMENT_TYPES);
export type SourceDocumentType = z.infer<typeof sourceDocumentTypeSchema>;

export const sourceDocumentSchema = z.object({
  id: z.string().uuid(),
  originalFilename: z.string(),
  mimeType: z.string(),
  extension: sourceDocumentTypeSchema,
  storageKey: z.string(),
  byteSize: z.number().int().nonnegative(),
  sha256: z.string().length(64),
  pageCount: z.number().int().nonnegative().nullable(),
  createdAt: z.coerce.date(),
});
export type SourceDocument = z.infer<typeof sourceDocumentSchema>;

/** Response returned by POST /uploads. */
export const uploadResponseSchema = z.object({
  documentId: z.string().uuid(),
  jobId: z.string().uuid(),
});
export type UploadResponse = z.infer<typeof uploadResponseSchema>;
