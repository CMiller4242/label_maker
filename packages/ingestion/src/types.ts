import type { SourceDocumentType } from "@label-maker/shared";

export type ExtractionCandidateField = "SKU" | "DESCRIPTION" | "PRICE" | "INCLUDE_STATUS";

export interface CandidateBoundingBox {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface IngestionCandidate {
  field: ExtractionCandidateField;
  rawValue: string;
  normalizedValue?: string;
  confidence: number;
  boundingBox?: CandidateBoundingBox;
  ruleName: string;
}

export interface IngestionProduct {
  sourcePageNumber?: number;
  sourceRowNumber?: number;
  sku: string | null;
  description: string | null;
  priceCents: number | null;
  include: boolean;
  confidence: number;
  candidates: IngestionCandidate[];
}

export interface IngestionResult {
  documentType: SourceDocumentType;
  pageCount?: number;
  sheetNames?: string[];
  products: IngestionProduct[];
  warnings: string[];
  needsReview: boolean;
}

/** Raw bytes plus the filename/mime type the API observed on upload. */
export interface RawSourceFile {
  buffer: Buffer;
  originalFilename: string;
  mimeType: string;
}
