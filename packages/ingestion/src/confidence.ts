import type { IngestionCandidate } from "./types.js";

/** Confidence levels used by structured (CSV/XLSX) row parsing. */
export const CONFIDENCE = {
  /** Value present, well-formed, unambiguous header mapping. */
  HIGH: 0.95,
  /** Value present but required a fallback rule or minor normalization. */
  MEDIUM: 0.7,
  /** Value missing, malformed, or derived from an ambiguous mapping. */
  LOW: 0.3,
  /** Generic PDF text-scan candidate with no structural confirmation. */
  PDF_GENERIC: 0.4,
} as const;

/**
 * Aggregate confidence for a product row: the minimum of its field
 * candidate confidences (a row is only as trustworthy as its weakest
 * field), or LOW if there are no candidates at all.
 */
export function aggregateRowConfidence(candidates: IngestionCandidate[]): number {
  if (candidates.length === 0) return CONFIDENCE.LOW;
  return Math.min(...candidates.map((c) => c.confidence));
}

/** Minimum text density (non-whitespace chars per page area unit) to trust native PDF extraction. */
export const PDF_MIN_TEXT_DENSITY_CHARS_PER_PAGE = 40;

/** Confidence threshold below which a Product is auto-flagged NEEDS_REVIEW. */
export const NEEDS_REVIEW_CONFIDENCE_THRESHOLD = 0.7;
