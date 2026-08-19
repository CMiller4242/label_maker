import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";
import type { TextItem, TextMarkedContent } from "pdfjs-dist/types/src/display/api.js";
import { createRequire } from "node:module";
import path from "node:path";
import { CONFIDENCE, PDF_MIN_TEXT_DENSITY_CHARS_PER_PAGE } from "./confidence.js";
import { normalizePriceToCents } from "./normalize-price.js";
import { DEFAULT_SKU_REGEX, findSkuLikeTokens } from "./sku.js";
import type { CandidateBoundingBox, IngestionCandidate, IngestionResult } from "./types.js";

export class PdfParseError extends Error {
  readonly code = "PDF_PARSE_ERROR";
}

const CURRENCY_REGEX = /\$\s?\d{1,3}(?:,\d{3})*(?:\.\d{2})?/g;
const AS_LOW_AS_REGEX = /as\s+low\s+as/gi;

// Points pdf.js at its bundled standard font metrics so it doesn't warn
// about missing glyph widths for non-embedded standard fonts (Helvetica, etc).
const require = createRequire(import.meta.url);
const STANDARD_FONT_DATA_URL = `${path.join(
  path.dirname(require.resolve("pdfjs-dist/package.json")),
  "standard_fonts",
)}${path.sep}`;

export interface PdfPageResult {
  pageNumber: number;
  viewportWidth: number;
  viewportHeight: number;
  rawText: string;
  /** Non-whitespace characters per page, used as a crude text-density score. */
  textDensityScore: number;
  needsReview: boolean;
  /**
   * Set when this page's native text density is too low to trust. OCR is
   * NOT implemented in this pass (see packages/ocr-provider interface,
   * TODO below) - this only records that OCR would be the next step.
   */
  suggestedFutureExtractionMethod?: "OCR_REQUIRED";
  candidates: IngestionCandidate[];
}

export interface PdfIngestionResult extends IngestionResult {
  pages: PdfPageResult[];
}

function isTextItem(item: TextItem | TextMarkedContent): item is TextItem {
  return "str" in item;
}

function boundingBoxFromItem(item: TextItem, viewportHeight: number): CandidateBoundingBox {
  // pdf.js text transforms are [scaleX, skewX, skewY, scaleY, translateX, translateY]
  // in PDF space (origin bottom-left); flip Y to a top-left-origin box for
  // consumers (e.g. a future review UI overlay) that expect screen space.
  const [, , , , translateX = 0, translateY = 0] = item.transform;
  const x = translateX;
  const y = viewportHeight - translateY - item.height;
  return { x, y, width: item.width, height: item.height };
}

function findCandidatesInLine(
  text: string,
  boundingBox: CandidateBoundingBox | undefined,
): IngestionCandidate[] {
  const candidates: IngestionCandidate[] = [];

  for (const token of findSkuLikeTokens(text, DEFAULT_SKU_REGEX)) {
    candidates.push({
      field: "SKU",
      rawValue: token,
      normalizedValue: token,
      confidence: CONFIDENCE.PDF_GENERIC,
      ruleName: "pdf-sku-like-token-regex",
      ...(boundingBox ? { boundingBox } : {}),
    });
  }

  const currencyMatches = text.match(CURRENCY_REGEX) ?? [];
  for (const rawValue of currencyMatches) {
    const { priceCents, malformed } = normalizePriceToCents(rawValue);
    candidates.push({
      field: "PRICE",
      rawValue,
      confidence: malformed ? CONFIDENCE.LOW : CONFIDENCE.PDF_GENERIC,
      ruleName: "pdf-currency-regex",
      ...(priceCents !== null ? { normalizedValue: String(priceCents) } : {}),
      ...(boundingBox ? { boundingBox } : {}),
    });
  }

  if (AS_LOW_AS_REGEX.test(text)) {
    AS_LOW_AS_REGEX.lastIndex = 0;
    candidates.push({
      field: "PRICE",
      rawValue: text,
      confidence: CONFIDENCE.PDF_GENERIC,
      ruleName: "pdf-as-low-as-label",
      ...(boundingBox ? { boundingBox } : {}),
    });
  }

  return candidates;
}

/**
 * Extracts native (embedded) text from a PDF using pdf.js, page by page.
 * This is programmatic, local extraction only - no OCR, no LLM. Generic
 * candidate discovery (SKU-like tokens, currency values, "AS LOW AS"
 * labels) is applied per line of text; this deliberately does NOT attempt
 * to pair candidates into finished Product records, since doing so
 * correctly requires deck-specific layout rules that are out of scope here.
 *
 * TODO(deck-specific-rules): a future pass should introduce per-deck layout
 * templates (e.g. "SKU appears N points below its price label") to pair
 * candidates into unambiguous product rows. Where that pairing is still
 * ambiguous after applying deck rules, route the page to a human review
 * queue; only as a last resort for handwritten proof markup should an LLM
 * be considered for that specific exception path (never for standard page
 * parsing).
 */
export async function parsePdf(buffer: Buffer): Promise<PdfIngestionResult> {
  let doc;
  try {
    const loadingTask = pdfjsLib.getDocument({
      data: new Uint8Array(buffer),
      useWorkerFetch: false,
      isEvalSupported: false,
      disableFontFace: true,
      standardFontDataUrl: STANDARD_FONT_DATA_URL,
    });
    doc = await loadingTask.promise;
  } catch (error) {
    throw new PdfParseError(
      `Failed to parse PDF: ${error instanceof Error ? error.message : String(error)}`,
    );
  }

  const pages: PdfPageResult[] = [];
  const warnings: string[] = [];

  for (let pageNumber = 1; pageNumber <= doc.numPages; pageNumber++) {
    const page = await doc.getPage(pageNumber);
    const viewport = page.getViewport({ scale: 1 });
    const textContent = await page.getTextContent();

    const items = textContent.items.filter(isTextItem);
    const rawText = items.map((item) => item.str).join(" ");
    const textDensityScore = rawText.replace(/\s/g, "").length;
    const needsReview = textDensityScore < PDF_MIN_TEXT_DENSITY_CHARS_PER_PAGE;

    if (needsReview) {
      warnings.push(
        `Page ${pageNumber}: low native text density (${textDensityScore} chars); flagged NEEDS_REVIEW, future extraction method OCR_REQUIRED.`,
      );
    }

    const candidates: IngestionCandidate[] = [];
    for (const item of items) {
      const boundingBox = boundingBoxFromItem(item, viewport.height);
      candidates.push(...findCandidatesInLine(item.str, boundingBox));
    }

    pages.push({
      pageNumber,
      viewportWidth: viewport.width,
      viewportHeight: viewport.height,
      rawText,
      textDensityScore,
      needsReview,
      ...(needsReview ? { suggestedFutureExtractionMethod: "OCR_REQUIRED" as const } : {}),
      candidates,
    });
  }

  return {
    documentType: "PDF",
    pageCount: doc.numPages,
    // Deliberately empty: pairing candidates into finished Product rows
    // requires deck-specific rules not implemented in this pass (see TODO
    // above). Candidates are still fully captured per-page for review.
    products: [],
    warnings,
    needsReview: true,
    pages,
  };
}
