import { parse } from "csv-parse/sync";
import { CONFIDENCE } from "./confidence.js";
import { normalizeHeaders } from "./normalize-headers.js";
import { normalizePriceToCents } from "./normalize-price.js";
import type { IngestionCandidate, IngestionProduct, IngestionResult } from "./types.js";

export class CsvParseError extends Error {
  readonly code = "CSV_PARSE_ERROR";
}

function readCell(row: Record<string, string>, header: string | undefined): string {
  if (!header) return "";
  return (row[header] ?? "").trim();
}

/**
 * Parses a CSV buffer into an IngestionResult. Handles BOM, CRLF/LF, quoted
 * commas, blank lines, and variable header casing/spacing via csv-parse and
 * normalizeHeaders(). Ambiguous header mappings (multiple source columns
 * mapping to the same logical field) are surfaced as warnings and force
 * needsReview rather than being silently resolved.
 */
export function parseCsv(buffer: Buffer): IngestionResult {
  let rows: Record<string, string>[];
  try {
    rows = parse(buffer, {
      bom: true,
      columns: true,
      skip_empty_lines: true,
      trim: true,
      relax_column_count: true,
    }) as Record<string, string>[];
  } catch (error) {
    throw new CsvParseError(
      `Failed to parse CSV: ${error instanceof Error ? error.message : String(error)}`,
    );
  }

  const warnings: string[] = [];

  if (rows.length === 0) {
    return {
      documentType: "CSV",
      products: [],
      warnings: ["CSV contains no data rows."],
      needsReview: true,
    };
  }

  const rawHeaders = Object.keys(rows[0] ?? {});
  const headerMapping = normalizeHeaders(rawHeaders);

  for (const field of headerMapping.ambiguousFields) {
    warnings.push(
      `Multiple columns map to logical field ${field}: ${headerMapping.matches[field].join(", ")}. Refusing to guess; row will need review.`,
    );
  }

  const skuHeader = headerMapping.matches.SKU[0];
  const descriptionHeader = headerMapping.matches.DESCRIPTION[0];
  const priceHeader = headerMapping.matches.PRICE[0];
  const includeHeader = headerMapping.matches.INCLUDE_STATUS[0];

  if (!skuHeader) warnings.push("No column recognized as SKU/item number.");
  if (!descriptionHeader) warnings.push("No column recognized as description.");
  if (!priceHeader) warnings.push("No column recognized as an 'As Low As' price.");

  const hasAmbiguity = headerMapping.ambiguousFields.length > 0;
  let anyRowNeedsReview = hasAmbiguity;

  const products: IngestionProduct[] = rows.map((row, index) => {
    const sourceRowNumber = index + 2; // +1 for 1-indexing, +1 for header row
    const candidates: IngestionCandidate[] = [];

    const rawSku = readCell(row, skuHeader);
    const skuCandidate: IngestionCandidate = {
      field: "SKU",
      rawValue: rawSku,
      confidence: rawSku.length > 0 ? CONFIDENCE.HIGH : CONFIDENCE.LOW,
      ruleName: "csv-header-mapping",
      ...(rawSku.length > 0 ? { normalizedValue: rawSku } : {}),
    };
    candidates.push(skuCandidate);

    const rawDescription = readCell(row, descriptionHeader);
    const descriptionCandidate: IngestionCandidate = {
      field: "DESCRIPTION",
      rawValue: rawDescription,
      confidence: rawDescription.length > 0 ? CONFIDENCE.HIGH : CONFIDENCE.LOW,
      ruleName: "csv-header-mapping",
      ...(rawDescription.length > 0 ? { normalizedValue: rawDescription } : {}),
    };
    candidates.push(descriptionCandidate);

    const rawPrice = readCell(row, priceHeader);
    const priceResult = normalizePriceToCents(rawPrice);
    const priceCandidate: IngestionCandidate = {
      field: "PRICE",
      rawValue: rawPrice,
      confidence:
        rawPrice.length === 0
          ? CONFIDENCE.LOW
          : priceResult.malformed
            ? CONFIDENCE.LOW
            : CONFIDENCE.HIGH,
      ruleName: "csv-price-normalization",
      ...(priceResult.priceCents !== null
        ? { normalizedValue: String(priceResult.priceCents) }
        : {}),
    };
    candidates.push(priceCandidate);

    let include = true;
    if (includeHeader) {
      const rawInclude = readCell(row, includeHeader).toLowerCase();
      if (rawInclude.length > 0) {
        include = !["false", "no", "0", "inactive", "excluded", "n"].includes(rawInclude);
        candidates.push({
          field: "INCLUDE_STATUS",
          rawValue: rawInclude,
          normalizedValue: String(include),
          confidence: CONFIDENCE.HIGH,
          ruleName: "csv-include-status",
        });
      }
    }

    if (priceResult.malformed) {
      warnings.push(`Row ${sourceRowNumber}: malformed price value "${rawPrice}".`);
    }

    const rowHasBlank = rawSku.length === 0 || rawDescription.length === 0 || rawPrice.length === 0;
    const rowNeedsReview = hasAmbiguity || rowHasBlank || priceResult.malformed;
    if (rowNeedsReview) anyRowNeedsReview = true;

    return {
      sourceRowNumber,
      sku: rawSku.length > 0 ? rawSku : null,
      description: rawDescription.length > 0 ? rawDescription : null,
      priceCents: priceResult.priceCents,
      include,
      confidence: Math.min(...candidates.map((c) => c.confidence)),
      candidates,
    };
  });

  return {
    documentType: "CSV",
    products,
    warnings,
    needsReview: anyRowNeedsReview || !skuHeader || !descriptionHeader || !priceHeader,
  };
}
