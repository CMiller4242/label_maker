export type LogicalField = "SKU" | "DESCRIPTION" | "PRICE" | "INCLUDE_STATUS";

/**
 * Known column-header spellings we map to each logical field, lowercased for
 * comparison. Ordered by nothing in particular; ambiguity (multiple source
 * headers mapping to the same logical field) is detected by the caller, not
 * resolved here.
 */
const HEADER_CANDIDATES: Record<LogicalField, string[]> = {
  SKU: [
    "sku",
    "item number",
    "item no",
    "item no.",
    "item #",
    "item",
    "product code",
    "product number",
  ],
  DESCRIPTION: ["description", "product description", "item description", "product name", "name"],
  PRICE: ["as low as", "as low as price", "price", "lowest price", "low price"],
  INCLUDE_STATUS: ["include", "status", "active", "available"],
};

function normalizeHeaderText(header: string): string {
  return header.trim().toLowerCase().replace(/\s+/g, " ");
}

export interface HeaderMapping {
  /** logical field -> matched source header(s), in original casing */
  matches: Record<LogicalField, string[]>;
  /** logical fields matched by more than one source header (ambiguous) */
  ambiguousFields: LogicalField[];
  /** source headers that did not match any known logical field */
  unmatchedHeaders: string[];
}

/**
 * Maps raw spreadsheet/CSV header strings to logical product fields.
 * Multiple headers matching the same logical field is treated as ambiguous
 * and surfaced to the caller rather than silently picking one.
 */
export function normalizeHeaders(rawHeaders: string[]): HeaderMapping {
  const matches: Record<LogicalField, string[]> = {
    SKU: [],
    DESCRIPTION: [],
    PRICE: [],
    INCLUDE_STATUS: [],
  };
  const unmatchedHeaders: string[] = [];

  for (const rawHeader of rawHeaders) {
    const normalized = normalizeHeaderText(rawHeader);
    let matchedField: LogicalField | undefined;

    for (const field of Object.keys(HEADER_CANDIDATES) as LogicalField[]) {
      if (HEADER_CANDIDATES[field].includes(normalized)) {
        matchedField = field;
        break;
      }
    }

    if (matchedField) {
      matches[matchedField].push(rawHeader);
    } else {
      unmatchedHeaders.push(rawHeader);
    }
  }

  const ambiguousFields = (Object.keys(matches) as LogicalField[]).filter(
    (field) => matches[field].length > 1,
  );

  return { matches, ambiguousFields, unmatchedHeaders };
}
