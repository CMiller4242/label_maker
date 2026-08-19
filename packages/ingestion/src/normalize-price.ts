export interface NormalizePriceResult {
  priceCents: number | null;
  /** true when input was non-blank but could not be parsed as a price */
  malformed: boolean;
}

/**
 * Parses common price strings into integer cents. Never performs float
 * arithmetic on the dollar amount; the decimal string is split and each
 * side is parsed as an integer to avoid binary floating-point rounding
 * (e.g. avoiding `14.49 * 100 === 1448.9999...`-style bugs).
 *
 * Accepts forms such as:
 *   "$14.49", "14.49", "$1,234.50", "14.49 each", "USD 14.49"
 */
export function normalizePriceToCents(rawInput: string): NormalizePriceResult {
  const trimmed = rawInput.trim();
  if (trimmed.length === 0) {
    return { priceCents: null, malformed: false };
  }

  // Strip currency symbols/codes, thousands separators, and trailing words
  // like "each"/"ea" while keeping the numeric core and its decimal point.
  const stripped = trimmed
    .replace(/usd/gi, "")
    .replace(/[$,]/g, "")
    .replace(/\beach\b|\bea\.?\b/gi, "")
    .trim();

  const match = /^(\d+)(?:\.(\d{1,2}))?$/.exec(stripped);
  if (!match) {
    return { priceCents: null, malformed: true };
  }

  const wholePart = match[1] ?? "0";
  const fractionalPart = (match[2] ?? "").padEnd(2, "0");

  const whole = Number.parseInt(wholePart, 10);
  const cents = Number.parseInt(fractionalPart, 10);

  if (!Number.isFinite(whole) || !Number.isFinite(cents)) {
    return { priceCents: null, malformed: true };
  }

  return { priceCents: whole * 100 + cents, malformed: false };
}
