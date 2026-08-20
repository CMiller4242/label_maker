/** Formats integer cents as a plain dollar string for an editable input, e.g. 1449 -> "14.49". Never touches floating point. */
export function centsToDollarsInputValue(cents: number | null): string {
  if (cents === null) return "";
  const dollars = Math.trunc(cents / 100);
  const remainder = Math.abs(cents % 100);
  return `${dollars}.${String(remainder).padStart(2, "0")}`;
}

/** Parses a user-typed dollar string (e.g. "14.49", "14", "$14.5") into integer cents, or null if unparseable/negative. */
export function parseDollarsToCents(value: string): number | null {
  const trimmed = value.trim().replace(/^\$/, "");
  if (trimmed.length === 0) return null;
  if (!/^\d+(\.\d{1,2})?$/.test(trimmed)) return null;
  const [wholePart, fractionPart = ""] = trimmed.split(".");
  const cents = Number(wholePart) * 100 + Number(fractionPart.padEnd(2, "0"));
  return Number.isSafeInteger(cents) ? cents : null;
}
