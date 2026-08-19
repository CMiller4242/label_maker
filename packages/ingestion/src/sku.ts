/**
 * Configurable regex for "SKU-like" tokens used for generic candidate
 * discovery (e.g. scanning raw PDF text). This is intentionally generic and
 * not tuned to any specific product deck's SKU format.
 *
 * Matches tokens of 3-20 characters made of uppercase letters, digits, and
 * internal hyphens/slashes/dots, containing at least one digit (to avoid
 * matching plain words).
 */
export const DEFAULT_SKU_REGEX =
  /\b(?=[A-Z0-9][A-Z0-9./-]{2,19}\b)(?=[A-Z0-9./-]*\d)[A-Z0-9./-]{3,20}\b/g;

export function findSkuLikeTokens(text: string, regex: RegExp = DEFAULT_SKU_REGEX): string[] {
  const flags = regex.flags.includes("g") ? regex.flags : `${regex.flags}g`;
  const matches = text.match(new RegExp(regex.source, flags));
  return matches ?? [];
}

/** Basic structural check: is this string plausible as a SKU value at all? */
export function isPlausibleSku(value: string): boolean {
  const trimmed = value.trim();
  if (trimmed.length < 2 || trimmed.length > 40) return false;
  return /[A-Za-z0-9]/.test(trimmed);
}
