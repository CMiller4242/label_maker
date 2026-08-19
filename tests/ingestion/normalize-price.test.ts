import { describe, expect, it } from "vitest";
import { normalizePriceToCents } from "@label-maker/ingestion";

describe("normalizePriceToCents", () => {
  it.each([
    ["$14.49", 1449],
    ["14.49", 1449],
    ["$1,234.50", 123450],
    ["14.49 each", 1449],
    ["USD 14.49", 1449],
    ["$5", 500],
    ["5.5", 550],
  ])("parses %s -> %i cents", (input, expected) => {
    const result = normalizePriceToCents(input);
    expect(result.priceCents).toBe(expected);
    expect(result.malformed).toBe(false);
  });

  it("treats blank input as null, not malformed", () => {
    const result = normalizePriceToCents("   ");
    expect(result.priceCents).toBeNull();
    expect(result.malformed).toBe(false);
  });

  it("flags non-numeric input as malformed", () => {
    const result = normalizePriceToCents("not-a-price");
    expect(result.priceCents).toBeNull();
    expect(result.malformed).toBe(true);
  });
});
