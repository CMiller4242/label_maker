import { describe, expect, it } from "vitest";
import { isProductListReady, productRowIssue } from "@label-maker/web/lib/productValidation";

function product(overrides: Partial<Parameters<typeof productRowIssue>[0]> = {}) {
  return {
    include: true,
    sku: "SKU-1",
    description: "Widget",
    priceCents: 1099,
    status: "APPROVED" as const,
    ...overrides,
  };
}

describe("productRowIssue", () => {
  it("returns null for a valid, approved, included row", () => {
    expect(productRowIssue(product())).toBeNull();
  });

  it("returns null for an excluded row regardless of other fields", () => {
    expect(productRowIssue(product({ include: false, sku: null, description: null }))).toBeNull();
  });

  it("flags a missing sku", () => {
    expect(productRowIssue(product({ sku: null }))).toBe("SKU is required.");
    expect(productRowIssue(product({ sku: "   " }))).toBe("SKU is required.");
  });

  it("flags a missing product name", () => {
    expect(productRowIssue(product({ description: null }))).toBe("Product name is required.");
  });

  it("flags a missing price", () => {
    expect(productRowIssue(product({ priceCents: null }))).toBe("Price is required.");
  });

  it("flags a negative price", () => {
    expect(productRowIssue(product({ priceCents: -1 }))).toBe("Price cannot be negative.");
  });

  it("flags a row still awaiting review", () => {
    expect(productRowIssue(product({ status: "NEEDS_REVIEW" }))).toBe(
      "Needs approval before it can be generated.",
    );
  });

  it("accepts AUTO_ACCEPTED as generation-ready, same as APPROVED", () => {
    expect(productRowIssue(product({ status: "AUTO_ACCEPTED" }))).toBeNull();
  });
});

describe("isProductListReady", () => {
  it("is false when no row is included", () => {
    expect(isProductListReady([product({ include: false })])).toBe(false);
  });

  it("is false when an included row has an issue", () => {
    expect(isProductListReady([product(), product({ sku: null })])).toBe(false);
  });

  it("is true when every included row is generation-ready", () => {
    expect(
      isProductListReady([product(), product({ include: false, sku: null }), product({ sku: "SKU-2" })]),
    ).toBe(true);
  });

  it("is false for an empty list", () => {
    expect(isProductListReady([])).toBe(false);
  });
});
