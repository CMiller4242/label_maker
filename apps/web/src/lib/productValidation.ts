import type { Product } from "@label-maker/shared";

/**
 * A single, shared definition of "ready to generate" for one product row,
 * used by both the product-review table (to block Continue) and the
 * label-plan preview (to decide what's actually includable). Mirrors
 * @label-maker/label-layout's buildPlacements() requirements exactly - a
 * row failing this check is exactly a row buildPlacements would reject
 * with MissingRequiredProductDataError - plus the server-side inclusion
 * filter createLabelRun() applies (include=true AND status IN (APPROVED,
 * AUTO_ACCEPTED)), so a row that "looks" ready but is still NEEDS_REVIEW
 * is caught here too, instead of silently vanishing from a real run.
 */
export function productRowIssue(product: Pick<Product, "include" | "sku" | "description" | "priceCents" | "status">): string | null {
  if (!product.include) return null; // excluded rows are never "invalid" - they're just not going

  if (!product.sku || product.sku.trim().length === 0) return "SKU is required.";
  if (!product.description || product.description.trim().length === 0) {
    return "Product name is required.";
  }
  if (product.priceCents === null || product.priceCents === undefined) return "Price is required.";
  if (product.priceCents < 0) return "Price cannot be negative.";
  if (product.status !== "APPROVED" && product.status !== "AUTO_ACCEPTED") {
    return "Needs approval before it can be generated.";
  }
  return null;
}

/** True when every included row is generation-ready and at least one row is included. */
export function isProductListReady(products: Array<Pick<Product, "include" | "sku" | "description" | "priceCents" | "status">>): boolean {
  const included = products.filter((p) => p.include);
  return included.length > 0 && included.every((p) => productRowIssue(p) === null);
}
