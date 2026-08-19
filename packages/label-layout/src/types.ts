import type { Placement, PlacementPlan } from "@label-maker/shared";

export type { Placement, PlacementPlan };

/** Minimal geometry the placement engine needs from a LabelTemplate. */
export interface LabelTemplateGeometry {
  id: string;
  columns: number;
  rows: number;
  labelsPerSheet: number;
}

/**
 * A product eligible for placement. Only APPROVED/AUTO_ACCEPTED, included
 * products with complete required data should ever reach this engine -
 * that filtering happens upstream (product-service); the engine itself
 * re-validates required fields defensively and rejects the whole run if
 * any included product is incomplete, rather than silently dropping rows.
 */
export interface PlacementInputProduct {
  id: string;
  sku: string | null;
  description: string | null;
  priceCents: number | null;
  include: boolean;
}

export class InvalidTemplateGeometryError extends Error {
  readonly code = "INVALID_TEMPLATE_GEOMETRY";
}

export class InvalidCopiesPerProductError extends Error {
  readonly code = "INVALID_COPIES_PER_PRODUCT";
}

export class MissingRequiredProductDataError extends Error {
  readonly code = "MISSING_REQUIRED_PRODUCT_DATA";
  readonly productIds: string[];

  constructor(productIds: string[]) {
    super(
      `${productIds.length} included product(s) are missing required sku/description/priceCents and cannot be placed: ${productIds.join(", ")}`,
    );
    this.productIds = productIds;
  }
}

export class PlacementValidationError extends Error {
  readonly code = "PLACEMENT_VALIDATION_ERROR";
  readonly issues: string[];

  constructor(issues: string[]) {
    super(`Placement plan failed validation: ${issues.join("; ")}`);
    this.issues = issues;
  }
}
