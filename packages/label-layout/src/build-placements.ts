import type { Placement } from "@label-maker/shared";
import {
  InvalidCopiesPerProductError,
  InvalidTemplateGeometryError,
  MissingRequiredProductDataError,
  type LabelTemplateGeometry,
  type PlacementInputProduct,
} from "./types.js";

export interface UnfilledSlot {
  sheetNumber: number;
  slotIndex: number;
  row: number;
  column: number;
}

export interface PlacementResult {
  labelTemplateId: string;
  copiesPerProduct: number;
  totalSheets: number;
  totalPlacements: number;
  placements: Placement[];
  /** Empty slots on the final (partial) sheet, so a renderer knows what to leave blank. */
  unfilledSlots: UnfilledSlot[];
}

function assertValidGeometry(template: LabelTemplateGeometry): void {
  if (
    !Number.isInteger(template.columns) ||
    template.columns <= 0 ||
    !Number.isInteger(template.rows) ||
    template.rows <= 0 ||
    !Number.isInteger(template.labelsPerSheet) ||
    template.labelsPerSheet <= 0
  ) {
    throw new InvalidTemplateGeometryError(
      `Template "${template.id}" has invalid geometry: columns=${template.columns}, rows=${template.rows}, labelsPerSheet=${template.labelsPerSheet} (all must be positive integers).`,
    );
  }
  if (template.columns * template.rows !== template.labelsPerSheet) {
    throw new InvalidTemplateGeometryError(
      `Template "${template.id}" geometry is inconsistent: columns (${template.columns}) x rows (${template.rows}) !== labelsPerSheet (${template.labelsPerSheet}).`,
    );
  }
}

/**
 * Deterministic label placement engine.
 *
 * - Filters out excluded products.
 * - Expands each included product exactly copiesPerProduct times.
 * - Fills sheet slots left-to-right, then top-to-bottom, with a single
 *   continuous global index (no page break forced after each product).
 * - Rejects the whole run (throws) if any included product is missing
 *   required sku/description/priceCents, or if geometry/copiesPerProduct
 *   are invalid - never silently drops or half-places a product.
 */
export function buildPlacements(
  products: PlacementInputProduct[],
  template: LabelTemplateGeometry,
  copiesPerProduct: number,
): PlacementResult {
  assertValidGeometry(template);

  if (!Number.isInteger(copiesPerProduct) || copiesPerProduct <= 0) {
    throw new InvalidCopiesPerProductError(
      `copiesPerProduct must be a positive integer, got ${copiesPerProduct}.`,
    );
  }

  const includedProducts = products.filter((p) => p.include);

  const invalidProductIds = includedProducts
    .filter((p) => !p.sku || !p.description || p.priceCents === null || p.priceCents === undefined)
    .map((p) => p.id);
  if (invalidProductIds.length > 0) {
    throw new MissingRequiredProductDataError(invalidProductIds);
  }

  const placements: Placement[] = [];
  let globalIndex = 0;

  for (const product of includedProducts) {
    for (let copyIndex = 0; copyIndex < copiesPerProduct; copyIndex++) {
      const sheetNumber = Math.floor(globalIndex / template.labelsPerSheet) + 1;
      const slotIndex = globalIndex % template.labelsPerSheet;
      const row = Math.floor(slotIndex / template.columns);
      const column = slotIndex % template.columns;

      placements.push({
        sheetNumber,
        slotIndex,
        row,
        column,
        productId: product.id,
        sku: product.sku,
        description: product.description,
        priceCents: product.priceCents,
        copyIndex,
      });

      globalIndex += 1;
    }
  }

  const totalPlacements = placements.length;
  const totalSheets =
    totalPlacements === 0 ? 0 : Math.ceil(totalPlacements / template.labelsPerSheet);

  const unfilledSlots: UnfilledSlot[] = [];
  if (totalPlacements > 0) {
    const slotsOnLastSheet = totalPlacements % template.labelsPerSheet;
    if (slotsOnLastSheet !== 0) {
      for (let slotIndex = slotsOnLastSheet; slotIndex < template.labelsPerSheet; slotIndex++) {
        unfilledSlots.push({
          sheetNumber: totalSheets,
          slotIndex,
          row: Math.floor(slotIndex / template.columns),
          column: slotIndex % template.columns,
        });
      }
    }
  }

  return {
    labelTemplateId: template.id,
    copiesPerProduct,
    totalSheets,
    totalPlacements,
    placements,
    unfilledSlots,
  };
}
