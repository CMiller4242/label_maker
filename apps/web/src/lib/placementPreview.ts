import { buildPlacements, type PlacementInputProduct, type PlacementResult } from "@label-maker/label-layout";
import type { FixedGridLogicalGrid, LabelTemplate, Placement } from "@label-maker/shared";

export interface PreviewProduct {
  id: string;
  sku: string;
  description: string;
  priceCents: number;
}

/**
 * Computes the exact same placement plan the backend would produce for
 * this template/products/copiesPerProduct, using the identical
 * buildPlacements() function from @label-maker/label-layout (a pure,
 * browser-safe package - no reimplementation). Given the same inputs, the
 * real POST /label-runs call in stage 4 will compute byte-identical
 * placements to what this preview shows.
 */
export function computePlacementPreview(
  products: PreviewProduct[],
  template: Pick<LabelTemplate, "id" | "columns" | "rows" | "labelsPerSheet">,
  copiesPerProduct: number,
): PlacementResult {
  const inputs: PlacementInputProduct[] = products.map((p) => ({
    id: p.id,
    sku: p.sku,
    description: p.description,
    priceCents: p.priceCents,
    include: true,
  }));
  return buildPlacements(inputs, template, copiesPerProduct);
}

export type PreviewCell =
  | { kind: "writable"; rawColumnIndex: number; slotIndex: number; placement: Placement | null }
  | { kind: "spacer"; rawColumnIndex: number };

/**
 * Builds one sheet's physical row/cell layout directly from the template's
 * inspected logicalGrid - never a generic `columns`-wide grid. Each row has
 * exactly `logicalGrid.rawGridColumnWidthsTwips.length` cells, in raw
 * column order, so spacer columns render in their real physical position
 * (e.g. Avery 5155's interleaved label/spacer/label/... pattern).
 */
export function buildSheetGrid(
  logicalGrid: FixedGridLogicalGrid,
  sheetNumber: number,
  placements: Placement[],
): PreviewCell[][] {
  const placementBySlot = new Map<number, Placement>();
  for (const placement of placements) {
    if (placement.sheetNumber === sheetNumber) placementBySlot.set(placement.slotIndex, placement);
  }

  const rawColumnCount = logicalGrid.rawGridColumnWidthsTwips.length;
  const writableSet = new Set(logicalGrid.writableRawColumnIndexes);
  // Raw writable column index -> its logical column position (0-based, left to right).
  const logicalColumnByRawIndex = new Map(
    [...logicalGrid.writableRawColumnIndexes].sort((a, b) => a - b).map((rawIndex, logicalIndex) => [rawIndex, logicalIndex]),
  );

  const rows: PreviewCell[][] = [];
  for (let rowIndex = 0; rowIndex < logicalGrid.logicalRows; rowIndex++) {
    const row: PreviewCell[] = [];
    for (let rawColumnIndex = 0; rawColumnIndex < rawColumnCount; rawColumnIndex++) {
      if (writableSet.has(rawColumnIndex)) {
        const logicalColumn = logicalColumnByRawIndex.get(rawColumnIndex) ?? 0;
        const slotIndex = rowIndex * logicalGrid.logicalColumns + logicalColumn;
        row.push({
          kind: "writable",
          rawColumnIndex,
          slotIndex,
          placement: placementBySlot.get(slotIndex) ?? null,
        });
      } else {
        row.push({ kind: "spacer", rawColumnIndex });
      }
    }
    rows.push(row);
  }
  return rows;
}

/** CSS grid-template-columns value proportional to each raw column's real twip width - the layout source of truth is the template's own inspected geometry, never a hardcoded CSS constant. */
export function gridTemplateColumns(widthsTwips: number[]): string {
  return widthsTwips.map((w) => `${w}fr`).join(" ");
}
