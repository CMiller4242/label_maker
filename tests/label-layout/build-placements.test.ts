import { describe, expect, it } from "vitest";
import {
  AVERY_5155_GEOMETRY,
  buildPlacements,
  InvalidCopiesPerProductError,
  InvalidTemplateGeometryError,
  MissingRequiredProductDataError,
  validatePlacements,
  type PlacementInputProduct,
} from "@label-maker/label-layout";

function product(
  id: string,
  overrides: Partial<PlacementInputProduct> = {},
): PlacementInputProduct {
  return {
    id,
    sku: `SKU-${id}`,
    description: `Product ${id}`,
    priceCents: 999,
    include: true,
    ...overrides,
  };
}

describe("buildPlacements on Avery 5155 (4 columns x 15 rows, 60/sheet)", () => {
  it("places 1 product x 8 copies as 8 placements on sheet 1, slots 0-7", () => {
    const result = buildPlacements([product("A")], AVERY_5155_GEOMETRY, 8);

    expect(result.totalPlacements).toBe(8);
    expect(result.totalSheets).toBe(1);
    expect(result.placements.every((p) => p.sheetNumber === 1)).toBe(true);
    expect(result.placements.map((p) => p.slotIndex)).toEqual([0, 1, 2, 3, 4, 5, 6, 7]);
    // 8 copies of one product occupy exactly two full rows of four labels.
    expect(result.placements.map((p) => p.row)).toEqual([0, 0, 0, 0, 1, 1, 1, 1]);
    expect(result.placements.map((p) => p.column)).toEqual([0, 1, 2, 3, 0, 1, 2, 3]);
  });

  it("places 7 products x 8 copies as 56 placements on sheet 1", () => {
    const products = Array.from({ length: 7 }, (_, i) => product(`P${i}`));
    const result = buildPlacements(products, AVERY_5155_GEOMETRY, 8);

    expect(result.totalPlacements).toBe(56);
    expect(result.totalSheets).toBe(1);
    expect(result.placements.every((p) => p.sheetNumber === 1)).toBe(true);
    expect(result.unfilledSlots).toHaveLength(4);
    expect(result.unfilledSlots.every((s) => s.sheetNumber === 1)).toBe(true);
  });

  it("places 8 products x 8 copies as 60 on sheet 1 and 4 on sheet 2", () => {
    const products = Array.from({ length: 8 }, (_, i) => product(`P${i}`));
    const result = buildPlacements(products, AVERY_5155_GEOMETRY, 8);

    expect(result.totalPlacements).toBe(64);
    expect(result.totalSheets).toBe(2);

    const sheet1 = result.placements.filter((p) => p.sheetNumber === 1);
    const sheet2 = result.placements.filter((p) => p.sheetNumber === 2);
    expect(sheet1).toHaveLength(60);
    expect(sheet2).toHaveLength(4);
    expect(sheet2.map((p) => p.slotIndex)).toEqual([0, 1, 2, 3]);

    validatePlacements(result.placements, AVERY_5155_GEOMETRY);
  });

  it("produces no placements for excluded products", () => {
    const result = buildPlacements(
      [product("A", { include: false }), product("B", { include: false })],
      AVERY_5155_GEOMETRY,
      8,
    );

    expect(result.totalPlacements).toBe(0);
    expect(result.totalSheets).toBe(0);
    expect(result.placements).toEqual([]);
  });

  it("fills left-to-right correctly with 4 copies per product", () => {
    const result = buildPlacements([product("A"), product("B")], AVERY_5155_GEOMETRY, 4);

    expect(result.totalPlacements).toBe(8);
    expect(
      result.placements.map((p) => ({ productId: p.productId, row: p.row, column: p.column })),
    ).toEqual([
      { productId: "A", row: 0, column: 0 },
      { productId: "A", row: 0, column: 1 },
      { productId: "A", row: 0, column: 2 },
      { productId: "A", row: 0, column: 3 },
      { productId: "B", row: 1, column: 0 },
      { productId: "B", row: 1, column: 1 },
      { productId: "B", row: 1, column: 2 },
      { productId: "B", row: 1, column: 3 },
    ]);
  });

  it("rejects included products with missing required data before placement", () => {
    expect(() =>
      buildPlacements(
        [product("A", { sku: null }), product("B", { description: null })],
        AVERY_5155_GEOMETRY,
        8,
      ),
    ).toThrow(MissingRequiredProductDataError);
  });

  it("does not reject missing data on excluded products", () => {
    const result = buildPlacements(
      [product("A", { include: false, sku: null }), product("B")],
      AVERY_5155_GEOMETRY,
      1,
    );
    expect(result.totalPlacements).toBe(1);
  });

  it("rejects copiesPerProduct <= 0", () => {
    expect(() => buildPlacements([product("A")], AVERY_5155_GEOMETRY, 0)).toThrow(
      InvalidCopiesPerProductError,
    );
    expect(() => buildPlacements([product("A")], AVERY_5155_GEOMETRY, -3)).toThrow(
      InvalidCopiesPerProductError,
    );
  });

  it("rejects invalid template geometry", () => {
    expect(() =>
      buildPlacements([product("A")], { id: "bad", columns: 0, rows: 15, labelsPerSheet: 60 }, 1),
    ).toThrow(InvalidTemplateGeometryError);
    expect(() =>
      buildPlacements([product("A")], { id: "bad", columns: 4, rows: 15, labelsPerSheet: 61 }, 1),
    ).toThrow(InvalidTemplateGeometryError);
  });
});
