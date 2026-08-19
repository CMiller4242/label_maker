import type { Placement } from "@label-maker/shared";
import { PlacementValidationError, type LabelTemplateGeometry } from "./types.js";

/**
 * Validates a previously-built placement plan against its template
 * geometry: no duplicate occupied (sheetNumber, slotIndex) pairs, and no
 * slotIndex outside [0, labelsPerSheet). Throws PlacementValidationError
 * (with all issues collected, not just the first) if invalid.
 *
 * Intended as a defensive check before a plan is persisted or handed to a
 * renderer, independent of buildPlacements()'s own internal invariants.
 */
export function validatePlacements(placements: Placement[], template: LabelTemplateGeometry): void {
  const issues: string[] = [];
  const seenSlots = new Set<string>();

  for (const placement of placements) {
    if (placement.slotIndex < 0 || placement.slotIndex >= template.labelsPerSheet) {
      issues.push(
        `slotIndex ${placement.slotIndex} on sheet ${placement.sheetNumber} is outside valid range [0, ${template.labelsPerSheet}).`,
      );
    }

    const expectedRow = Math.floor(placement.slotIndex / template.columns);
    const expectedColumn = placement.slotIndex % template.columns;
    if (placement.row !== expectedRow || placement.column !== expectedColumn) {
      issues.push(
        `slotIndex ${placement.slotIndex} on sheet ${placement.sheetNumber} has row/column (${placement.row},${placement.column}) but expected (${expectedRow},${expectedColumn}).`,
      );
    }

    const slotKey = `${placement.sheetNumber}:${placement.slotIndex}`;
    if (seenSlots.has(slotKey)) {
      issues.push(
        `Duplicate placement at sheet ${placement.sheetNumber}, slot ${placement.slotIndex}.`,
      );
    }
    seenSlots.add(slotKey);
  }

  if (issues.length > 0) {
    throw new PlacementValidationError(issues);
  }
}
