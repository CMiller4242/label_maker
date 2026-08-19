import type { LabelTemplate, PlacementPlan } from "@label-maker/shared";
import { loadTemplateDocx, type TemplatePackage } from "./ooxml.js";
import { validateTemplateLayout } from "./template-validation.js";
import {
  DocxNotImplementedError,
  type LabelRenderer,
  type RenderResult,
  type TemplateLayoutValidationResult,
} from "./types.js";

export interface DebugPlacementEntry {
  sheetNumber: number;
  slotIndex: number;
  row: number;
  column: number;
  product: {
    productId: string;
    sku: string | null;
    description: string | null;
    priceCents: number | null;
    copyIndex: number;
  };
}

export interface DebugArtifact {
  artifactType: "label-run-debug-json";
  generatedAt: string;
  labelTemplateId: string;
  templateDisplayName: string;
  renderingMode: LabelTemplate["renderingMode"];
  copiesPerProduct: number;
  totalSheets: number;
  totalPlacements: number;
  placements: DebugPlacementEntry[];
}

/**
 * Builds the structured debug JSON artifact for a placement plan. Useful as
 * an optional development-time output alongside (or instead of) a real
 * DOCX, since it captures every field a renderer would need (sheet, slot,
 * row, column, product fields, template id) in a format that's trivial to
 * diff in tests and code review.
 */
export function buildDebugArtifact(plan: PlacementPlan, template: LabelTemplate): DebugArtifact {
  return {
    artifactType: "label-run-debug-json",
    generatedAt: new Date().toISOString(),
    labelTemplateId: template.id,
    templateDisplayName: template.displayName,
    renderingMode: template.renderingMode,
    copiesPerProduct: plan.copiesPerProduct,
    totalSheets: plan.totalSheets,
    totalPlacements: plan.totalPlacements,
    placements: plan.placements.map((p) => ({
      sheetNumber: p.sheetNumber,
      slotIndex: p.slotIndex,
      row: p.row,
      column: p.column,
      product: {
        productId: p.productId,
        sku: p.sku,
        description: p.description,
        priceCents: p.priceCents,
        copyIndex: p.copyIndex,
      },
    })),
  };
}

/** Serializes a DebugArtifact to a RenderResult (bytes + content type) ready for storage. */
export function renderDebugJsonArtifact(
  plan: PlacementPlan,
  template: LabelTemplate,
): RenderResult {
  const artifact = buildDebugArtifact(plan, template);
  const buffer = Buffer.from(JSON.stringify(artifact, null, 2), "utf-8");
  return { buffer, mimeType: "application/json", fileExtension: "json" };
}

/**
 * A LabelRenderer implementation that only ever produces the debug JSON
 * artifact - its renderFixedGridDocx/renderFloatingDocx methods always
 * throw. Use the standalone renderFixedGridDocx() from
 * fixed-grid-renderer.ts directly for real DOCX output; this class exists
 * for callers that want the debug-only artifact behind the same interface.
 */
export class DebugJsonRenderer implements LabelRenderer {
  private readonly readTemplateBytes: (templateStorageKey: string) => Promise<Buffer>;

  constructor(readTemplateBytes: (templateStorageKey: string) => Promise<Buffer>) {
    this.readTemplateBytes = readTemplateBytes;
  }

  async loadTemplateDocx(templateStorageKey: string): Promise<TemplatePackage> {
    const buffer = await this.readTemplateBytes(templateStorageKey);
    return loadTemplateDocx(buffer, templateStorageKey);
  }

  validateTemplateLayout(
    templatePackage: TemplatePackage,
    template: LabelTemplate,
  ): TemplateLayoutValidationResult {
    return validateTemplateLayout(templatePackage, template);
  }

  renderFixedGridDocx(_plan: PlacementPlan, _template: LabelTemplate): Promise<RenderResult> {
    throw new DocxNotImplementedError("DebugJsonRenderer.renderFixedGridDocx");
  }

  renderFloatingDocx(_plan: PlacementPlan, _template: LabelTemplate): Promise<RenderResult> {
    throw new DocxNotImplementedError("DebugJsonRenderer.renderFloatingDocx");
  }
}
