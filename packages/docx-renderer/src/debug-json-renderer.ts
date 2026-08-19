import type { LabelTemplate, PlacementPlan } from "@label-maker/shared";
import {
  loadTemplateDocx as loadTemplateDocxFromBuffer,
  validateTemplateLayout,
} from "./template-loader.js";
import {
  DocxNotImplementedError,
  type LabelRenderer,
  type RenderResult,
  type TemplateLayoutValidationResult,
  type TemplatePackage,
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
 * Builds the structured debug JSON artifact for a placement plan: this is
 * the stand-in for real DOCX output until WordprocessingML rendering is
 * implemented (see README.md). It captures every field a real renderer
 * would need (sheet, slot, row, column, product fields, template id) so
 * downstream consumers/tests can assert on placement correctness today.
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
 * LabelRenderer implementation backing this milestone: template
 * loading/validation are implemented against real .docx bytes when given
 * (see template-loader.ts), but actual DOCX rendering
 * (renderFixedGridDocx/renderFloatingDocx) is intentionally not
 * implemented yet - callers needing an artifact today should call
 * renderDebugJsonArtifact() directly instead.
 */
export class DebugJsonRenderer implements LabelRenderer {
  private readonly readTemplateBytes: (templateStorageKey: string) => Promise<Buffer>;

  constructor(readTemplateBytes: (templateStorageKey: string) => Promise<Buffer>) {
    this.readTemplateBytes = readTemplateBytes;
  }

  async loadTemplateDocx(templateStorageKey: string): Promise<TemplatePackage> {
    const buffer = await this.readTemplateBytes(templateStorageKey);
    return loadTemplateDocxFromBuffer(templateStorageKey, buffer);
  }

  validateTemplateLayout(
    templatePackage: TemplatePackage,
    template: LabelTemplate,
  ): TemplateLayoutValidationResult {
    return validateTemplateLayout(templatePackage, template);
  }

  renderFixedGridDocx(_plan: PlacementPlan, _template: LabelTemplate): Promise<RenderResult> {
    throw new DocxNotImplementedError("renderFixedGridDocx");
  }

  renderFloatingDocx(_plan: PlacementPlan, _template: LabelTemplate): Promise<RenderResult> {
    throw new DocxNotImplementedError("renderFloatingDocx");
  }
}
