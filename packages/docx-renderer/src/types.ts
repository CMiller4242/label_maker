import type JSZip from "jszip";
import type { LabelTemplate, PlacementPlan } from "@label-maker/shared";

export class DocxNotImplementedError extends Error {
  readonly code = "DOCX_NOT_IMPLEMENTED";

  constructor(operation: string) {
    super(
      `${operation} is not implemented yet. DOCX generation is the next milestone; ` +
        `see packages/docx-renderer/README.md for the planned WordprocessingML strategy.`,
    );
    this.name = "DocxNotImplementedError";
  }
}

/** An opened .docx/.dotx template package, prior to any content manipulation. */
export interface TemplatePackage {
  templateStorageKey: string;
  zip: JSZip;
  /** Raw document.xml contents, for future WordprocessingML manipulation. */
  documentXml: string;
}

export interface TemplateLayoutValidationIssue {
  code: string;
  message: string;
}

export interface TemplateLayoutValidationResult {
  valid: boolean;
  issues: TemplateLayoutValidationIssue[];
}

export interface RenderResult {
  /** Bytes of the produced artifact (a debug JSON document today; a .docx once implemented). */
  buffer: Buffer;
  mimeType: string;
  fileExtension: string;
}

/**
 * The renderer contract every concrete implementation (debug JSON today,
 * WordprocessingML DOCX renderer eventually) must satisfy. Keeping this
 * interface stable lets apps/worker's build-label-run processor swap
 * implementations without changing call sites.
 */
export interface LabelRenderer {
  /** Opens a template package (.docx/.dotx) from storage. TODO: real implementation reads template bytes via a template loader (see template-loader.ts) once verified templates exist. */
  loadTemplateDocx(templateStorageKey: string): Promise<TemplatePackage>;

  /** Checks a loaded template's table/row/column structure against the LabelTemplate's declared geometry. */
  validateTemplateLayout(
    templatePackage: TemplatePackage,
    template: LabelTemplate,
  ): TemplateLayoutValidationResult;

  /** Renders a FIXED_GRID template (e.g. Avery 5155) by cloning fixed-layout Word tables. Throws DocxNotImplementedError today. */
  renderFixedGridDocx(plan: PlacementPlan, template: LabelTemplate): Promise<RenderResult>;

  /** Renders a FLOATING template (free-positioned tags/labels). Throws DocxNotImplementedError today. */
  renderFloatingDocx(plan: PlacementPlan, template: LabelTemplate): Promise<RenderResult>;
}
