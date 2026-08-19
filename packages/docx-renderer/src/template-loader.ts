import JSZip from "jszip";
import type { LabelTemplate } from "@label-maker/shared";
import type { TemplateLayoutValidationResult, TemplatePackage } from "./types.js";

const WORD_DOCUMENT_XML_PATH = "word/document.xml";

/**
 * Opens a .docx/.dotx template package (a zip container per OOXML) and
 * extracts document.xml for later WordprocessingML manipulation.
 *
 * TODO: once a verified physical Avery template exists (see
 * fixtures/label-templates/README.md), this will read templateBuffer from
 * storage via templateStorageKey; for now callers pass the bytes directly
 * since no verified template is checked in yet.
 */
export async function loadTemplateDocx(
  templateStorageKey: string,
  templateBuffer: Buffer,
): Promise<TemplatePackage> {
  const zip = await JSZip.loadAsync(templateBuffer);
  const documentXmlFile = zip.file(WORD_DOCUMENT_XML_PATH);
  if (!documentXmlFile) {
    throw new Error(
      `Template "${templateStorageKey}" is not a valid .docx package (missing ${WORD_DOCUMENT_XML_PATH}).`,
    );
  }
  const documentXml = await documentXmlFile.async("text");

  return { templateStorageKey, zip, documentXml };
}

/**
 * Validates that a loaded template's fixed-grid table structure matches
 * the LabelTemplate's declared columns/rows/labelsPerSheet.
 *
 * TODO: this is a structural placeholder. The real implementation will
 * parse document.xml's <w:tbl> elements, count <w:tr>/<w:tc>, and confirm
 * `tblLayout w:type="fixed"` plus exact row heights, per the rendering
 * strategy documented in README.md.
 */
export function validateTemplateLayout(
  templatePackage: TemplatePackage,
  template: LabelTemplate,
): TemplateLayoutValidationResult {
  const issues: TemplateLayoutValidationResult["issues"] = [];

  if (template.renderingMode === "FIXED_GRID" && !templatePackage.documentXml.includes("<w:tbl")) {
    issues.push({
      code: "NO_TABLE_FOUND",
      message: "FIXED_GRID template does not contain any Word table (<w:tbl>) in document.xml.",
    });
  }

  // TODO: verify tblLayout is fixed (not autofit), row count matches
  // template.rows, column count matches template.columns, and cell margins/
  // paragraph spacing are explicit, per README.md's rendering strategy.

  return { valid: issues.length === 0, issues };
}
