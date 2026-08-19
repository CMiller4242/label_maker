import JSZip from "jszip";
import type { LabelTemplate, Placement, PlacementPlan } from "@label-maker/shared";
import {
  assertGeneratedDocxPackageIsValid,
  childrenOf,
  cloneNode,
  findDirectChildren,
  findFirstByTag,
  getAttr,
  loadTemplateDocx,
  makeElement,
  makeEmptyElement,
  makeTextElement,
  setChildren,
  buildXml,
  tagNameOf,
  type XmlNode,
} from "./ooxml.js";
import { findTableNode, validateFixedGridTemplate } from "./template-validation.js";
import {
  DocxNotImplementedError,
  UnsupportedFloatingTemplateError,
  InvalidFixedGridTemplateError,
  MissingLabelTextStyleConfigError,
  labelTextStyleConfigSchema,
  type LabelTextStyleConfig,
  type RenderResult,
} from "./types.js";

const DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

/**
 * Formats integer cents as a dollar string, e.g. 1449 -> "$14.49". Never
 * touches floating point: dollars/cents are derived via integer division
 * and modulo only.
 */
export function formatCentsAsDollars(cents: number): string {
  if (!Number.isInteger(cents) || cents < 0) {
    throw new RangeError(`Expected a non-negative integer number of cents, got ${cents}.`);
  }
  const dollars = Math.trunc(cents / 100);
  const remainder = cents % 100;
  return `$${dollars}.${String(remainder).padStart(2, "0")}`;
}

function readLabelTextStyle(template: LabelTemplate): LabelTextStyleConfig {
  const raw = template.configJson["labelTextStyle"];
  const result = labelTextStyleConfigSchema.safeParse(raw);
  if (!result.success) {
    throw new MissingLabelTextStyleConfigError(template.id, result.error.message);
  }
  return result.data;
}

function buildTextParagraph(
  text: string,
  fontSizeHalfPoints: number,
  style: LabelTextStyleConfig,
): XmlNode {
  const pPr = makeElement("w:pPr", {}, [
    makeEmptyElement("w:spacing", {
      "w:before": String(style.paragraphSpacingBeforeTwips),
      "w:after": String(style.paragraphSpacingAfterTwips),
      "w:line": String(style.lineSpacing.line),
      "w:lineRule": style.lineSpacing.lineRule,
    }),
    makeEmptyElement("w:jc", { "w:val": style.horizontalAlignment }),
  ]);

  const rPrChildren: XmlNode[] = [
    makeEmptyElement("w:rFonts", {
      "w:ascii": style.fontFamily,
      "w:hAnsi": style.fontFamily,
      "w:cs": style.fontFamily,
    }),
  ];
  if (style.bold) rPrChildren.push(makeEmptyElement("w:b", {}));
  rPrChildren.push(makeEmptyElement("w:color", { "w:val": style.colorHex }));
  rPrChildren.push(makeEmptyElement("w:sz", { "w:val": String(fontSizeHalfPoints) }));
  rPrChildren.push(makeEmptyElement("w:szCs", { "w:val": String(fontSizeHalfPoints) }));
  const rPr = makeElement("w:rPr", {}, rPrChildren);

  const run = makeElement("w:r", {}, [
    rPr,
    makeTextElement("w:t", { "xml:space": "preserve" }, text),
  ]);

  return makeElement("w:p", {}, [pPr, run]);
}

/**
 * Ensures a cell's tcPr carries explicit margins consistent with `override`
 * when provided (config-driven safe insets), otherwise leaves whatever
 * margins were already cloned from the source template untouched. Only
 * touches the single w:tcMar child; every other tcPr child (tcW, borders,
 * etc.) is preserved as-is. NOTE: this assumes a simple label-cell tcPr
 * shape (no textDirection/tcFitText/hideMark) - true for the grid templates
 * this renderer targets.
 */
function ensureExplicitCellMargins(
  tcPr: XmlNode,
  override: LabelTextStyleConfig["safeInsets"] | undefined,
): void {
  if (!override) return;
  const children = childrenOf(tcPr);
  const existingIndex = children.findIndex((c) => tagNameOf(c) === "w:tcMar");
  const tcMar = makeElement("w:tcMar", {}, [
    makeEmptyElement("w:top", { "w:w": String(override.topTwips), "w:type": "dxa" }),
    makeEmptyElement("w:left", { "w:w": String(override.leftTwips), "w:type": "dxa" }),
    makeEmptyElement("w:bottom", { "w:w": String(override.bottomTwips), "w:type": "dxa" }),
    makeEmptyElement("w:right", { "w:w": String(override.rightTwips), "w:type": "dxa" }),
  ]);
  if (existingIndex >= 0) {
    children[existingIndex] = tcMar;
  } else {
    children.push(tcMar);
  }
  setChildren(tcPr, children);
}

function ensureCellVerticalAlignment(
  tcPr: XmlNode,
  vAlign: LabelTextStyleConfig["verticalAlignment"],
): void {
  const children = childrenOf(tcPr).filter((c) => tagNameOf(c) !== "w:vAlign");
  children.push(makeEmptyElement("w:vAlign", { "w:val": vAlign }));
  setChildren(tcPr, children);
}

/** Replaces a cell's body content (paragraphs) with the 3 label lines, preserving/adjusting tcPr. */
function fillCell(
  tc: XmlNode,
  texts: { sku: string; description: string; priceLine: string },
  style: LabelTextStyleConfig,
): void {
  const children = childrenOf(tc);
  let tcPr = findDirectChildren(children, "w:tcPr")[0];
  if (!tcPr) {
    tcPr = makeElement("w:tcPr", {}, []);
  }

  ensureExplicitCellMargins(tcPr, style.safeInsets);
  ensureCellVerticalAlignment(tcPr, style.verticalAlignment);

  const paragraphs = [
    buildTextParagraph(texts.sku, style.skuFontSizeHalfPoints, style),
    buildTextParagraph(texts.description, style.descriptionFontSizeHalfPoints, style),
    buildTextParagraph(texts.priceLine, style.priceFontSizeHalfPoints, style),
  ];

  // tcPr must be the first child of w:tc when present; everything else
  // (previously blank paragraph(s)) is fully replaced by our 3 label lines,
  // so no blank/default paragraph is left before the label content.
  setChildren(tc, [tcPr, ...paragraphs]);
}

/** Forces every row's trPr to declare <w:cantSplit/> and, when a height is known, w:hRule="exact". */
function enforceRowProperties(tr: XmlNode): void {
  const trChildren = childrenOf(tr);
  const trPr = findDirectChildren(trChildren, "w:trPr")[0];
  const otherChildren = trChildren.filter((c) => tagNameOf(c) !== "w:trPr");

  const existing = trPr ? childrenOf(trPr) : [];
  const heightNode = existing.find((c) => tagNameOf(c) === "w:trHeight");
  const passthrough = existing.filter(
    (c) => tagNameOf(c) !== "w:trHeight" && tagNameOf(c) !== "w:cantSplit",
  );

  // CT_TrPrBase schema order (subset we care about): cantSplit, trHeight, ...
  const newTrPrChildren: XmlNode[] = [makeEmptyElement("w:cantSplit", {})];
  if (heightNode) {
    const val = getAttr(heightNode, "w:val");
    newTrPrChildren.push(
      val !== undefined
        ? makeEmptyElement("w:trHeight", { "w:val": val, "w:hRule": "exact" })
        : heightNode,
    );
  }
  newTrPrChildren.push(...passthrough);

  setChildren(tr, [makeElement("w:trPr", {}, newTrPrChildren), ...otherChildren]);
}

function makePageBreakParagraph(): XmlNode {
  return makeElement("w:p", {}, [
    makeElement("w:r", {}, [makeEmptyElement("w:br", { "w:type": "page" })]),
  ]);
}

/**
 * Renders a real, valid .docx for a validated FIXED_GRID template (e.g.
 * Avery 5155) from a deterministic placement plan.
 *
 * - Starts from a byte-for-byte copy of the source template (JSZip), and
 *   only ever rewrites word/document.xml - every other package part
 *   (styles, fonts, theme, settings, relationships, media) is preserved
 *   untouched.
 * - Clones the single validated grid table once per sheet; blank slots are
 *   left exactly as they were in the source (no placeholder text).
 * - Inserts an explicit page break paragraph only between completed sheet
 *   tables (never after the last one).
 * - Preserves the source's original section properties (page size/margins)
 *   verbatim.
 *
 * Throws InvalidFixedGridTemplateError if the template does not actually
 * validate as a `template.columns x template.rows = template.labelsPerSheet`
 * fixed grid, and MissingLabelTextStyleConfigError if
 * template.configJson.labelTextStyle is absent/invalid.
 */
export async function renderFixedGridDocx(
  templateBuffer: Buffer,
  template: LabelTemplate,
  plan: PlacementPlan,
): Promise<RenderResult> {
  if (template.renderingMode !== "FIXED_GRID") {
    throw new InvalidFixedGridTemplateError(template.id, [
      `Template renderingMode is "${template.renderingMode}", not FIXED_GRID.`,
    ]);
  }

  const style = readLabelTextStyle(template);

  const templatePackage = await loadTemplateDocx(templateBuffer, template.id);
  // Captured before any mutation: renderFixedGridDocx only ever overwrites
  // the existing word/document.xml entry's content, never adds or removes
  // a package part, so the final output must have exactly this same set.
  const expectedPackageEntryNames = new Set(Object.keys(templatePackage.zip.files));
  const validated = validateFixedGridTemplate(templatePackage, {
    columns: template.columns,
    rows: template.rows,
    labelsPerSheet: template.labelsPerSheet,
  });

  const sourceTable = findTableNode(templatePackage.documentTree, validated.tableIndex);

  const body = findFirstByTag(templatePackage.documentTree, "w:body");
  if (!body) {
    throw new InvalidFixedGridTemplateError(template.id, ["Document has no <w:body>."]);
  }
  const sectPr = findDirectChildren(childrenOf(body), "w:sectPr")[0];
  if (!sectPr) {
    throw new InvalidFixedGridTemplateError(template.id, [
      "Document has no body-level <w:sectPr>; cannot preserve original page geometry.",
    ]);
  }

  const placementsBySheet = new Map<number, Map<number, Placement>>();
  for (const placement of plan.placements) {
    let bySlot = placementsBySheet.get(placement.sheetNumber);
    if (!bySlot) {
      bySlot = new Map();
      placementsBySheet.set(placement.sheetNumber, bySlot);
    }
    bySlot.set(placement.slotIndex, placement);
  }

  // Reverse lookup: physical (row, cell) -> logical slot index. Only cells
  // present in this map are writable label cells; every other physical cell
  // (spacer/gutter columns in the INTERLEAVED_SPACER pattern) is never
  // touched - fillCell() is never called on it, so its XML (including any
  // <w:vMerge>) is left exactly as cloned from the source.
  const physicalCellToSlot = new Map<string, number>();
  for (const mapping of validated.pattern.writableCellMap) {
    physicalCellToSlot.set(`${mapping.physicalRowIndex}:${mapping.physicalCellIndex}`, mapping.logicalSlotIndex);
  }

  const sheetTables: XmlNode[] = [];
  for (let sheetNumber = 1; sheetNumber <= plan.totalSheets; sheetNumber++) {
    const clonedTable = cloneNode(sourceTable);
    const rows = findDirectChildren(childrenOf(clonedTable), "w:tr");
    const sheetPlacements = placementsBySheet.get(sheetNumber);

    rows.forEach((row, rowIndex) => {
      const cells = findDirectChildren(childrenOf(row), "w:tc");
      cells.forEach((cell, columnIndex) => {
        const slotIndex = physicalCellToSlot.get(`${rowIndex}:${columnIndex}`);
        if (slotIndex === undefined) {
          // Not a writable label cell (e.g. a spacer/gutter column) - leave
          // it completely untouched.
          return;
        }
        const placement = sheetPlacements?.get(slotIndex);
        if (placement) {
          const priceText =
            placement.priceCents === null
              ? ""
              : `As Low As: ${formatCentsAsDollars(placement.priceCents)}`;
          fillCell(
            cell,
            {
              sku: placement.sku ?? "",
              description: placement.description ?? "",
              priceLine: priceText,
            },
            style,
          );
        }
        // No placement for this slot: leave the cloned (blank) cell untouched.
      });
      enforceRowProperties(row);
    });

    sheetTables.push(clonedTable);
  }

  const newBodyChildren: XmlNode[] = [];
  sheetTables.forEach((table, index) => {
    newBodyChildren.push(table);
    if (index < sheetTables.length - 1) {
      newBodyChildren.push(makePageBreakParagraph());
    }
  });
  newBodyChildren.push(sectPr);

  setChildren(body, newBodyChildren);

  // Preserve the source template's own word/document.xml timestamp rather
  // than letting JSZip stamp the current wall-clock time on this entry, and
  // suppress JSZip's automatic parent-folder synthesis (createFolders:
  // false) since real .docx packages store flat file entries with no
  // explicit "word/" directory entry - without both of these, two renders
  // of byte-identical content (same template + same placement plan)
  // produce different output bytes (a fresh "word/" folder entry stamped
  // with the current wall-clock time gets added on every call), which
  // undermines reproducibility (a stable sha256 for identical inputs) and
  // silently adds a zip entry the source template never had. Every other
  // zip entry is untouched and keeps its original date.
  const originalDocumentEntry = templatePackage.zip.file("word/document.xml");
  const documentEntryDate = originalDocumentEntry?.date ?? new Date(0);

  const newDocumentXml = buildXml(templatePackage.documentTree);
  templatePackage.zip.file("word/document.xml", newDocumentXml, {
    date: documentEntryDate,
    createFolders: false,
  });

  const buffer = await templatePackage.zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
  });

  // Validate the actual output bytes - not the in-memory node tree, and
  // not a lenient re-parse - before ever returning them. This is what
  // catches a real-Word-rejects-the-file class of bug (e.g. the
  // duplicate-XML-declaration defect buildXml() used to produce) that a
  // structurally-valid-zip / re-parses-fine check does not.
  const outputZip = await JSZip.loadAsync(buffer);
  await assertGeneratedDocxPackageIsValid(outputZip, expectedPackageEntryNames);

  return { buffer, mimeType: DOCX_MIME_TYPE, fileExtension: "docx" };
}

/**
 * Floating/tag-style DOCX generation (e.g. Avery 22802) is intentionally
 * out of scope for this milestone. This always throws
 * UnsupportedFloatingTemplateError - it never produces partial or malformed
 * output for a floating template.
 */
export async function renderFloatingDocx(
  _templateBuffer: Buffer,
  template: LabelTemplate,
  _plan: PlacementPlan,
): Promise<RenderResult> {
  if (template.renderingMode !== "FLOATING") {
    throw new DocxNotImplementedError("renderFloatingDocx");
  }
  throw new UnsupportedFloatingTemplateError(template.id);
}
