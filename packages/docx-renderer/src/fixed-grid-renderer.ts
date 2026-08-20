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
    // Explicit zero indentation - defense in depth against any inherited
    // left/right/first-line/hanging indent (from styles.xml or a future
    // template) silently shifting content off-center. This label renderer
    // never wants paragraph indentation; centering comes entirely from
    // w:jc below plus the cell's own width/margins.
    makeEmptyElement("w:ind", { "w:left": "0", "w:right": "0", "w:firstLine": "0", "w:hanging": "0" }),
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

/**
 * CT_TcPrBase's element sequence (the subset this renderer ever touches).
 * `w:tcPr` children must appear in this relative order or Word's strict
 * OOXML parser can silently ignore/misplace an out-of-order element - so
 * every helper that adds/replaces a tcPr child must insert it at the
 * position this order implies, not just append it.
 */
const TC_PR_CHILD_ORDER = [
  "w:cnfStyle",
  "w:tcW",
  "w:gridSpan",
  "w:hMerge",
  "w:vMerge",
  "w:tcBorders",
  "w:shd",
  "w:noWrap",
  "w:tcMar",
  "w:textDirection",
  "w:tcFitText",
  "w:vAlign",
  "w:hideMark",
];

/** Replaces (or inserts) one `w:tcPr` child at its schema-correct position, per TC_PR_CHILD_ORDER. */
function setTcPrChildInOrder(tcPr: XmlNode, tag: string, newChild: XmlNode): void {
  const children = childrenOf(tcPr).filter((c) => tagNameOf(c) !== tag);
  const orderIndex = TC_PR_CHILD_ORDER.indexOf(tag);
  const insertAt = children.findIndex((c) => {
    const otherIndex = TC_PR_CHILD_ORDER.indexOf(tagNameOf(c) ?? "");
    return otherIndex === -1 || otherIndex > orderIndex;
  });
  if (insertAt === -1) children.push(newChild);
  else children.splice(insertAt, 0, newChild);
  setChildren(tcPr, children);
}

function ensureCellVerticalAlignment(
  tcPr: XmlNode,
  vAlign: LabelTextStyleConfig["verticalAlignment"],
): void {
  setTcPrChildInOrder(tcPr, "w:vAlign", makeEmptyElement("w:vAlign", { "w:val": vAlign }));
}

/** Thin single-line border color used for the review-only cell-outline mode - never applied by default. */
const REVIEW_BORDER_COLOR = "BFBFBF";

/**
 * Adds a thin, uniform border to every side of a cell - purely to make the
 * real physical cell rectangle visible on screen for centering/geometry
 * review. This is opt-in only (`renderFixedGridDocx`'s `reviewOutlines`
 * option) and MUST NOT be used for a template that will actually be
 * printed on real Avery stock: unlike Word's built-in non-printing "Table
 * Gridlines" view (View > Gridlines, which already shows for any
 * borderless table - true for every cell here by default, on screen only,
 * never printed), a `w:tcBorders` border is real cell-boundary ink that
 * Word WILL print.
 */
function ensureCellReviewBorder(tcPr: XmlNode): void {
  const edge = (tag: string) =>
    makeEmptyElement(tag, {
      "w:val": "single",
      "w:sz": "2",
      "w:space": "0",
      "w:color": REVIEW_BORDER_COLOR,
    });
  const tcBorders = makeElement("w:tcBorders", {}, [
    edge("w:top"),
    edge("w:left"),
    edge("w:bottom"),
    edge("w:right"),
  ]);
  setTcPrChildInOrder(tcPr, "w:tcBorders", tcBorders);
}

/**
 * Sets the table's overall preferred width (`w:tblW`) to the exact sum of
 * its own already-declared `w:tblGrid` column widths, as an explicit `dxa`
 * value. Never invents a new physical dimension - `widthsTwips` always
 * comes from the table's own inspected/validated grid.
 *
 * Why this matters: `w:tblLayout w:type="fixed"` alone does not reliably
 * force Word to honor per-column widths when `w:tblW` is `type="auto"`
 * (`w:w="0"`, i.e. "let Word calculate") - a well-documented Word
 * compatibility hazard where the table's effective width/column
 * allocation can still be recomputed from available page width rather
 * than the declared grid, especially once real content fills more cells
 * than the template's blank preview state exercised. Pairing "fixed"
 * layout with an explicit total `w:tblW` removes that ambiguity.
 */
function normalizeTableWidth(tblPr: XmlNode, widthsTwips: number[]): void {
  const totalWidth = widthsTwips.reduce((sum, w) => sum + w, 0);
  const children = childrenOf(tblPr).filter((c) => tagNameOf(c) !== "w:tblW");
  const tblW = makeEmptyElement("w:tblW", { "w:w": String(totalWidth), "w:type": "dxa" });
  // CT_TblPrBase: tblW is the first sizing-related element, ordered before
  // tblLayout/tblCellMar/tblLook/etc. - insert before whichever of those
  // comes first (or at the front if tblPr had no such children yet).
  const insertAt = children.findIndex((c) =>
    ["w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblBorders", "w:jc", "w:tblInd"].includes(
      tagNameOf(c) ?? "",
    ),
  );
  if (insertAt === -1) children.unshift(tblW);
  else children.splice(insertAt, 0, tblW);
  setChildren(tblPr, children);
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
export interface RenderFixedGridDocxOptions {
  /**
   * Adds a thin, non-default border to every cell (writable and
   * spacer/gutter alike) so the real physical cell rectangles are visible
   * on screen for centering/geometry review. NEVER enable this for a
   * template that will actually be printed on real Avery stock - unlike
   * Word's built-in non-printing "Table Gridlines" view (View > Gridlines,
   * on by default for any borderless table, including this renderer's
   * normal output), these borders are real ink Word WILL print. Defaults
   * to false/unset - normal rendering never adds borders.
   */
  reviewOutlines?: boolean;
}

export async function renderFixedGridDocx(
  templateBuffer: Buffer,
  template: LabelTemplate,
  plan: PlacementPlan,
  options: RenderFixedGridDocxOptions = {},
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

  // Fix a well-documented Word compatibility hazard once, before cloning:
  // an explicit total table width removes any ambiguity between the
  // declared "fixed" layout and an "auto" w:tblW, so every cloned sheet
  // inherits the correction automatically. See normalizeTableWidth()'s
  // doc comment.
  const tableTblPr = findDirectChildren(childrenOf(sourceTable), "w:tblPr")[0];
  if (tableTblPr) {
    normalizeTableWidth(tableTblPr, validated.pattern.rawGridColumnWidthsTwips);
  }

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
        if (options.reviewOutlines) {
          // Every cell, writable or spacer/gutter, blank or filled - the
          // point is to make the real physical grid visible for review.
          const cellChildren = childrenOf(cell);
          let cellTcPr = findDirectChildren(cellChildren, "w:tcPr")[0];
          if (!cellTcPr) {
            cellTcPr = makeElement("w:tcPr", {}, []);
            setChildren(cell, [cellTcPr, ...cellChildren]);
          }
          ensureCellReviewBorder(cellTcPr);
        }

        const slotIndex = physicalCellToSlot.get(`${rowIndex}:${columnIndex}`);
        if (slotIndex === undefined) {
          // Not a writable label cell (e.g. a spacer/gutter column) - leave
          // it completely untouched (beyond the review border above, when
          // requested).
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
