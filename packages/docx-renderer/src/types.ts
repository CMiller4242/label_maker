import { z } from "zod";
import type { LabelTemplate, PlacementPlan } from "@label-maker/shared";
import type { TemplatePackage } from "./ooxml.js";

export class DocxNotImplementedError extends Error {
  readonly code = "DOCX_NOT_IMPLEMENTED";

  constructor(operation: string) {
    super(
      `${operation} is not implemented yet. See packages/docx-renderer/README.md ` +
        `for the planned WordprocessingML strategy and current support status.`,
    );
    this.name = "DocxNotImplementedError";
  }
}

/** Thrown by renderFloatingDocx(): floating/tag-style DOCX generation (e.g. Avery 22802) is out of scope for this milestone. */
export class UnsupportedFloatingTemplateError extends Error {
  readonly code = "UNSUPPORTED_FLOATING_TEMPLATE";

  constructor(templateId: string) {
    super(
      `Template "${templateId}" uses FLOATING rendering mode. Floating/tag-style DOCX ` +
        `generation is intentionally unsupported in this milestone; only FIXED_GRID ` +
        `templates (e.g. Avery 5155) can be rendered to a real DOCX artifact today.`,
    );
    this.name = "UnsupportedFloatingTemplateError";
  }
}

/** Thrown when a template's source .docx does not actually contain a valid fixed grid matching the expected geometry. */
export class InvalidFixedGridTemplateError extends Error {
  readonly code = "INVALID_FIXED_GRID_TEMPLATE";
  readonly issues: string[];

  constructor(templateLabel: string, issues: string[]) {
    super(
      `Template "${templateLabel}" failed fixed-grid validation:\n` +
        issues.map((issue) => `  - ${issue}`).join("\n"),
    );
    this.name = "InvalidFixedGridTemplateError";
    this.issues = issues;
  }
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
  /** Bytes of the produced artifact (a debug JSON document, or a real .docx once rendered). */
  buffer: Buffer;
  mimeType: string;
  fileExtension: string;
}

// --- Inspection report types ---------------------------------------------

export interface PageGeometry {
  widthTwips: number | null;
  heightTwips: number | null;
  marginTopTwips: number | null;
  marginBottomTwips: number | null;
  marginLeftTwips: number | null;
  marginRightTwips: number | null;
}

/** Vertical-merge state of one raw table cell, per OOXML `<w:vMerge>` semantics. */
export type CellVerticalMergeState = "restart" | "continue" | null;

export interface CellInspection {
  /** Raw cell index within its row, left to right (0-based). */
  cellIndex: number;
  widthTwips: number | null;
  vMerge: CellVerticalMergeState;
}

export interface RowInspection {
  rowIndex: number;
  cellCount: number;
  heightTwips: number | null;
  heightRule: string | null;
  cantSplit: boolean;
  /** Per-cell detail (width, vertical-merge state) for this row's raw cells. */
  cells: CellInspection[];
}

export interface CellMarginInspection {
  topTwips: number | null;
  bottomTwips: number | null;
  leftTwips: number | null;
  rightTwips: number | null;
}

export interface FontSample {
  fontFamily: string | null;
  sizeHalfPoints: number | null;
  bold: boolean;
  colorHex: string | null;
}

export interface ParagraphSample {
  spacingBeforeTwips: number | null;
  spacingAfterTwips: number | null;
  lineSpacing: string | null;
  lineSpacingRule: string | null;
  alignment: string | null;
}

export interface TablePositioning {
  /** true when the table carries <w:tblpPr> (floating/anchored positioning) rather than normal in-flow placement. */
  isFloating: boolean;
  anchorVertical: string | null;
  anchorHorizontal: string | null;
  positionXTwips: number | null;
  positionYTwips: number | null;
}

/**
 * One of the fixed-grid table shapes this package knows how to map logical
 * label slots onto physical `<w:tc>` cells for:
 *
 * - "SIMPLE": every raw grid column is a writable label column (e.g. the
 *   package's synthetic test fixture: 4 raw columns, 4 logical columns,
 *   no vertical merging).
 * - "INTERLEAVED_SPACER": raw grid columns alternate label/spacer/label/...
 *   /label (an odd count), with the narrow spacer columns vertically
 *   merged (`<w:vMerge>`) down the table for horizontal pitch - the real
 *   Avery 5155 template's shape (7 raw columns: 4 label + 3 spacer).
 */
export type FixedGridPatternType = "SIMPLE" | "INTERLEAVED_SPACER";

/** One logical label slot's location in the physical table. */
export interface WritableCellMapping {
  /** Row-major logical slot index: physicalRowIndex * logicalColumns + logicalColumnIndex. */
  logicalSlotIndex: number;
  physicalRowIndex: number;
  /** Index within the row's raw `<w:tc>` cells (0-based, left to right). */
  physicalCellIndex: number;
}

export interface FixedGridPattern {
  patternType: FixedGridPatternType;
  /** Raw `<w:tblGrid>` column widths (twips), left to right - preserved verbatim from the source. */
  rawGridColumnWidthsTwips: number[];
  /** Raw column indexes that hold writable label content. */
  logicalLabelColumnIndexes: number[];
  /** Raw column indexes that are spacer/gutter columns and must never be written to. */
  spacerColumnIndexes: number[];
  logicalColumns: number;
  logicalRows: number;
  /** Every logical label slot's physical (row, cell) location - the renderer's only source of truth for where to write. */
  writableCellMap: WritableCellMapping[];
}

export interface TableInspection {
  tableIndex: number;
  layoutType: string | null;
  isFixedLayout: boolean;
  tableWidthTwips: number | null;
  tableWidthType: string | null;
  gridColumnWidthsTwips: number[];
  rowCount: number;
  rows: RowInspection[];
  totalWritableCells: number;
  cellMargins: CellMarginInspection;
  cellVerticalAlignment: string | null;
  paragraphSample: ParagraphSample | null;
  fontSample: FontSample | null;
  positioning: TablePositioning;
  /** The recognized fixed-grid pattern, or null if this table doesn't match either known shape. */
  fixedGridPattern: FixedGridPattern | null;
  /**
   * When `fixedGridPattern` is null, explains exactly which condition(s)
   * were close-but-not-satisfied (e.g. "spacer columns aren't uniformly
   * narrower", "row 3 spacer column 1 has no <w:vMerge>") - a
   * near-miss should never be silently indistinguishable from "not a grid
   * at all".
   */
  fixedGridPatternDiagnostics: string[];
}

export type TemplateClassification =
  "AVERY_5155_LIKE_FIXED_GRID" | "AVERY_22802_LIKE_FLOATING" | "AMBIGUOUS";

export interface TableClassificationInput {
  tables: TableInspection[];
}

export interface DocxInspectionReport {
  filePath: string | null;
  sha256: string;
  byteSize: number;
  pageGeometry: PageGeometry;
  tableCount: number;
  tables: TableInspection[];
  classification: TemplateClassification;
  classificationConfidence: number;
  warnings: string[];
}

// --- Fixed-grid validation -------------------------------------------------

export interface ExpectedFixedGridGeometry {
  columns: number;
  rows: number;
  labelsPerSheet: number;
}

/** The specific table (and its structural shape) validated as usable for cloning by the fixed-grid renderer. */
export interface ValidatedFixedGridTemplate {
  templateStorageKey: string;
  tableIndex: number;
  columns: number;
  rows: number;
  labelsPerSheet: number;
  inspection: TableInspection;
  /** The recognized pattern (SIMPLE or INTERLEAVED_SPACER) - the renderer's only source of truth for which physical cells are writable. */
  pattern: FixedGridPattern;
}

// --- Label cell text style configuration -----------------------------------

export const horizontalAlignmentSchema = z.enum(["left", "center", "right"]);
export const verticalAlignmentSchema = z.enum(["top", "center", "bottom"]);

export const lineSpacingConfigSchema = z.object({
  /** w:spacing/@w:lineRule value: "auto" | "exact" | "atLeast". */
  lineRule: z.enum(["auto", "exact", "atLeast"]),
  /** w:spacing/@w:line value. For "auto", this is 240ths-of-a-line; for "exact"/"atLeast", twips. */
  line: z.number().int().positive(),
});

export const safeInsetsConfigSchema = z.object({
  topTwips: z.number().int().nonnegative(),
  bottomTwips: z.number().int().nonnegative(),
  leftTwips: z.number().int().nonnegative(),
  rightTwips: z.number().int().nonnegative(),
});

/**
 * Explicit, configurable style for the 3 lines of text rendered into each
 * filled label cell (SKU / description / price). Stored under
 * LabelTemplate.configJson.labelTextStyle. Deliberately not hardcoded in
 * renderer code - see fixed-grid-renderer.ts.
 */
export const labelTextStyleConfigSchema = z.object({
  fontFamily: z.string().min(1),
  skuFontSizeHalfPoints: z.number().int().positive(),
  descriptionFontSizeHalfPoints: z.number().int().positive(),
  priceFontSizeHalfPoints: z.number().int().positive(),
  bold: z.boolean(),
  colorHex: z.string().regex(/^[0-9A-Fa-f]{6}$/),
  horizontalAlignment: horizontalAlignmentSchema,
  verticalAlignment: verticalAlignmentSchema,
  lineSpacing: lineSpacingConfigSchema,
  paragraphSpacingBeforeTwips: z.number().int().nonnegative(),
  paragraphSpacingAfterTwips: z.number().int().nonnegative(),
  safeInsets: safeInsetsConfigSchema.optional(),
});
export type LabelTextStyleConfig = z.infer<typeof labelTextStyleConfigSchema>;

export class MissingLabelTextStyleConfigError extends Error {
  readonly code = "MISSING_LABEL_TEXT_STYLE_CONFIG";

  constructor(templateId: string, cause: string) {
    super(
      `Template "${templateId}" is missing a valid configJson.labelTextStyle - ` +
        `renderFixedGridDocx() requires explicit, configured text style rather than ` +
        `hardcoded defaults. Cause: ${cause}`,
    );
    this.name = "MissingLabelTextStyleConfigError";
  }
}

/**
 * The renderer contract every concrete implementation must satisfy. Keeping
 * this interface stable lets callers (apps/api, apps/worker) swap
 * implementations without changing call sites.
 */
export interface LabelRenderer {
  /** Opens a template package (.docx/.dotx) from storage. */
  loadTemplateDocx(templateStorageKey: string): Promise<TemplatePackage>;

  /** Checks a loaded template's table/row/column structure against the LabelTemplate's declared geometry. */
  validateTemplateLayout(
    templatePackage: TemplatePackage,
    template: LabelTemplate,
  ): TemplateLayoutValidationResult;

  /** Renders a FIXED_GRID template (e.g. Avery 5155) by cloning its validated fixed-grid table. */
  renderFixedGridDocx(plan: PlacementPlan, template: LabelTemplate): Promise<RenderResult>;

  /** Renders a FLOATING template (free-positioned tags/labels). Always throws UnsupportedFloatingTemplateError today. */
  renderFloatingDocx(plan: PlacementPlan, template: LabelTemplate): Promise<RenderResult>;
}
