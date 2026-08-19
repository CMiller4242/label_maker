import * as XLSX from "xlsx";
import type { SpreadsheetMapping } from "@label-maker/shared";
import { CONFIDENCE } from "./confidence.js";
import { normalizeHeaders, type HeaderMapping } from "./normalize-headers.js";
import { normalizePriceToCents } from "./normalize-price.js";
import type { IngestionCandidate, IngestionProduct, IngestionResult } from "./types.js";

export class WorkbookParseError extends Error {
  readonly code = "WORKBOOK_PARSE_ERROR";
}

const MAX_PREVIEW_ROWS = 20;

export interface SheetPreview {
  sheetName: string;
  range: string;
  headers: string[];
  headerMapping: HeaderMapping;
  /** First N non-empty rows, formatted cell values, keyed by header. */
  previewRows: Record<string, string>[];
  isViable: boolean;
}

export interface WorkbookIngestionResult extends IngestionResult {
  sheetPreviews: SheetPreview[];
}

function sheetToRows(sheet: XLSX.WorkSheet): string[][] {
  // raw: false -> formatted display strings (so "$14.49" style formatting and
  // decimal precision are preserved instead of routing prices through raw
  // JS floats), defval: "" -> keep sparse rows aligned by column.
  const rows: unknown[][] = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: "",
    blankrows: false,
  });
  return rows.map((row) => row.map((cell) => String(cell ?? "").trim()));
}

function buildSheetPreview(sheetName: string, sheet: XLSX.WorkSheet): SheetPreview {
  const range = sheet["!ref"] ?? "A1";
  const rows = sheetToRows(sheet);
  const headerRow = rows[0] ?? [];
  const headerMapping = normalizeHeaders(headerRow);

  const dataRows = rows.slice(1, 1 + MAX_PREVIEW_ROWS);
  const previewRows: Record<string, string>[] = dataRows.map((row) => {
    const record: Record<string, string> = {};
    headerRow.forEach((header, columnIndex) => {
      record[header] = row[columnIndex] ?? "";
    });
    return record;
  });

  const isViable =
    headerMapping.ambiguousFields.length === 0 &&
    headerMapping.matches.SKU.length === 1 &&
    headerMapping.matches.DESCRIPTION.length === 1 &&
    headerMapping.matches.PRICE.length === 1 &&
    rows.length > 1;

  return {
    sheetName,
    range,
    headers: headerRow,
    headerMapping,
    previewRows,
    isViable,
  };
}

function buildProductsFromSheet(
  sheet: XLSX.WorkSheet,
  mapping: {
    skuHeader: string;
    descriptionHeader: string;
    priceHeader: string;
    includeHeader?: string;
  },
  warnings: string[],
): { products: IngestionProduct[]; anyNeedsReview: boolean } {
  const rows = sheetToRows(sheet);
  const headerRow = rows[0] ?? [];
  const dataRows = rows.slice(1);

  const columnIndex = (header: string): number => headerRow.indexOf(header);
  const skuIdx = columnIndex(mapping.skuHeader);
  const descIdx = columnIndex(mapping.descriptionHeader);
  const priceIdx = columnIndex(mapping.priceHeader);
  const includeIdx = mapping.includeHeader ? columnIndex(mapping.includeHeader) : -1;

  let anyNeedsReview = false;

  const products: IngestionProduct[] = dataRows.map((row, index) => {
    const sourceRowNumber = index + 2;
    const rawSku = (row[skuIdx] ?? "").trim();
    const rawDescription = (row[descIdx] ?? "").trim();
    const rawPrice = (row[priceIdx] ?? "").trim();
    const priceResult = normalizePriceToCents(rawPrice);

    if (priceResult.malformed) {
      warnings.push(`Row ${sourceRowNumber}: malformed price value "${rawPrice}".`);
    }

    const candidates: IngestionCandidate[] = [
      {
        field: "SKU",
        rawValue: rawSku,
        confidence: rawSku.length > 0 ? CONFIDENCE.HIGH : CONFIDENCE.LOW,
        ruleName: "xlsx-header-mapping",
        ...(rawSku.length > 0 ? { normalizedValue: rawSku } : {}),
      },
      {
        field: "DESCRIPTION",
        rawValue: rawDescription,
        confidence: rawDescription.length > 0 ? CONFIDENCE.HIGH : CONFIDENCE.LOW,
        ruleName: "xlsx-header-mapping",
        ...(rawDescription.length > 0 ? { normalizedValue: rawDescription } : {}),
      },
      {
        field: "PRICE",
        rawValue: rawPrice,
        confidence:
          rawPrice.length === 0
            ? CONFIDENCE.LOW
            : priceResult.malformed
              ? CONFIDENCE.LOW
              : CONFIDENCE.HIGH,
        ruleName: "xlsx-price-normalization",
        ...(priceResult.priceCents !== null
          ? { normalizedValue: String(priceResult.priceCents) }
          : {}),
      },
    ];

    let include = true;
    if (includeIdx >= 0) {
      const rawInclude = (row[includeIdx] ?? "").trim().toLowerCase();
      if (rawInclude.length > 0) {
        include = !["false", "no", "0", "inactive", "excluded", "n"].includes(rawInclude);
        candidates.push({
          field: "INCLUDE_STATUS",
          rawValue: rawInclude,
          normalizedValue: String(include),
          confidence: CONFIDENCE.HIGH,
          ruleName: "xlsx-include-status",
        });
      }
    }

    const rowHasBlank = rawSku.length === 0 || rawDescription.length === 0 || rawPrice.length === 0;
    if (rowHasBlank || priceResult.malformed) anyNeedsReview = true;

    return {
      sourceRowNumber,
      sku: rawSku.length > 0 ? rawSku : null,
      description: rawDescription.length > 0 ? rawDescription : null,
      priceCents: priceResult.priceCents,
      include,
      confidence: Math.min(...candidates.map((c) => c.confidence)),
      candidates,
    };
  });

  return { products, anyNeedsReview };
}

/**
 * Parses an XLSX/XLS buffer. Never auto-combines worksheets:
 * - If exactly one sheet is "viable" (non-empty, unambiguous SKU/description/
 *   price headers), it is parsed into candidate products.
 * - Otherwise (0 or 2+ viable sheets, or an explicit mapping wasn't given
 *   for an ambiguous workbook) the result is returned with needsReview=true
 *   and no products; callers should present sheetPreviews for the user to
 *   pick from and resubmit with an explicit SpreadsheetMapping.
 */
export function parseWorkbook(
  buffer: Buffer,
  explicitMapping?: SpreadsheetMapping,
): WorkbookIngestionResult {
  let workbook: XLSX.WorkBook;
  try {
    workbook = XLSX.read(buffer, { type: "buffer", cellDates: false });
  } catch (error) {
    throw new WorkbookParseError(
      `Failed to parse workbook: ${error instanceof Error ? error.message : String(error)}`,
    );
  }

  const warnings: string[] = [];
  const nonEmptySheetNames = workbook.SheetNames.filter((name) => {
    const sheet = workbook.Sheets[name];
    return sheet !== undefined && sheet["!ref"] !== undefined;
  });

  const sheetPreviews: SheetPreview[] = nonEmptySheetNames.map((name) => {
    const sheet = workbook.Sheets[name];
    if (!sheet) {
      throw new WorkbookParseError(`Sheet "${name}" was listed but not found in workbook.`);
    }
    return buildSheetPreview(name, sheet);
  });

  if (explicitMapping) {
    const sheet = workbook.Sheets[explicitMapping.sheetName];
    if (!sheet) {
      throw new WorkbookParseError(`Sheet "${explicitMapping.sheetName}" not found in workbook.`);
    }
    const { products, anyNeedsReview } = buildProductsFromSheet(
      sheet,
      {
        skuHeader: explicitMapping.skuColumn,
        descriptionHeader: explicitMapping.descriptionColumn,
        priceHeader: explicitMapping.priceColumn,
        ...(explicitMapping.includeColumn ? { includeHeader: explicitMapping.includeColumn } : {}),
      },
      warnings,
    );
    return {
      documentType: "XLSX",
      sheetNames: workbook.SheetNames,
      products,
      warnings,
      needsReview: anyNeedsReview,
      sheetPreviews,
    };
  }

  const viableSheets = sheetPreviews.filter((s) => s.isViable);

  if (viableSheets.length === 1) {
    const viable = viableSheets[0];
    if (!viable) {
      throw new WorkbookParseError("Internal error selecting the viable sheet.");
    }
    const sheet = workbook.Sheets[viable.sheetName];
    if (!sheet) {
      throw new WorkbookParseError(`Sheet "${viable.sheetName}" not found in workbook.`);
    }
    const skuHeader = viable.headerMapping.matches.SKU[0];
    const descriptionHeader = viable.headerMapping.matches.DESCRIPTION[0];
    const priceHeader = viable.headerMapping.matches.PRICE[0];
    const includeHeader = viable.headerMapping.matches.INCLUDE_STATUS[0];
    if (!skuHeader || !descriptionHeader || !priceHeader) {
      throw new WorkbookParseError("Viable sheet unexpectedly missing required header mapping.");
    }

    const { products, anyNeedsReview } = buildProductsFromSheet(
      sheet,
      { skuHeader, descriptionHeader, priceHeader, ...(includeHeader ? { includeHeader } : {}) },
      warnings,
    );

    return {
      documentType: "XLSX",
      sheetNames: workbook.SheetNames,
      products,
      warnings,
      needsReview: anyNeedsReview,
      sheetPreviews,
    };
  }

  if (viableSheets.length === 0) {
    warnings.push(
      "No sheet has unambiguous SKU/description/price headers. Manual mapping required.",
    );
  } else {
    warnings.push(
      `Multiple viable sheets found (${viableSheets.map((s) => s.sheetName).join(", ")}). Refusing to guess which one to use; manual sheet selection required.`,
    );
  }

  return {
    documentType: "XLSX",
    sheetNames: workbook.SheetNames,
    products: [],
    warnings,
    needsReview: true,
    sheetPreviews,
  };
}
