import {
  parseCsv,
  parseWorkbook,
  type IngestionResult,
  type WorkbookIngestionResult,
} from "@label-maker/ingestion";
import type { SourceDocumentType } from "@label-maker/shared";

/**
 * Dispatches CSV vs XLSX/XLS parsing. CSV is header-row tabular data with no
 * concept of multiple sheets, so it always returns a plain IngestionResult;
 * XLSX/XLS may return sheetPreviews for a future sheet-selection UI when the
 * workbook is ambiguous (see parseWorkbook in @label-maker/ingestion).
 */
export function extractSpreadsheet(
  buffer: Buffer,
  documentType: Extract<SourceDocumentType, "CSV" | "XLSX" | "XLS">,
): IngestionResult | WorkbookIngestionResult {
  if (documentType === "CSV") {
    return parseCsv(buffer);
  }
  return parseWorkbook(buffer);
}
