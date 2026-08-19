import type {
  IngestionResult,
  PdfIngestionResult,
  WorkbookIngestionResult,
} from "@label-maker/ingestion";
import type { SourceDocumentType } from "@label-maker/shared";
import { extractPdf } from "../processors/extract-pdf.js";
import { extractSpreadsheet } from "../processors/extract-spreadsheet.js";

export type DocumentIngestionResult =
  | { kind: "pdf"; result: PdfIngestionResult }
  | { kind: "spreadsheet"; result: IngestionResult | WorkbookIngestionResult };

/** Routes a source document to the correct extraction processor by its detected type. */
export async function routeDocumentForExtraction(
  buffer: Buffer,
  documentType: SourceDocumentType,
): Promise<DocumentIngestionResult> {
  if (documentType === "PDF") {
    return { kind: "pdf", result: await extractPdf(buffer) };
  }
  return { kind: "spreadsheet", result: extractSpreadsheet(buffer, documentType) };
}
