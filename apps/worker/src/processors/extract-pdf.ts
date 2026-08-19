import { parsePdf, type PdfIngestionResult } from "@label-maker/ingestion";

/** Thin wrapper over the ingestion package's native PDF text extraction. */
export async function extractPdf(buffer: Buffer): Promise<PdfIngestionResult> {
  return parsePdf(buffer);
}
