import type { SourceDocumentType } from "@label-maker/shared";

export class UnsupportedFileTypeError extends Error {
  readonly code = "UNSUPPORTED_FILE_TYPE";

  constructor(message: string) {
    super(message);
    this.name = "UnsupportedFileTypeError";
  }
}

const EXTENSION_MAP: Record<string, SourceDocumentType> = {
  pdf: "PDF",
  csv: "CSV",
  xlsx: "XLSX",
  xls: "XLS",
};

const MIME_MAP: Record<string, SourceDocumentType> = {
  "application/pdf": "PDF",
  "text/csv": "CSV",
  "application/csv": "CSV",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "XLSX",
  "application/vnd.ms-excel": "XLS",
};

// Magic byte signatures used as a sanity check against the claimed extension/mime.
const PDF_MAGIC = Buffer.from("%PDF");
const ZIP_MAGIC = Buffer.from([0x50, 0x4b, 0x03, 0x04]); // XLSX is a zip container

function extensionFromFilename(filename: string): string | undefined {
  const match = /\.([a-zA-Z0-9]+)$/.exec(filename);
  return match?.[1]?.toLowerCase();
}

function matchesMagicBytes(buffer: Buffer, type: SourceDocumentType): boolean {
  switch (type) {
    case "PDF":
      return buffer.subarray(0, PDF_MAGIC.length).equals(PDF_MAGIC);
    case "XLSX":
      return buffer.subarray(0, ZIP_MAGIC.length).equals(ZIP_MAGIC);
    case "XLS":
      // Legacy .xls exports are sometimes actually CSV/HTML with an .xls
      // extension (a common spreadsheet-export quirk), so OLE2 magic bytes
      // are informative but not required; SheetJS handles both transparently.
      return true;
    case "CSV":
      // CSV has no reliable magic bytes; anything not matching a binary
      // signature above is presumed textual.
      return true;
  }
}

/**
 * Detects the SourceDocumentType from uploaded bytes plus filename/mime hints.
 * Extension is the primary signal (mime types from multipart uploads are
 * often generic or missing); magic bytes are used as a sanity check for the
 * binary formats we can verify (PDF, XLSX).
 */
export function detectFileType(input: {
  buffer: Buffer;
  originalFilename: string;
  mimeType?: string;
}): SourceDocumentType {
  const extension = extensionFromFilename(input.originalFilename);
  const fromExtension = extension ? EXTENSION_MAP[extension] : undefined;
  const fromMime = input.mimeType ? MIME_MAP[input.mimeType] : undefined;

  const detected = fromExtension ?? fromMime;
  if (!detected) {
    throw new UnsupportedFileTypeError(
      `Unsupported file type for "${input.originalFilename}" (mime: ${input.mimeType ?? "unknown"}). Supported: .pdf, .csv, .xlsx, .xls`,
    );
  }

  if (detected !== "CSV" && !matchesMagicBytes(input.buffer, detected)) {
    throw new UnsupportedFileTypeError(
      `File "${input.originalFilename}" has extension/mime indicating ${detected} but its contents do not match the expected file signature.`,
    );
  }

  return detected;
}
