import type { PrismaClient } from "@label-maker/database";
import { detectFileType, UnsupportedFileTypeError } from "@label-maker/ingestion";
import type { LocalStorageService } from "@label-maker/storage";
import { QUEUE_NAMES, type SourceDocumentType } from "@label-maker/shared";
import type { Queue } from "bullmq";
import { PayloadTooLargeAppError, UnsupportedMediaTypeAppError } from "../lib/errors.js";

export interface UploadInput {
  buffer: Buffer;
  originalFilename: string;
  mimeType: string;
  maxUploadBytes: number;
}

export interface UploadResult {
  documentId: string;
  jobId: string;
}

const EXTENSION_BY_TYPE: Record<SourceDocumentType, string> = {
  PDF: "pdf",
  CSV: "csv",
  XLSX: "xlsx",
  XLS: "xls",
};

export async function handleUpload(
  prisma: PrismaClient,
  storage: LocalStorageService,
  ingestDocumentQueue: Queue,
  input: UploadInput,
): Promise<UploadResult> {
  if (input.buffer.byteLength > input.maxUploadBytes) {
    throw new PayloadTooLargeAppError(
      `Upload exceeds maximum allowed size of ${input.maxUploadBytes} bytes.`,
    );
  }

  let documentType: SourceDocumentType;
  try {
    documentType = detectFileType({
      buffer: input.buffer,
      originalFilename: input.originalFilename,
      mimeType: input.mimeType,
    });
  } catch (error) {
    if (error instanceof UnsupportedFileTypeError) {
      throw new UnsupportedMediaTypeAppError(error.message);
    }
    throw error;
  }

  const saved = await storage.save("uploads", input.buffer, EXTENSION_BY_TYPE[documentType]);

  const sourceDocument = await prisma.sourceDocument.create({
    data: {
      originalFilename: input.originalFilename,
      mimeType: input.mimeType,
      extension: documentType,
      storageKey: saved.storageKey,
      byteSize: saved.byteSize,
      sha256: saved.sha256,
      pageCount: null,
    },
  });

  const processingJob = await prisma.processingJob.create({
    data: {
      sourceDocumentId: sourceDocument.id,
      jobType: "INGEST_DOCUMENT",
      status: "QUEUED",
    },
  });

  await ingestDocumentQueue.add(
    QUEUE_NAMES.INGEST_DOCUMENT,
    { processingJobId: processingJob.id, sourceDocumentId: sourceDocument.id },
    // BullMQ custom job IDs can't contain ":"; using the document id alone
    // as the idempotency key is enough (INGEST_DOCUMENT is the only job
    // type on this queue).
    { jobId: `ingest-document-${sourceDocument.id}` },
  );

  return { documentId: sourceDocument.id, jobId: processingJob.id };
}
