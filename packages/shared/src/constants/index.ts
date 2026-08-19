export const SOURCE_DOCUMENT_TYPES = ["PDF", "CSV", "XLSX", "XLS"] as const;

export const PROCESSING_JOB_TYPES = ["INGEST_DOCUMENT", "GENERATE_LABEL_RUN"] as const;

export const PROCESSING_JOB_STATUSES = [
  "QUEUED",
  "PROCESSING",
  "COMPLETED",
  "FAILED",
  "NEEDS_REVIEW",
] as const;

export const EXTRACTION_METHODS = ["NATIVE_TEXT", "OCR", "SPREADSHEET", "MANUAL"] as const;

export const EXTRACTION_FIELDS = ["SKU", "DESCRIPTION", "PRICE", "INCLUDE_STATUS"] as const;

export const PRODUCT_STATUSES = ["AUTO_ACCEPTED", "NEEDS_REVIEW", "APPROVED", "EXCLUDED"] as const;

export const LABEL_TEMPLATE_RENDERING_MODES = ["FIXED_GRID", "FLOATING"] as const;

export const LABEL_RUN_STATUSES = ["DRAFT", "READY", "GENERATED", "FAILED"] as const;

/** Default maximum accepted upload size in bytes (50 MB). */
export const DEFAULT_MAX_UPLOAD_BYTES = 50 * 1024 * 1024;

/** Storage subdirectories used by local disk storage in development. */
export const STORAGE_UPLOADS_SUBDIR = "uploads";
export const STORAGE_ARTIFACTS_SUBDIR = "artifacts";

/** BullMQ queue names shared between apps/api (producer) and apps/worker (consumer). */
export const QUEUE_NAMES = {
  INGEST_DOCUMENT: "ingest-document",
  GENERATE_LABEL_RUN: "generate-label-run",
} as const;
