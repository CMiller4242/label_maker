import { useId, useRef, type ChangeEvent } from "react";
import { useIngestUpload } from "../hooks/useIngestUpload";
import { ErrorBanner } from "./ErrorBanner";
import { LoadingIndicator } from "./LoadingIndicator";
import "./UploadStage.css";

const ACCEPTED_EXTENSIONS = [".csv", ".xlsx", ".xls", ".pdf"];

export interface UploadStageProps {
  onDocumentReady: (documentId: string) => void;
}

function fileTypeLabel(file: File): string {
  const extension = file.name.split(".").pop()?.toUpperCase();
  return extension ?? (file.type || "Unknown");
}

/**
 * Stage 1: upload/import. Calls the existing POST /uploads + GET /jobs/:id
 * endpoints only - no new parsing backend. Supports CSV/XLSX/XLS/PDF, the
 * same set apps/api's detectFileType() already accepts.
 */
export function UploadStage({ onDocumentReady }: UploadStageProps) {
  const { status, file, documentId, job, errorMessage, upload, retry, reset } = useIngestUpload();
  const inputRef = useRef<HTMLInputElement>(null);
  const inputId = useId();

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const selected = event.target.files?.[0];
    if (selected) upload(selected);
  };

  const handleChooseDifferent = () => {
    reset();
    if (inputRef.current) inputRef.current.value = "";
    inputRef.current?.focus();
  };

  return (
    <section aria-labelledby="upload-stage-heading" className="upload-stage">
      <h2 id="upload-stage-heading">1. Upload a product file</h2>
      <p className="upload-stage__hint">
        Supported file types: CSV, XLSX, XLS, PDF. Your file is parsed by the existing ingestion
        pipeline - no data is generated or invented here.
      </p>

      <div className="upload-stage__field">
        <label htmlFor={inputId}>Choose a file to upload</label>
        <input
          ref={inputRef}
          id={inputId}
          type="file"
          accept={ACCEPTED_EXTENSIONS.join(",")}
          onChange={handleFileChange}
          disabled={status === "uploading" || status === "processing"}
        />
      </div>

      {file && (
        <dl className="upload-stage__file-summary">
          <div>
            <dt>Selected file</dt>
            <dd>{file.name}</dd>
          </div>
          <div>
            <dt>Type</dt>
            <dd>{fileTypeLabel(file)}</dd>
          </div>
        </dl>
      )}

      {status === "uploading" && <LoadingIndicator label="Uploading file…" />}
      {status === "processing" && <LoadingIndicator label="Processing document…" />}

      {status === "failed" && (
        <ErrorBanner message={errorMessage ?? "Upload failed."} onRetry={file ? retry : undefined} />
      )}

      {status === "succeeded" && job && documentId && (
        <div className="upload-stage__success" role="status">
          <p>
            <strong>{file?.name}</strong> processed successfully
            {job.status === "NEEDS_REVIEW" ? " - some rows need review." : "."}
          </p>
          <div className="upload-stage__actions">
            <button type="button" onClick={handleChooseDifferent} className="button button--secondary">
              Choose a different file
            </button>
            <button
              type="button"
              className="button button--primary"
              onClick={() => onDocumentReady(documentId)}
            >
              Continue to product review
            </button>
          </div>
        </div>
      )}
    </section>
  );
}
