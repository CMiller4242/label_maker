import { useCallback, useEffect, useRef, useState } from "react";
import type { ProcessingJob } from "@label-maker/shared";
import { ApiError, NetworkError } from "../api/client";
import { getJob, uploadDocument } from "../api/endpoints";

const POLL_INTERVAL_MS = 1200;
/** Terminal ProcessingJob statuses - ingestion is done, one way or another (see apps/worker's ingest-source-document.ts). */
const TERMINAL_STATUSES = new Set<ProcessingJob["status"]>(["COMPLETED", "NEEDS_REVIEW", "FAILED"]);

export type IngestUploadStatus = "idle" | "uploading" | "processing" | "succeeded" | "failed";

export interface UseIngestUploadResult {
  status: IngestUploadStatus;
  file: File | null;
  documentId: string | null;
  job: ProcessingJob | null;
  /** Human-readable message for the current failure, or null when there isn't one. */
  errorMessage: string | null;
  upload: (file: File) => void;
  /** Re-attempts the same file after a failure - never silently retries; the user triggers it. */
  retry: () => void;
  reset: () => void;
}

function describeError(error: unknown): string {
  if (error instanceof ApiError) return error.message;
  if (error instanceof NetworkError) return error.message;
  if (error instanceof Error) return error.message;
  return "Something went wrong.";
}

/**
 * Drives POST /uploads -> poll GET /jobs/:jobId until the INGEST_DOCUMENT
 * job reaches a terminal state. Ingestion runs asynchronously in
 * apps/worker (a BullMQ consumer) - the worker must be running locally for
 * this to ever leave "processing".
 */
export function useIngestUpload(): UseIngestUploadResult {
  const [status, setStatus] = useState<IngestUploadStatus>("idle");
  const [file, setFile] = useState<File | null>(null);
  const [documentId, setDocumentId] = useState<string | null>(null);
  const [job, setJob] = useState<ProcessingJob | null>(null);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const pollTimeoutRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const requestIdRef = useRef(0);

  useEffect(
    () => () => {
      if (pollTimeoutRef.current) clearTimeout(pollTimeoutRef.current);
    },
    [],
  );

  const pollJob = useCallback((jobId: string, requestId: number) => {
    const tick = async () => {
      if (requestIdRef.current !== requestId) return; // superseded by a newer upload/retry
      try {
        const nextJob = await getJob(jobId);
        if (requestIdRef.current !== requestId) return;
        setJob(nextJob);
        if (TERMINAL_STATUSES.has(nextJob.status)) {
          if (nextJob.status === "FAILED") {
            setStatus("failed");
            setErrorMessage(nextJob.errorMessage ?? "Document processing failed.");
          } else {
            setStatus("succeeded");
          }
          return;
        }
        pollTimeoutRef.current = setTimeout(() => void tick(), POLL_INTERVAL_MS);
      } catch (error) {
        if (requestIdRef.current !== requestId) return;
        setStatus("failed");
        setErrorMessage(describeError(error));
      }
    };
    void tick();
  }, []);

  const startUpload = useCallback(
    (targetFile: File) => {
      const requestId = ++requestIdRef.current;
      if (pollTimeoutRef.current) clearTimeout(pollTimeoutRef.current);
      setFile(targetFile);
      setStatus("uploading");
      setErrorMessage(null);
      setJob(null);
      setDocumentId(null);

      void (async () => {
        try {
          const result = await uploadDocument(targetFile);
          if (requestIdRef.current !== requestId) return;
          setDocumentId(result.documentId);
          setStatus("processing");
          pollJob(result.jobId, requestId);
        } catch (error) {
          if (requestIdRef.current !== requestId) return;
          setStatus("failed");
          setErrorMessage(describeError(error));
        }
      })();
    },
    [pollJob],
  );

  const retry = useCallback(() => {
    if (file) startUpload(file);
  }, [file, startUpload]);

  const reset = useCallback(() => {
    requestIdRef.current += 1;
    if (pollTimeoutRef.current) clearTimeout(pollTimeoutRef.current);
    setStatus("idle");
    setFile(null);
    setDocumentId(null);
    setJob(null);
    setErrorMessage(null);
  }, []);

  return { status, file, documentId, job, errorMessage, upload: startUpload, retry, reset };
}
