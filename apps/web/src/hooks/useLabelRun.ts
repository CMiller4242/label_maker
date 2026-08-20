import { useCallback, useState } from "react";
import type { CreateLabelRunRequest, CreateLabelRunResponse } from "@label-maker/shared";
import { ApiError, NetworkError } from "../api/client";
import { createLabelRun, downloadLabelRunArtifact } from "../api/endpoints";

export type LabelRunStatus = "idle" | "pending" | "succeeded" | "failed";

export interface UseLabelRunResult {
  status: LabelRunStatus;
  errorMessage: string | null;
  errorCode: string | null;
  result: CreateLabelRunResponse | null;
  generate: (input: CreateLabelRunRequest) => void;
  retry: () => void;
  download: () => Promise<void>;
  downloadError: string | null;
}

function describeError(error: unknown): { message: string; code: string | null } {
  if (error instanceof ApiError) return { message: error.message, code: error.code };
  if (error instanceof NetworkError) return { message: error.message, code: null };
  if (error instanceof Error) return { message: error.message, code: null };
  return { message: "Something went wrong.", code: null };
}

/** Drives POST /label-runs, then GET /label-runs/:id/download on request. */
export function useLabelRun(): UseLabelRunResult {
  const [status, setStatus] = useState<LabelRunStatus>("idle");
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [errorCode, setErrorCode] = useState<string | null>(null);
  const [result, setResult] = useState<CreateLabelRunResponse | null>(null);
  const [lastInput, setLastInput] = useState<CreateLabelRunRequest | null>(null);
  const [downloadError, setDownloadError] = useState<string | null>(null);

  const run = useCallback((input: CreateLabelRunRequest) => {
    setLastInput(input);
    setStatus("pending");
    setErrorMessage(null);
    setErrorCode(null);
    setDownloadError(null);
    void (async () => {
      try {
        const response = await createLabelRun(input);
        setResult(response);
        setStatus("succeeded");
      } catch (error) {
        const described = describeError(error);
        setErrorMessage(described.message);
        setErrorCode(described.code);
        setStatus("failed");
      }
    })();
  }, []);

  const retry = useCallback(() => {
    if (lastInput) run(lastInput);
  }, [lastInput, run]);

  const download = useCallback(async () => {
    if (!result) return;
    setDownloadError(null);
    try {
      const { blob, filename } = await downloadLabelRunArtifact(result.labelRun.id);
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = filename ?? `label-run-${result.labelRun.id}.docx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
    } catch (error) {
      setDownloadError(describeError(error).message);
    }
  }, [result]);

  return { status, errorMessage, errorCode, result, generate: run, retry, download, downloadError };
}
