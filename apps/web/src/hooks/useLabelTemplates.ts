import { useCallback, useEffect, useState } from "react";
import type { LabelTemplate } from "@label-maker/shared";
import { ApiError, NetworkError } from "../api/client";
import { listLabelTemplates } from "../api/endpoints";

export type TemplateLoadStatus = "idle" | "loading" | "loaded" | "failed";

export interface UseLabelTemplatesResult {
  status: TemplateLoadStatus;
  errorMessage: string | null;
  templates: LabelTemplate[];
  retry: () => void;
}

function describeError(error: unknown): string {
  if (error instanceof ApiError) return error.message;
  if (error instanceof NetworkError) return error.message;
  if (error instanceof Error) return error.message;
  return "Something went wrong.";
}

/** GET /label-templates, once. */
export function useLabelTemplates(): UseLabelTemplatesResult {
  const [status, setStatus] = useState<TemplateLoadStatus>("idle");
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [templates, setTemplates] = useState<LabelTemplate[]>([]);

  const load = useCallback(() => {
    setStatus("loading");
    setErrorMessage(null);
    void (async () => {
      try {
        const result = await listLabelTemplates();
        setTemplates(result.templates);
        setStatus("loaded");
      } catch (error) {
        setStatus("failed");
        setErrorMessage(describeError(error));
      }
    })();
  }, []);

  useEffect(load, [load]);

  return { status, errorMessage, templates, retry: load };
}
