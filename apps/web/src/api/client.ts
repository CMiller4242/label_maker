import type { ApiErrorBody } from "@label-maker/shared";

/** Base path every API request is issued against. In dev, Vite proxies "/api/*" to the Fastify API (see vite.config.ts) - the frontend never hardcodes the API's origin. */
const API_BASE_URL = (import.meta.env["VITE_API_BASE_URL"] as string | undefined) ?? "/api";

/** A structured error response the API actually returned (4xx/5xx with a parseable `{ error }` body). */
export class ApiError extends Error {
  readonly code: string;
  readonly statusCode: number;
  readonly details?: unknown;

  constructor(statusCode: number, body: ApiErrorBody) {
    super(body.error.message);
    this.name = "ApiError";
    this.code = body.error.code;
    this.statusCode = statusCode;
    this.details = body.error.details;
  }
}

/** The request never reached the server (offline, DNS failure, connection refused, aborted). Distinct from ApiError so the UI can show "check your connection" vs. a specific server-reported problem. */
export class NetworkError extends Error {
  constructor(cause: unknown) {
    super("Could not reach the server. Check your connection and try again.");
    this.name = "NetworkError";
    this.cause = cause;
  }
}

async function parseErrorBody(response: Response): Promise<ApiErrorBody> {
  try {
    const parsed: unknown = await response.json();
    if (
      parsed &&
      typeof parsed === "object" &&
      "error" in parsed &&
      parsed.error &&
      typeof parsed.error === "object" &&
      "code" in parsed.error &&
      "message" in parsed.error
    ) {
      return parsed as ApiErrorBody;
    }
  } catch {
    // Body wasn't JSON (or was empty) - fall through to a generic error below.
  }
  return {
    error: {
      code: "UNKNOWN_ERROR",
      message: `Request failed with status ${response.status}.`,
    },
  };
}

async function doFetch(path: string, init?: RequestInit): Promise<Response> {
  let response: Response;
  try {
    response = await fetch(`${API_BASE_URL}${path}`, init);
  } catch (error) {
    throw new NetworkError(error);
  }
  if (!response.ok) {
    throw new ApiError(response.status, await parseErrorBody(response));
  }
  return response;
}

/** GET/POST/PATCH with a JSON body and a JSON response. */
export async function apiJson<T>(
  path: string,
  init?: { method?: "GET" | "POST" | "PATCH" | "DELETE"; body?: unknown },
): Promise<T> {
  const response = await doFetch(path, {
    method: init?.method ?? "GET",
    headers: init?.body !== undefined ? { "Content-Type": "application/json" } : undefined,
    body: init?.body !== undefined ? JSON.stringify(init.body) : undefined,
  });
  return (await response.json()) as T;
}

/** POST a File as multipart/form-data (matches apps/api's @fastify/multipart POST /uploads route). */
export async function apiUpload<T>(path: string, file: File): Promise<T> {
  const formData = new FormData();
  formData.append("file", file);
  const response = await doFetch(path, { method: "POST", body: formData });
  return (await response.json()) as T;
}

/** GET a binary response (e.g. the generated DOCX download). Returns the blob plus the server-suggested filename from Content-Disposition, when present. */
export async function apiDownload(path: string): Promise<{ blob: Blob; filename: string | null }> {
  const response = await doFetch(path);
  const disposition = response.headers.get("Content-Disposition");
  const filenameMatch = disposition?.match(/filename="([^"]+)"/);
  return { blob: await response.blob(), filename: filenameMatch?.[1] ?? null };
}
