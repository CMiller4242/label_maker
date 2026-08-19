export abstract class AppError extends Error {
  abstract readonly statusCode: number;
  abstract readonly code: string;
  readonly details?: unknown;

  constructor(message: string, details?: unknown) {
    super(message);
    this.details = details;
  }
}

export class NotFoundError extends AppError {
  readonly statusCode = 404;
  readonly code = "NOT_FOUND";
}

export class ValidationFailedError extends AppError {
  readonly statusCode = 400;
  readonly code = "VALIDATION_FAILED";
}

export class UnsupportedMediaTypeAppError extends AppError {
  readonly statusCode = 415;
  readonly code = "UNSUPPORTED_MEDIA_TYPE";
}

export class PayloadTooLargeAppError extends AppError {
  readonly statusCode = 413;
  readonly code = "PAYLOAD_TOO_LARGE";
}

export class ConflictError extends AppError {
  readonly statusCode = 409;
  readonly code = "CONFLICT";
}

/** The requested template's rendering mode does not support DOCX export yet (e.g. FLOATING). */
export class UnsupportedRenderingModeError extends AppError {
  readonly statusCode = 422;
  readonly code = "UNSUPPORTED_RENDERING_MODE";
}

/** The template row exists, but its source file is missing, unreadable, or not a usable/valid template. */
export class TemplateUnavailableError extends AppError {
  readonly statusCode = 422;
  readonly code = "TEMPLATE_UNAVAILABLE";
}
