import "./ErrorBanner.css";

export interface ErrorBannerProps {
  message: string;
  onRetry?: () => void;
  retryLabel?: string;
}

/** An accessible, actionable error message with an optional retry action. Announced to screen readers via role="alert". */
export function ErrorBanner({ message, onRetry, retryLabel = "Retry" }: ErrorBannerProps) {
  return (
    <div className="error-banner" role="alert">
      <span className="error-banner__message">{message}</span>
      {onRetry && (
        <button type="button" className="error-banner__retry" onClick={onRetry}>
          {retryLabel}
        </button>
      )}
    </div>
  );
}
