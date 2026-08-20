import "./LoadingIndicator.css";

export interface LoadingIndicatorProps {
  label: string;
}

/** A small inline loading state, announced politely to screen readers. */
export function LoadingIndicator({ label }: LoadingIndicatorProps) {
  return (
    <div className="loading-indicator" role="status" aria-live="polite">
      <span className="loading-indicator__spinner" aria-hidden="true" />
      <span>{label}</span>
    </div>
  );
}
