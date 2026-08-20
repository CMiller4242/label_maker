import "./StepIndicator.css";

export interface Step {
  id: string;
  label: string;
}

export interface StepIndicatorProps {
  steps: Step[];
  activeIndex: number;
}

/** Horizontal (wraps on narrow screens) step tracker for the 4-stage workflow. */
export function StepIndicator({ steps, activeIndex }: StepIndicatorProps) {
  return (
    <nav aria-label="Workflow progress" className="step-indicator">
      <ol>
        {steps.map((step, index) => {
          const state = index < activeIndex ? "done" : index === activeIndex ? "active" : "upcoming";
          return (
            <li key={step.id} className={`step-indicator__item step-indicator__item--${state}`}>
              <span className="step-indicator__badge" aria-hidden="true">
                {state === "done" ? "✓" : index + 1}
              </span>
              <span className="step-indicator__label">
                {step.label}
                {state === "active" && <span className="visually-hidden"> (current step)</span>}
              </span>
            </li>
          );
        })}
      </ol>
    </nav>
  );
}
