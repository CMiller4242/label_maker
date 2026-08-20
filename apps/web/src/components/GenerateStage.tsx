import { useEffect, useRef } from "react";
import type { CreateLabelRunRequest } from "@label-maker/shared";
import { useLabelRun } from "../hooks/useLabelRun";
import { ErrorBanner } from "./ErrorBanner";
import { LoadingIndicator } from "./LoadingIndicator";
import "./GenerateStage.css";

export interface GenerateStageProps {
  request: CreateLabelRunRequest;
  onBack: () => void;
  onStartOver: () => void;
}

const PRINT_GUIDANCE =
  "Open in Microsoft Word and print at Actual Size / 100%. Disable Fit to Page, " +
  "Scale to Paper Size, and printer-driver scaling.";

export function GenerateStage({ request, onBack, onStartOver }: GenerateStageProps) {
  const { status, errorMessage, errorCode, result, generate, retry, download, downloadError } = useLabelRun();
  const startedRef = useRef(false);

  // Auto-starts once when this stage mounts - the user already confirmed
  // everything in review/preview, so "Continue to generate" (stage 3's
  // button) is the explicit trigger; this stage shows pending/success/error.
  useEffect(() => {
    if (!startedRef.current) {
      startedRef.current = true;
      generate(request);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps -- runs exactly once per mount, by design
  }, []);

  return (
    <section aria-labelledby="generate-stage-heading" className="generate-stage">
      <h2 id="generate-stage-heading">4. Generate and download</h2>

      {status === "pending" && <LoadingIndicator label="Generating your Avery 5155 DOCX…" />}

      {status === "failed" && (
        <>
          <ErrorBanner message={errorMessage ?? "Label run generation failed."} onRetry={retry} />
          {errorCode === "TEMPLATE_UNAVAILABLE" && (
            <p className="generate-stage__hint">
              The Avery 5155 source template could not be used. This is a server-side template
              issue, not something you can fix from here.
            </p>
          )}
          <button type="button" className="button button--secondary" onClick={onBack}>
            Back to preview
          </button>
        </>
      )}

      {status === "succeeded" && result && (
        <div className="generate-stage__success" role="status">
          <p>DOCX generated successfully.</p>
          <dl className="generate-stage__summary">
            <div>
              <dt>Sheets</dt>
              <dd>{result.labelRun.sheetCount ?? "-"}</dd>
            </div>
            <div>
              <dt>Filled labels</dt>
              <dd>{result.labelRun.filledSlotCount ?? "-"}</dd>
            </div>
            <div>
              <dt>Blank labels</dt>
              <dd>{result.labelRun.emptySlotCount ?? "-"}</dd>
            </div>
          </dl>

          <button type="button" className="button button--primary" onClick={() => void download()}>
            Download DOCX
          </button>
          {downloadError && <ErrorBanner message={downloadError} onRetry={() => void download()} />}

          <div className="generate-stage__print-guidance">
            <h3>Before you print</h3>
            <p>{PRINT_GUIDANCE}</p>
            <p className="generate-stage__disclaimer">
              This confirms the DOCX generated successfully - it does not confirm physical print
              alignment on real Avery 5155 stock. Perform a plain-paper overlay test at 100% scale
              before printing a full batch.
            </p>
          </div>

          <button type="button" className="button button--secondary" onClick={onStartOver}>
            Start a new label run
          </button>
        </div>
      )}
    </section>
  );
}
