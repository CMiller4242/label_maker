import { useMemo, useState } from "react";
import { fixedGridTemplateConfigSchema, type LabelTemplate } from "@label-maker/shared";
import { buildSheetGrid, computePlacementPreview, gridTemplateColumns, type PreviewProduct } from "../lib/placementPreview";
import "./LabelPlanPreview.css";

export interface LabelPlanPreviewProps {
  template: LabelTemplate;
  products: PreviewProduct[];
  copiesPerProduct: number;
  onBack: () => void;
  onContinue: () => void;
}

export function LabelPlanPreview({ template, products, copiesPerProduct, onBack, onContinue }: LabelPlanPreviewProps) {
  const [showOutlines, setShowOutlines] = useState(false);
  const [sheetIndex, setSheetIndex] = useState(1);

  const configParse = fixedGridTemplateConfigSchema.safeParse(template.configJson);
  const logicalGrid = configParse.success ? (configParse.data.logicalGrid ?? null) : null;

  const plan = useMemo(
    () => computePlacementPreview(products, template, copiesPerProduct),
    [products, template, copiesPerProduct],
  );

  const totalSheets = plan.totalSheets;
  const currentSheet = Math.min(Math.max(sheetIndex, 1), Math.max(totalSheets, 1));
  const emptySlotCount = totalSheets * template.labelsPerSheet - plan.totalPlacements;

  const sheetGrid = useMemo(() => {
    if (!logicalGrid || totalSheets === 0) return null;
    return buildSheetGrid(logicalGrid, currentSheet, plan.placements);
  }, [logicalGrid, totalSheets, currentSheet, plan.placements]);

  return (
    <section aria-labelledby="preview-stage-heading" className="preview-stage">
      <h2 id="preview-stage-heading">3. Label plan preview</h2>
      <p className="preview-stage__disclaimer">
        This is a visual review aid rendered in your browser - it is <strong>not</strong> a
        print-accurate guarantee. Physical alignment can only be confirmed by printing the
        generated DOCX at Actual Size / 100% and checking it against a real sheet.
      </p>

      <dl className="preview-stage__summary">
        <div>
          <dt>Pages</dt>
          <dd>{totalSheets}</dd>
        </div>
        <div>
          <dt>Filled labels</dt>
          <dd>{plan.totalPlacements}</dd>
        </div>
        <div>
          <dt>Blank labels</dt>
          <dd>{Math.max(emptySlotCount, 0)}</dd>
        </div>
        <div>
          <dt>Products x copies</dt>
          <dd>
            {products.length} x {copiesPerProduct}
          </dd>
        </div>
      </dl>

      <div className="preview-stage__controls">
        <label className="preview-stage__toggle">
          <input
            type="checkbox"
            checked={showOutlines}
            onChange={(event) => setShowOutlines(event.target.checked)}
          />
          Show review outlines
        </label>

        {totalSheets > 1 && (
          <div className="preview-stage__pager">
            <button
              type="button"
              className="button button--secondary"
              onClick={() => setSheetIndex((n) => Math.max(1, n - 1))}
              disabled={currentSheet <= 1}
            >
              Previous page
            </button>
            <span aria-live="polite">
              Page {currentSheet} of {totalSheets}
            </span>
            <button
              type="button"
              className="button button--secondary"
              onClick={() => setSheetIndex((n) => Math.min(totalSheets, n + 1))}
              disabled={currentSheet >= totalSheets}
            >
              Next page
            </button>
          </div>
        )}
      </div>

      {totalSheets === 0 && <p className="preview-stage__empty">No labels to preview yet.</p>}

      {totalSheets > 0 && !logicalGrid && (
        <p className="preview-stage__empty">
          This template's physical layout has not been inspected yet, so a real preview can't be
          shown. Generation is unaffected - it uses the actual source template geometry directly.
        </p>
      )}

      {sheetGrid && logicalGrid && (
        <div
          className={`preview-stage__sheet${showOutlines ? " preview-stage__sheet--outlined" : ""}`}
          style={{ gridTemplateColumns: gridTemplateColumns(logicalGrid.rawGridColumnWidthsTwips) }}
          role="img"
          aria-label={`Preview of sheet ${currentSheet}: ${logicalGrid.logicalColumns} label columns by ${logicalGrid.logicalRows} rows, ${logicalGrid.spacerRawColumnIndexes.length} spacer columns between them`}
        >
          {sheetGrid.flatMap((row, rowIndex) =>
            row.map((cell) => {
              const key = `${rowIndex}-${cell.rawColumnIndex}`;
              if (cell.kind === "spacer") {
                return <div key={key} className="preview-stage__cell preview-stage__cell--spacer" aria-hidden="true" />;
              }
              return (
                <div key={key} className="preview-stage__cell preview-stage__cell--writable">
                  {cell.placement && (
                    <>
                      <span className="preview-stage__cell-sku">{cell.placement.sku}</span>
                      <span className="preview-stage__cell-description">{cell.placement.description}</span>
                    </>
                  )}
                </div>
              );
            }),
          )}
        </div>
      )}

      <div className="preview-stage__actions">
        <button type="button" className="button button--secondary" onClick={onBack}>
          Back to product review
        </button>
        <button type="button" className="button button--primary" onClick={onContinue}>
          Continue to generate
        </button>
      </div>
    </section>
  );
}
