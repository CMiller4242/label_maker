import type { LabelTemplate } from "@label-maker/shared";
import { ErrorBanner } from "./ErrorBanner";
import { LoadingIndicator } from "./LoadingIndicator";
import { centsToDollarsInputValue, parseDollarsToCents } from "../lib/money";
import { productRowIssue } from "../lib/productValidation";
import type { DraftRow, ProductRow, UseProductsResult } from "../hooks/useProducts";
import type { UseLabelTemplatesResult } from "../hooks/useLabelTemplates";
import "./ProductReviewStage.css";

/** Only Avery 5155 is supported end to end right now - see the project's current scope. */
const SUPPORTED_TEMPLATE_ID = "avery-5155";

function rowKey(row: ProductRow): string {
  return row.isNew ? row.tempId : row.id;
}

function draftRowIssue(row: DraftRow): string | null {
  if (row.sku.trim().length === 0) return "SKU is required.";
  if (row.description.trim().length === 0) return "Product name is required.";
  if (row.priceCents === null) return "Price is required.";
  return null;
}

function rowIssue(row: ProductRow): string | null {
  return row.isNew ? draftRowIssue(row) : productRowIssue(row);
}

export interface ProductReviewStageProps {
  products: UseProductsResult;
  templatesResult: UseLabelTemplatesResult;
  selectedTemplateId: string | null;
  onSelectTemplate: (templateId: string) => void;
  copiesPerProduct: number;
  onChangeCopiesPerProduct: (value: number) => void;
  onContinue: () => void;
}

export function ProductReviewStage({
  products,
  templatesResult,
  selectedTemplateId,
  onSelectTemplate,
  copiesPerProduct,
  onChangeCopiesPerProduct,
  onContinue,
}: ProductReviewStageProps) {
  const includedRows = products.rows.filter((r) => r.include);
  const blockingIssueCount = includedRows.filter((r) => rowIssue(r) !== null).length;
  const canContinue =
    products.status === "loaded" &&
    includedRows.length > 0 &&
    blockingIssueCount === 0 &&
    selectedTemplateId === SUPPORTED_TEMPLATE_ID &&
    copiesPerProduct > 0;

  return (
    <section aria-labelledby="review-stage-heading" className="review-stage">
      <h2 id="review-stage-heading">2. Review products</h2>

      {products.status === "loading" && <LoadingIndicator label="Loading products…" />}
      {products.status === "failed" && (
        <ErrorBanner message={products.errorMessage ?? "Could not load products."} onRetry={products.refetch} />
      )}

      {products.status === "loaded" && (
        <>
          {products.rows.length === 0 ? (
            <p className="review-stage__empty">
              No products were found in this file. Add a row manually below, or go back and upload a
              different file.
            </p>
          ) : (
            <div className="review-stage__table-wrap">
              <table>
                <caption className="visually-hidden">Products parsed from the uploaded file, editable</caption>
                <thead>
                  <tr>
                    <th scope="col">Include</th>
                    <th scope="col">SKU</th>
                    <th scope="col">Product name</th>
                    <th scope="col">Price</th>
                    <th scope="col">Status</th>
                    <th scope="col">Row status</th>
                    <th scope="col">
                      <span className="visually-hidden">Actions</span>
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {products.rows.map((row) => {
                    const key = rowKey(row);
                    const issue = rowIssue(row);
                    const isSaving = products.savingRowIds.has(key);
                    const saveError = products.rowErrors[key];
                    return (
                      <tr key={key} className={issue ? "review-stage__row--invalid" : undefined}>
                        <td>
                          <input
                            type="checkbox"
                            checked={row.include}
                            disabled={row.isNew || isSaving}
                            aria-label={`Include ${row.sku || "this row"} in the label run`}
                            onChange={(event) => products.setInclude(key, event.target.checked)}
                          />
                        </td>
                        <td>
                          <label className="visually-hidden" htmlFor={`sku-${key}`}>
                            SKU
                          </label>
                          <input
                            id={`sku-${key}`}
                            type="text"
                            defaultValue={row.sku ?? ""}
                            disabled={isSaving}
                            onBlur={(event) =>
                              products.editRow(key, { sku: event.target.value.trim() || null })
                            }
                          />
                        </td>
                        <td>
                          <label className="visually-hidden" htmlFor={`description-${key}`}>
                            Product name
                          </label>
                          <input
                            id={`description-${key}`}
                            type="text"
                            defaultValue={row.description ?? ""}
                            disabled={isSaving}
                            onBlur={(event) =>
                              products.editRow(key, { description: event.target.value.trim() || null })
                            }
                          />
                        </td>
                        <td>
                          <label className="visually-hidden" htmlFor={`price-${key}`}>
                            Price in dollars
                          </label>
                          <div className="review-stage__price-field">
                            <span aria-hidden="true">$</span>
                            <input
                              id={`price-${key}`}
                              type="text"
                              inputMode="decimal"
                              defaultValue={centsToDollarsInputValue(row.priceCents)}
                              disabled={isSaving}
                              onBlur={(event) =>
                                products.editRow(key, { priceCents: parseDollarsToCents(event.target.value) })
                              }
                            />
                          </div>
                        </td>
                        <td>
                          {row.isNew ? (
                            <span className="status-pill status-pill--new">Not saved yet</span>
                          ) : (
                            <span className={`status-pill status-pill--${row.status.toLowerCase()}`}>
                              {row.status.replace("_", " ")}
                            </span>
                          )}
                        </td>
                        <td>
                          {isSaving && <LoadingIndicator label="Saving…" />}
                          {!isSaving && saveError && <ErrorBanner message={saveError} />}
                          {!isSaving && !saveError && issue && row.include && (
                            <span className="review-stage__issue">{issue}</span>
                          )}
                          {!isSaving &&
                            !saveError &&
                            !row.isNew &&
                            row.include &&
                            row.status === "NEEDS_REVIEW" && (
                              <button
                                type="button"
                                className="button button--secondary"
                                onClick={() => products.approveRow(key)}
                              >
                                Approve
                              </button>
                            )}
                        </td>
                        <td>
                          <button
                            type="button"
                            className="button button--danger"
                            disabled={isSaving}
                            onClick={() => products.removeRow(key)}
                          >
                            Remove
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}

          <button type="button" className="button button--secondary" onClick={products.addRow}>
            + Add row
          </button>
        </>
      )}

      <div className="review-stage__settings card">
        <div className="review-stage__field">
          <label htmlFor="copies-per-product">Copies per product</label>
          <input
            id="copies-per-product"
            type="number"
            min={1}
            step={1}
            value={copiesPerProduct}
            onChange={(event) => onChangeCopiesPerProduct(Math.max(1, Number(event.target.value) || 1))}
          />
          <p className="review-stage__field-hint">
            Applies to every included product equally - the current label run API does not support a
            different quantity per row.
          </p>
        </div>

        <fieldset className="review-stage__templates">
          <legend>Label template</legend>
          {templatesResult.status === "loading" && <LoadingIndicator label="Loading templates…" />}
          {templatesResult.status === "failed" && (
            <ErrorBanner
              message={templatesResult.errorMessage ?? "Could not load label templates."}
              onRetry={templatesResult.retry}
            />
          )}
          {templatesResult.status === "loaded" &&
            templatesResult.templates.map((template: LabelTemplate) => {
              const supported = template.id === SUPPORTED_TEMPLATE_ID;
              return (
                <label key={template.id} className="review-stage__template-option">
                  <input
                    type="radio"
                    name="label-template"
                    value={template.id}
                    disabled={!supported}
                    checked={selectedTemplateId === template.id}
                    onChange={() => onSelectTemplate(template.id)}
                  />
                  <span>
                    {template.displayName}
                    {!supported && <span className="review-stage__coming-soon"> (coming soon)</span>}
                  </span>
                </label>
              );
            })}
        </fieldset>
      </div>

      <div className="review-stage__actions">
        <button type="button" className="button button--primary" disabled={!canContinue} onClick={onContinue}>
          Continue to label plan preview
        </button>
        {!canContinue && products.status === "loaded" && (
          <p className="review-stage__blocked-hint">
            {includedRows.length === 0
              ? "Include at least one product to continue."
              : blockingIssueCount > 0
                ? `Resolve ${blockingIssueCount} row issue${blockingIssueCount === 1 ? "" : "s"} to continue.`
                : "Select the Avery 5155 template and a valid copy count to continue."}
          </p>
        )}
      </div>
    </section>
  );
}
