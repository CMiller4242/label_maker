import type { ComponentProps } from "react";
import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import type { LabelTemplate, Product } from "@label-maker/shared";
import type { PersistedRow, UseProductsResult } from "@label-maker/web/hooks/useProducts";
import type { UseLabelTemplatesResult } from "@label-maker/web/hooks/useLabelTemplates";

const { ProductReviewStage } = await import("@label-maker/web/components/ProductReviewStage");

function makeProduct(overrides: Partial<Product> = {}): PersistedRow {
  return {
    id: "11111111-1111-1111-1111-111111111111",
    sourceDocumentId: "22222222-2222-2222-2222-222222222222",
    sourcePageId: null,
    sourceRowNumber: null,
    sku: "SKU-1",
    description: "Widget",
    priceCents: 1099,
    include: true,
    status: "APPROVED",
    confidence: 1,
    extractionNotesJson: null,
    createdAt: new Date(),
    updatedAt: new Date(),
    isNew: false,
    ...overrides,
  };
}

function makeProductsResult(overrides: Partial<UseProductsResult> = {}): UseProductsResult {
  return {
    status: "loaded",
    errorMessage: null,
    rows: [makeProduct()],
    rowErrors: {},
    savingRowIds: new Set(),
    refetch: vi.fn(),
    addRow: vi.fn(),
    removeRow: vi.fn(),
    editRow: vi.fn(),
    approveRow: vi.fn(),
    setInclude: vi.fn(),
    ...overrides,
  };
}

const TEMPLATES: LabelTemplate[] = [
  {
    id: "avery-5155",
    displayName: "Avery 5155",
    renderingMode: "FIXED_GRID",
    columns: 2,
    rows: 10,
    labelsPerSheet: 20,
    templateStorageKey: null,
    templateVersion: null,
    sourceTemplateSha256: null,
    configJson: {},
    isPreset: true,
    createdAt: new Date(),
    updatedAt: new Date(),
  },
  {
    id: "avery-22802",
    displayName: "Avery 22802",
    renderingMode: "FIXED_GRID",
    columns: 3,
    rows: 10,
    labelsPerSheet: 30,
    templateStorageKey: null,
    templateVersion: null,
    sourceTemplateSha256: null,
    configJson: {},
    isPreset: true,
    createdAt: new Date(),
    updatedAt: new Date(),
  },
];

function makeTemplatesResult(overrides: Partial<UseLabelTemplatesResult> = {}): UseLabelTemplatesResult {
  return {
    status: "loaded",
    errorMessage: null,
    templates: TEMPLATES,
    retry: vi.fn(),
    ...overrides,
  };
}

function renderStage(props: Partial<ComponentProps<typeof ProductReviewStage>> = {}) {
  const onContinue = vi.fn();
  const onSelectTemplate = vi.fn();
  const utils = render(
    <ProductReviewStage
      products={makeProductsResult()}
      templatesResult={makeTemplatesResult()}
      selectedTemplateId="avery-5155"
      onSelectTemplate={onSelectTemplate}
      copiesPerProduct={8}
      onChangeCopiesPerProduct={vi.fn()}
      onContinue={onContinue}
      {...props}
    />,
  );
  return { ...utils, onContinue, onSelectTemplate };
}

afterEach(() => {
  cleanup();
});

describe("ProductReviewStage", () => {
  it("shows a loading indicator while products are loading", () => {
    renderStage({ products: makeProductsResult({ status: "loading", rows: [] }) });
    expect(screen.getByRole("status")).toHaveTextContent(/loading products/i);
  });

  it("shows a retryable error when products fail to load", () => {
    const refetch = vi.fn();
    renderStage({
      products: makeProductsResult({ status: "failed", errorMessage: "Server exploded.", rows: [], refetch }),
    });
    expect(screen.getByRole("alert")).toHaveTextContent("Server exploded.");
    fireEvent.click(screen.getByRole("button", { name: /retry/i }));
    expect(refetch).toHaveBeenCalledTimes(1);
  });

  it("enables Continue once every row is valid and Avery 5155 is selected", () => {
    const { onContinue } = renderStage();
    const continueButton = screen.getByRole("button", { name: /continue to label plan preview/i });
    expect(continueButton).toBeEnabled();
    fireEvent.click(continueButton);
    expect(onContinue).toHaveBeenCalledTimes(1);
  });

  it("blocks Continue and explains why when an included row is missing required data", () => {
    renderStage({
      products: makeProductsResult({ rows: [makeProduct({ sku: null })] }),
    });
    const continueButton = screen.getByRole("button", { name: /continue to label plan preview/i });
    expect(continueButton).toBeDisabled();
    expect(screen.getByText(/resolve 1 row issue/i)).toBeInTheDocument();
  });

  it("blocks Continue and explains why when a row needs approval", () => {
    renderStage({
      products: makeProductsResult({ rows: [makeProduct({ status: "NEEDS_REVIEW" })] }),
    });
    expect(screen.getByRole("button", { name: /continue to label plan preview/i })).toBeDisabled();
    expect(screen.getByText(/needs approval before it can be generated/i)).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /^approve$/i })).toBeInTheDocument();
  });

  it("blocks Continue when no product is included, even if all rows are individually valid", () => {
    renderStage({
      products: makeProductsResult({ rows: [makeProduct({ include: false })] }),
    });
    expect(screen.getByRole("button", { name: /continue to label plan preview/i })).toBeDisabled();
    expect(screen.getByText(/include at least one product/i)).toBeInTheDocument();
  });

  it("blocks Continue while no template is selected, and disables unsupported templates", () => {
    renderStage({ selectedTemplateId: null });
    expect(screen.getByRole("button", { name: /continue to label plan preview/i })).toBeDisabled();
    expect(screen.getByRole("radio", { name: /avery 22802/i })).toBeDisabled();
    expect(screen.getByText(/select the avery 5155 template/i)).toBeInTheDocument();
  });

  it("calls removeRow for a persisted row's Remove button", () => {
    const removeRow = vi.fn();
    renderStage({ products: makeProductsResult({ removeRow }) });
    fireEvent.click(screen.getByRole("button", { name: /remove/i }));
    expect(removeRow).toHaveBeenCalledWith("11111111-1111-1111-1111-111111111111");
  });

  it("calls approveRow when a NEEDS_REVIEW row is approved", () => {
    const approveRow = vi.fn();
    renderStage({
      products: makeProductsResult({ rows: [makeProduct({ status: "NEEDS_REVIEW" })], approveRow }),
    });
    fireEvent.click(screen.getByRole("button", { name: /^approve$/i }));
    expect(approveRow).toHaveBeenCalledWith("11111111-1111-1111-1111-111111111111");
  });

  it("commits a SKU edit on blur via editRow", () => {
    const editRow = vi.fn();
    renderStage({ products: makeProductsResult({ editRow }) });
    const skuInput = screen.getByLabelText(/^sku$/i);
    fireEvent.change(skuInput, { target: { value: "NEW-SKU" } });
    fireEvent.blur(skuInput);
    expect(editRow).toHaveBeenCalledWith("11111111-1111-1111-1111-111111111111", { sku: "NEW-SKU" });
  });
});
