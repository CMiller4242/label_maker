import { useMemo, useState } from "react";
import { StepIndicator, type Step } from "./components/StepIndicator";
import { UploadStage } from "./components/UploadStage";
import { ProductReviewStage } from "./components/ProductReviewStage";
import { LabelPlanPreview } from "./components/LabelPlanPreview";
import { GenerateStage } from "./components/GenerateStage";
import { useProducts } from "./hooks/useProducts";
import { useLabelTemplates } from "./hooks/useLabelTemplates";
import { productRowIssue } from "./lib/productValidation";
import type { PreviewProduct } from "./lib/placementPreview";
import "./App.css";

type Stage = "upload" | "review" | "preview" | "generate";

const STEPS: Step[] = [
  { id: "upload", label: "Upload" },
  { id: "review", label: "Review products" },
  { id: "preview", label: "Preview" },
  { id: "generate", label: "Generate" },
];
const STAGE_INDEX: Record<Stage, number> = { upload: 0, review: 1, preview: 2, generate: 3 };

const DEFAULT_COPIES_PER_PRODUCT = 8;

export function App() {
  const [stage, setStage] = useState<Stage>("upload");
  const [documentId, setDocumentId] = useState<string | null>(null);
  const [selectedTemplateId, setSelectedTemplateId] = useState<string | null>(null);
  const [copiesPerProduct, setCopiesPerProduct] = useState(DEFAULT_COPIES_PER_PRODUCT);

  const products = useProducts(documentId);
  const templatesResult = useLabelTemplates();

  const selectedTemplate = useMemo(
    () => templatesResult.templates.find((t) => t.id === selectedTemplateId) ?? null,
    [templatesResult.templates, selectedTemplateId],
  );

  const readyProducts: PreviewProduct[] = useMemo(
    () =>
      products.rows
        .filter((r): r is Extract<typeof r, { isNew: false }> => !r.isNew)
        .filter((r) => r.include && productRowIssue(r) === null)
        .map((r) => ({
          id: r.id,
          sku: r.sku as string,
          description: r.description as string,
          priceCents: r.priceCents as number,
        })),
    [products.rows],
  );

  const startOver = () => {
    setStage("upload");
    setDocumentId(null);
    setSelectedTemplateId(null);
    setCopiesPerProduct(DEFAULT_COPIES_PER_PRODUCT);
  };

  return (
    <div className="app">
      <header className="app__header">
        <h1>Label Maker</h1>
        <p className="app__tagline">Avery 5155 label generation</p>
      </header>

      <StepIndicator steps={STEPS} activeIndex={STAGE_INDEX[stage]} />

      <main className="app__main">
        {stage === "upload" && (
          <UploadStage
            onDocumentReady={(id) => {
              setDocumentId(id);
              setStage("review");
            }}
          />
        )}

        {stage === "review" && documentId && (
          <ProductReviewStage
            products={products}
            templatesResult={templatesResult}
            selectedTemplateId={selectedTemplateId}
            onSelectTemplate={setSelectedTemplateId}
            copiesPerProduct={copiesPerProduct}
            onChangeCopiesPerProduct={setCopiesPerProduct}
            onContinue={() => setStage("preview")}
          />
        )}

        {stage === "preview" && documentId && selectedTemplate && (
          <LabelPlanPreview
            template={selectedTemplate}
            products={readyProducts}
            copiesPerProduct={copiesPerProduct}
            onBack={() => setStage("review")}
            onContinue={() => setStage("generate")}
          />
        )}

        {stage === "generate" && documentId && selectedTemplate && (
          <GenerateStage
            request={{
              sourceDocumentId: documentId,
              labelTemplateId: selectedTemplate.id,
              copiesPerProduct,
              includeDebugArtifact: false,
            }}
            onBack={() => setStage("preview")}
            onStartOver={startOver}
          />
        )}
      </main>
    </div>
  );
}
