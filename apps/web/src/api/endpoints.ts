import {
  createLabelRunRequestSchema,
  createLabelRunResponseSchema,
  createProductRequestSchema,
  labelTemplateListResponseSchema,
  processingJobSchema,
  productListResponseSchema,
  productSchema,
  updateProductRequestSchema,
  uploadResponseSchema,
  type CreateLabelRunRequest,
  type CreateLabelRunResponse,
  type CreateProductRequest,
  type LabelTemplateListResponse,
  type ProcessingJob,
  type Product,
  type ProductListResponse,
  type UpdateProductRequest,
  type UploadResponse,
} from "@label-maker/shared";
import { apiDownload, apiJson, apiUpload } from "./client";

/** POST /uploads - matches apps/api/src/routes/uploads.ts exactly (multipart, field name "file"). */
export async function uploadDocument(file: File): Promise<UploadResponse> {
  const raw = await apiUpload<unknown>("/uploads", file);
  return uploadResponseSchema.parse(raw);
}

/** GET /jobs/:jobId */
export async function getJob(jobId: string): Promise<ProcessingJob> {
  const raw = await apiJson<unknown>(`/jobs/${jobId}`);
  return processingJobSchema.parse(raw);
}

/** GET /documents/:documentId/products */
export async function listProducts(documentId: string): Promise<ProductListResponse> {
  const raw = await apiJson<unknown>(`/documents/${documentId}/products`);
  return productListResponseSchema.parse(raw);
}

/** PATCH /products/:productId */
export async function updateProduct(productId: string, input: UpdateProductRequest): Promise<Product> {
  const body = updateProductRequestSchema.parse(input);
  const raw = await apiJson<unknown>(`/products/${productId}`, { method: "PATCH", body });
  return productSchema.parse(raw);
}

/** POST /documents/:documentId/products */
export async function createProduct(documentId: string, input: CreateProductRequest): Promise<Product> {
  const body = createProductRequestSchema.parse(input);
  const raw = await apiJson<unknown>(`/documents/${documentId}/products`, { method: "POST", body });
  return productSchema.parse(raw);
}

/** GET /label-templates */
export async function listLabelTemplates(): Promise<LabelTemplateListResponse> {
  const raw = await apiJson<unknown>("/label-templates");
  return labelTemplateListResponseSchema.parse(raw);
}

/** POST /label-runs */
export async function createLabelRun(input: CreateLabelRunRequest): Promise<CreateLabelRunResponse> {
  const body = createLabelRunRequestSchema.parse(input);
  const raw = await apiJson<unknown>("/label-runs", { method: "POST", body });
  return createLabelRunResponseSchema.parse(raw);
}

/** GET /label-runs/:labelRunId/download */
export async function downloadLabelRunArtifact(
  labelRunId: string,
): Promise<{ blob: Blob; filename: string | null }> {
  return apiDownload(`/label-runs/${labelRunId}/download`);
}
