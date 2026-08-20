import type { PrismaClient, Product } from "@label-maker/database";
import type { CreateProductRequest, UpdateProductRequest } from "@label-maker/shared";
import { NotFoundError } from "../lib/errors.js";

export async function listProductsForDocument(
  prisma: PrismaClient,
  documentId: string,
): Promise<Product[]> {
  const sourceDocument = await prisma.sourceDocument.findUnique({ where: { id: documentId } });
  if (!sourceDocument) {
    throw new NotFoundError(`SourceDocument "${documentId}" not found.`);
  }

  return prisma.product.findMany({
    where: { sourceDocumentId: documentId },
    orderBy: { createdAt: "asc" },
  });
}

/**
 * Adds a single manually-entered product row to a document's product list -
 * the UI-driven equivalent of a row an extractor would otherwise have
 * proposed. No sourcePageId/sourceRowNumber (nothing backs it), confidence
 * 1 (a human typed it directly), status APPROVED (already reviewed, by
 * construction) - so it's immediately eligible for a label run alongside
 * ingested rows once `include` is true.
 */
export async function createManualProduct(
  prisma: PrismaClient,
  documentId: string,
  input: CreateProductRequest,
): Promise<Product> {
  const sourceDocument = await prisma.sourceDocument.findUnique({ where: { id: documentId } });
  if (!sourceDocument) {
    throw new NotFoundError(`SourceDocument "${documentId}" not found.`);
  }

  return prisma.product.create({
    data: {
      sourceDocumentId: documentId,
      sourcePageId: null,
      sourceRowNumber: null,
      sku: input.sku,
      description: input.description,
      priceCents: input.priceCents,
      include: input.include,
      status: "APPROVED",
      confidence: 1,
      extractionNotesJson: undefined,
    },
  });
}

/**
 * Updates only the fields present in `input` (sku/description/priceCents/
 * include/status). Editing content does not implicitly change status - a
 * caller correcting a NEEDS_REVIEW row's sku must also send
 * status: "APPROVED" for it to become eligible for a label run, mirroring
 * createLabelRun's existing include=true AND status IN (APPROVED,
 * AUTO_ACCEPTED) filter (unchanged here).
 */
export async function updateProduct(
  prisma: PrismaClient,
  productId: string,
  input: UpdateProductRequest,
): Promise<Product> {
  const existing = await prisma.product.findUnique({ where: { id: productId } });
  if (!existing) {
    throw new NotFoundError(`Product "${productId}" not found.`);
  }

  return prisma.product.update({
    where: { id: productId },
    data: {
      ...(input.sku !== undefined && { sku: input.sku }),
      ...(input.description !== undefined && { description: input.description }),
      ...(input.priceCents !== undefined && { priceCents: input.priceCents }),
      ...(input.include !== undefined && { include: input.include }),
      ...(input.status !== undefined && { status: input.status }),
    },
  });
}
