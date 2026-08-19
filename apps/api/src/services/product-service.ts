import type { PrismaClient, Product } from "@label-maker/database";
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
