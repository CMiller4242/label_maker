import type { FastifyInstance } from "fastify";
import { z } from "zod";
import { productListResponseSchema } from "@label-maker/shared";
import { listProductsForDocument } from "../services/product-service.js";

const paramsSchema = z.object({ documentId: z.string().uuid() });

export default async function productRoutes(fastify: FastifyInstance): Promise<void> {
  fastify.get("/documents/:documentId/products", async (request) => {
    const { documentId } = paramsSchema.parse(request.params);
    const products = await listProductsForDocument(fastify.prisma, documentId);
    return productListResponseSchema.parse({ documentId, products });
  });
}
