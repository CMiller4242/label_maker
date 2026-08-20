import type { FastifyInstance } from "fastify";
import { z } from "zod";
import {
  createProductRequestSchema,
  productListResponseSchema,
  productSchema,
  updateProductRequestSchema,
} from "@label-maker/shared";
import {
  createManualProduct,
  listProductsForDocument,
  updateProduct,
} from "../services/product-service.js";

const documentParamsSchema = z.object({ documentId: z.string().uuid() });
const productParamsSchema = z.object({ productId: z.string().uuid() });

export default async function productRoutes(fastify: FastifyInstance): Promise<void> {
  fastify.get("/documents/:documentId/products", async (request) => {
    const { documentId } = documentParamsSchema.parse(request.params);
    const products = await listProductsForDocument(fastify.prisma, documentId);
    return productListResponseSchema.parse({ documentId, products });
  });

  fastify.post("/documents/:documentId/products", async (request, reply) => {
    const { documentId } = documentParamsSchema.parse(request.params);
    const input = createProductRequestSchema.parse(request.body);
    const product = await createManualProduct(fastify.prisma, documentId, input);
    return reply.code(201).send(productSchema.parse(product));
  });

  fastify.patch("/products/:productId", async (request) => {
    const { productId } = productParamsSchema.parse(request.params);
    const input = updateProductRequestSchema.parse(request.body);
    const product = await updateProduct(fastify.prisma, productId, input);
    return productSchema.parse(product);
  });
}
