import type { FastifyInstance } from "fastify";
import { createLabelRunRequestSchema } from "@label-maker/shared";
import { createLabelRun } from "../services/label-run-service.js";

export default async function labelRunRoutes(fastify: FastifyInstance): Promise<void> {
  fastify.post("/label-runs", async (request, reply) => {
    const input = createLabelRunRequestSchema.parse(request.body);
    const result = await createLabelRun(fastify.prisma, fastify.storage, input);
    return reply.code(201).send(result);
  });

  fastify.get("/label-templates", async () => {
    const templates = await fastify.prisma.labelTemplate.findMany({
      orderBy: { createdAt: "asc" },
    });
    return { templates };
  });
}
