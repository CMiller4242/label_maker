import type { FastifyInstance } from "fastify";
import { z } from "zod";
import { createLabelRunRequestSchema } from "@label-maker/shared";
import { createLabelRun, getLabelRunArtifactForDownload } from "../services/label-run-service.js";

const paramsSchema = z.object({ labelRunId: z.string().uuid() });

export default async function labelRunRoutes(fastify: FastifyInstance): Promise<void> {
  fastify.post("/label-runs", async (request, reply) => {
    const input = createLabelRunRequestSchema.parse(request.body);
    const result = await createLabelRun(fastify.prisma, fastify.storage, fastify.templates, input);
    return reply.code(201).send(result);
  });

  fastify.get("/label-runs/:labelRunId/download", async (request, reply) => {
    const { labelRunId } = paramsSchema.parse(request.params);
    const download = await getLabelRunArtifactForDownload(
      fastify.prisma,
      fastify.storage,
      labelRunId,
    );

    reply.header("Content-Type", download.mimeType);
    reply.header("Content-Disposition", `attachment; filename="${download.filename}"`);
    return reply.send(download.buffer);
  });

  fastify.get("/label-templates", async () => {
    const templates = await fastify.prisma.labelTemplate.findMany({
      orderBy: { createdAt: "asc" },
    });
    return { templates };
  });
}
