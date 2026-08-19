import type { FastifyInstance } from "fastify";
import { z } from "zod";
import { processingJobSchema } from "@label-maker/shared";
import { getJobById } from "../services/job-service.js";

const paramsSchema = z.object({ jobId: z.string().uuid() });

export default async function jobRoutes(fastify: FastifyInstance): Promise<void> {
  fastify.get("/jobs/:jobId", async (request) => {
    const { jobId } = paramsSchema.parse(request.params);
    const job = await getJobById(fastify.prisma, jobId);
    return processingJobSchema.parse(job);
  });
}
