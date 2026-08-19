import type { FastifyInstance } from "fastify";
import { uploadResponseSchema } from "@label-maker/shared";
import { ValidationFailedError } from "../lib/errors.js";
import { handleUpload } from "../services/upload-service.js";

export default async function uploadRoutes(fastify: FastifyInstance): Promise<void> {
  fastify.post("/uploads", async (request, reply) => {
    const file = await request.file();
    if (!file) {
      throw new ValidationFailedError("Multipart request must include a file field.");
    }

    const buffer = await file.toBuffer();

    const result = await handleUpload(
      fastify.prisma,
      fastify.storage,
      fastify.queues.ingestDocument,
      {
        buffer,
        originalFilename: file.filename,
        mimeType: file.mimetype,
        maxUploadBytes: fastify.config.maxUploadBytes,
      },
    );

    const body = uploadResponseSchema.parse(result);
    return reply.code(201).send(body);
  });
}
