import Fastify, { type FastifyInstance } from "fastify";
import multipart from "@fastify/multipart";
import { ZodError } from "zod";
import type { AppConfig } from "./config.js";
import { AppError } from "./lib/errors.js";
import prismaPlugin from "./plugins/prisma.js";
import storagePlugin from "./plugins/storage.js";
import queuesPlugin from "./plugins/queues.js";
import healthRoutes from "./routes/health.js";
import uploadRoutes from "./routes/uploads.js";
import jobRoutes from "./routes/jobs.js";
import productRoutes from "./routes/products.js";
import labelRunRoutes from "./routes/label-runs.js";

export async function buildApp(config: AppConfig): Promise<FastifyInstance> {
  const fastify = Fastify({ logger: { level: config.logLevel } });

  await fastify.register(storagePlugin, { config });
  await fastify.register(prismaPlugin);
  await fastify.register(queuesPlugin, { config });
  await fastify.register(multipart, { limits: { fileSize: config.maxUploadBytes } });

  fastify.setErrorHandler((error, request, reply) => {
    if (error instanceof AppError) {
      return reply.code(error.statusCode).send({
        error: { code: error.code, message: error.message, details: error.details },
      });
    }

    if (error instanceof ZodError) {
      return reply.code(400).send({
        error: {
          code: "VALIDATION_FAILED",
          message: "Request validation failed.",
          details: error.issues,
        },
      });
    }

    // @fastify/multipart signals an oversized part with this code.
    if ((error as { code?: string }).code === "FST_REQ_FILE_TOO_LARGE") {
      return reply.code(413).send({
        error: {
          code: "PAYLOAD_TOO_LARGE",
          message: "Uploaded file exceeds the maximum allowed size.",
        },
      });
    }

    request.log.error(error);
    return reply.code(500).send({
      error: { code: "INTERNAL_SERVER_ERROR", message: "An unexpected error occurred." },
    });
  });

  await fastify.register(healthRoutes);
  await fastify.register(uploadRoutes);
  await fastify.register(jobRoutes);
  await fastify.register(productRoutes);
  await fastify.register(labelRunRoutes);

  return fastify;
}
