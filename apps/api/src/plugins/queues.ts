import fp from "fastify-plugin";
import type { FastifyInstance } from "fastify";
import { Queue } from "bullmq";
import { Redis } from "ioredis";
import { QUEUE_NAMES } from "@label-maker/shared";
import type { AppConfig } from "../config.js";

export interface AppQueues {
  ingestDocument: Queue;
  generateLabelRun: Queue;
}

declare module "fastify" {
  interface FastifyInstance {
    queues: AppQueues;
  }
}

export default fp(async function queuesPlugin(
  fastify: FastifyInstance,
  opts: { config: AppConfig },
): Promise<void> {
  const connection = new Redis(opts.config.redisUrl, { maxRetriesPerRequest: null });

  const queues: AppQueues = {
    ingestDocument: new Queue(QUEUE_NAMES.INGEST_DOCUMENT, { connection }),
    generateLabelRun: new Queue(QUEUE_NAMES.GENERATE_LABEL_RUN, { connection }),
  };

  fastify.decorate("queues", queues);
  fastify.addHook("onClose", async () => {
    await Promise.all([queues.ingestDocument.close(), queues.generateLabelRun.close()]);
    await connection.quit();
  });
});
