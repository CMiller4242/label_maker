import { Worker, type Job } from "bullmq";
import { getPrismaClient, disconnectPrismaClient } from "@label-maker/database";
import {
  generateLabelRunJobPayloadSchema,
  ingestDocumentJobPayloadSchema,
} from "@label-maker/shared";
import { createRedisConnection, QUEUE_NAMES } from "./queues.js";
import { processIngestDocumentJob } from "./processors/ingest-source-document.js";
import { processBuildLabelRunJob } from "./processors/build-label-run.js";
import { logger } from "./logger.js";

const prisma = getPrismaClient();
const connection = createRedisConnection();

const ingestDocumentWorker = new Worker(
  QUEUE_NAMES.INGEST_DOCUMENT,
  async (job: Job) => {
    const payload = ingestDocumentJobPayloadSchema.parse(job.data);
    await processIngestDocumentJob(prisma, payload);
  },
  { connection, concurrency: 2 },
);

const generateLabelRunWorker = new Worker(
  QUEUE_NAMES.GENERATE_LABEL_RUN,
  async (job: Job) => {
    const payload = generateLabelRunJobPayloadSchema.parse(job.data);
    await processBuildLabelRunJob(prisma, payload);
  },
  { connection, concurrency: 2 },
);

for (const worker of [ingestDocumentWorker, generateLabelRunWorker]) {
  worker.on("failed", (job, error) => {
    logger.error({ queue: worker.name, jobId: job?.id, err: error.message }, "job failed");
  });
  worker.on("completed", (job) => {
    logger.info({ queue: worker.name, jobId: job.id }, "job completed");
  });
}

logger.info(
  { queues: [QUEUE_NAMES.INGEST_DOCUMENT, QUEUE_NAMES.GENERATE_LABEL_RUN] },
  "worker started",
);

async function shutdown(): Promise<void> {
  logger.info("worker shutting down");
  await Promise.all([ingestDocumentWorker.close(), generateLabelRunWorker.close()]);
  await disconnectPrismaClient();
  await connection.quit();
  process.exit(0);
}

process.on("SIGTERM", () => void shutdown());
process.on("SIGINT", () => void shutdown());
