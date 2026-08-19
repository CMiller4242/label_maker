import type { PrismaClient, ProcessingJob } from "@label-maker/database";
import { NotFoundError } from "../lib/errors.js";

export async function getJobById(prisma: PrismaClient, jobId: string): Promise<ProcessingJob> {
  const job = await prisma.processingJob.findUnique({ where: { id: jobId } });
  if (!job) {
    throw new NotFoundError(`ProcessingJob "${jobId}" not found.`);
  }
  return job;
}
