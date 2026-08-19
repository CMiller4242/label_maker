import { z } from "zod";
import { PROCESSING_JOB_STATUSES, PROCESSING_JOB_TYPES } from "../constants/index.js";

export const processingJobTypeSchema = z.enum(PROCESSING_JOB_TYPES);
export type ProcessingJobType = z.infer<typeof processingJobTypeSchema>;

export const processingJobStatusSchema = z.enum(PROCESSING_JOB_STATUSES);
export type ProcessingJobStatus = z.infer<typeof processingJobStatusSchema>;

export const processingJobSchema = z.object({
  id: z.string().uuid(),
  sourceDocumentId: z.string().uuid(),
  jobType: processingJobTypeSchema,
  status: processingJobStatusSchema,
  progressCurrent: z.number().int().nonnegative(),
  progressTotal: z.number().int().nonnegative(),
  errorCode: z.string().nullable(),
  errorMessage: z.string().nullable(),
  createdAt: z.coerce.date(),
  startedAt: z.coerce.date().nullable(),
  completedAt: z.coerce.date().nullable(),
});
export type ProcessingJob = z.infer<typeof processingJobSchema>;

/** BullMQ job payload for INGEST_DOCUMENT jobs. */
export const ingestDocumentJobPayloadSchema = z.object({
  processingJobId: z.string().uuid(),
  sourceDocumentId: z.string().uuid(),
});
export type IngestDocumentJobPayload = z.infer<typeof ingestDocumentJobPayloadSchema>;

/** BullMQ job payload for GENERATE_LABEL_RUN jobs. */
export const generateLabelRunJobPayloadSchema = z.object({
  processingJobId: z.string().uuid(),
  labelRunId: z.string().uuid(),
});
export type GenerateLabelRunJobPayload = z.infer<typeof generateLabelRunJobPayloadSchema>;
