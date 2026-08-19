import path from "node:path";
import { z } from "zod";
import { DEFAULT_MAX_UPLOAD_BYTES } from "@label-maker/shared";

const envSchema = z.object({
  DATABASE_URL: z.string().min(1),
  REDIS_URL: z.string().min(1).default("redis://localhost:6379"),
  API_PORT: z.coerce.number().int().positive().default(3000),
  API_HOST: z.string().min(1).default("0.0.0.0"),
  STORAGE_UPLOADS_DIR: z.string().min(1).default("./storage/uploads"),
  STORAGE_ARTIFACTS_DIR: z.string().min(1).default("./storage/artifacts"),
  TEMPLATES_ROOT_DIR: z.string().min(1).default("./fixtures/label-templates"),
  MAX_UPLOAD_BYTES: z.coerce.number().int().positive().default(DEFAULT_MAX_UPLOAD_BYTES),
  LOG_LEVEL: z.string().min(1).default("info"),
  NODE_ENV: z.string().min(1).default("development"),
});

export interface AppConfig {
  databaseUrl: string;
  redisUrl: string;
  port: number;
  host: string;
  storageUploadsDir: string;
  storageArtifactsDir: string;
  templatesRootDir: string;
  maxUploadBytes: number;
  logLevel: string;
  nodeEnv: string;
}

export function loadConfig(env: NodeJS.ProcessEnv = process.env): AppConfig {
  const parsed = envSchema.parse(env);
  return {
    databaseUrl: parsed.DATABASE_URL,
    redisUrl: parsed.REDIS_URL,
    port: parsed.API_PORT,
    host: parsed.API_HOST,
    storageUploadsDir: path.resolve(parsed.STORAGE_UPLOADS_DIR),
    storageArtifactsDir: path.resolve(parsed.STORAGE_ARTIFACTS_DIR),
    templatesRootDir: path.resolve(parsed.TEMPLATES_ROOT_DIR),
    maxUploadBytes: parsed.MAX_UPLOAD_BYTES,
    logLevel: parsed.LOG_LEVEL,
    nodeEnv: parsed.NODE_ENV,
  };
}
