import type { AppConfig } from "../config.js";

/** Fastify-compatible pino logger options, derived from AppConfig. */
export function loggerOptions(config: Pick<AppConfig, "logLevel">): { level: string } {
  return { level: config.logLevel };
}
