import fp from "fastify-plugin";
import type { FastifyInstance } from "fastify";
import { LocalStorageService, TemplateFileStorage } from "@label-maker/storage";
import type { AppConfig } from "../config.js";

declare module "fastify" {
  interface FastifyInstance {
    storage: LocalStorageService;
    templates: TemplateFileStorage;
    config: AppConfig;
  }
}

export default fp(async function storagePlugin(
  fastify: FastifyInstance,
  opts: { config: AppConfig },
): Promise<void> {
  const storage = new LocalStorageService({
    uploadsDir: opts.config.storageUploadsDir,
    artifactsDir: opts.config.storageArtifactsDir,
  });
  await storage.ensureDirectories();

  const templates = new TemplateFileStorage(opts.config.templatesRootDir);

  fastify.decorate("config", opts.config);
  fastify.decorate("storage", storage);
  fastify.decorate("templates", templates);
});
