import { buildApp } from "./app.js";
import { loadConfig } from "./config.js";

const config = loadConfig();
const fastify = await buildApp(config);

try {
  await fastify.listen({ port: config.port, host: config.host });
} catch (error) {
  fastify.log.error(error);
  process.exit(1);
}

async function shutdown(): Promise<void> {
  await fastify.close();
  process.exit(0);
}

process.on("SIGTERM", () => void shutdown());
process.on("SIGINT", () => void shutdown());
