import fp from "fastify-plugin";
import type { FastifyInstance } from "fastify";
import { disconnectPrismaClient, getPrismaClient, type PrismaClient } from "@label-maker/database";

declare module "fastify" {
  interface FastifyInstance {
    prisma: PrismaClient;
  }
}

export default fp(async function prismaPlugin(fastify: FastifyInstance): Promise<void> {
  const prisma = getPrismaClient();
  fastify.decorate("prisma", prisma);
  fastify.addHook("onClose", async () => {
    await disconnectPrismaClient();
  });
});
