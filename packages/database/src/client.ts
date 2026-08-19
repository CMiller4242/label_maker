import { PrismaClient } from "@prisma/client";

export type { PrismaClient } from "@prisma/client";
export * from "@prisma/client";

let cachedClient: PrismaClient | undefined;

/**
 * Returns a process-wide singleton PrismaClient. Reused across API/worker
 * request handling so we don't exhaust Postgres connections under load.
 */
export function getPrismaClient(): PrismaClient {
  if (!cachedClient) {
    cachedClient = new PrismaClient();
  }
  return cachedClient;
}

export async function disconnectPrismaClient(): Promise<void> {
  if (cachedClient) {
    await cachedClient.$disconnect();
    cachedClient = undefined;
  }
}
