import { Redis } from "ioredis";
import { QUEUE_NAMES } from "@label-maker/shared";

export { QUEUE_NAMES };

export function createRedisConnection(): Redis {
  const redisUrl = process.env.REDIS_URL ?? "redis://localhost:6379";
  return new Redis(redisUrl, { maxRetriesPerRequest: null });
}
