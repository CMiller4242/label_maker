import { defineConfig } from "vitest/config";

// Mirrors .env.example / docker-compose.yml defaults. Integration tests
// need a real Postgres + Redis (see README.md "First run" for how to start
// them, either via `docker compose up -d` or local services).
export default defineConfig({
  test: {
    name: "integration",
    root: import.meta.dirname,
    environment: "node",
    testTimeout: 30_000,
    env: {
      DATABASE_URL:
        process.env.DATABASE_URL ??
        "postgresql://label_maker:label_maker@localhost:5432/label_maker?schema=public",
      REDIS_URL: process.env.REDIS_URL ?? "redis://localhost:6379",
      STORAGE_UPLOADS_DIR: process.env.STORAGE_UPLOADS_DIR ?? "./storage/uploads",
      STORAGE_ARTIFACTS_DIR: process.env.STORAGE_ARTIFACTS_DIR ?? "./storage/artifacts",
      MAX_UPLOAD_BYTES: process.env.MAX_UPLOAD_BYTES ?? "52428800",
      LOG_LEVEL: "silent",
      NODE_ENV: "test",
    },
  },
});
