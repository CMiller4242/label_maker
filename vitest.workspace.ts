import { defineWorkspace } from "vitest/config";

export default defineWorkspace([
  "tests/ingestion/vitest.config.ts",
  "tests/label-layout/vitest.config.ts",
  "tests/integration/vitest.config.ts",
]);
