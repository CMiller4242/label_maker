import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    name: "ingestion",
    root: import.meta.dirname,
    environment: "node",
  },
});
