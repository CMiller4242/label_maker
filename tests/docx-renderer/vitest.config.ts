import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    name: "docx-renderer",
    root: import.meta.dirname,
    environment: "node",
  },
});
