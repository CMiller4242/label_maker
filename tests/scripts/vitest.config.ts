import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    name: "scripts",
    root: import.meta.dirname,
    environment: "node",
  },
});
