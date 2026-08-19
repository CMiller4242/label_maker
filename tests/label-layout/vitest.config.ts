import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    name: "label-layout",
    root: import.meta.dirname,
    environment: "node",
  },
});
