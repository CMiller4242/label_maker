import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

/**
 * Dev-server proxy: apps/api's Fastify routes are mounted at the root
 * (e.g. "/uploads", not "/api/uploads" - see apps/api/src/app.ts), so every
 * frontend request goes to "/api/*" and gets rewritten to "/*" before
 * reaching the API. This means the API needs no CORS configuration for
 * local development, and the frontend never hardcodes the API's origin.
 */
const API_PROXY_TARGET = process.env["VITE_API_PROXY_TARGET"] ?? "http://localhost:3000";

export default defineConfig({
  plugins: [react()],
  server: {
    port: 5173,
    proxy: {
      "/api": {
        target: API_PROXY_TARGET,
        changeOrigin: true,
        rewrite: (path) => path.replace(/^\/api/, ""),
      },
    },
  },
});
