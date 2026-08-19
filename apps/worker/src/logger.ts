import pino from "pino";

export const logger = pino({
  name: "label-maker-worker",
  level: process.env.LOG_LEVEL ?? "info",
});
