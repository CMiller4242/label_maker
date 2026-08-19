import path from "node:path";
import { LocalStorageService } from "@label-maker/storage";

const uploadsDir = process.env.STORAGE_UPLOADS_DIR ?? "./storage/uploads";
const artifactsDir = process.env.STORAGE_ARTIFACTS_DIR ?? "./storage/artifacts";

export const storageService = new LocalStorageService({
  uploadsDir: path.resolve(uploadsDir),
  artifactsDir: path.resolve(artifactsDir),
});
