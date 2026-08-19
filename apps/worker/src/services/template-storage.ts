import path from "node:path";
import { TemplateFileStorage } from "@label-maker/storage";

const templatesRootDir = process.env.TEMPLATES_ROOT_DIR ?? "./fixtures/label-templates";

export const templateStorage = new TemplateFileStorage(path.resolve(templatesRootDir));
