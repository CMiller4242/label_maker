import { readFile } from "node:fs/promises";
import path from "node:path";

export class InvalidTemplateKeyError extends Error {
  readonly code = "INVALID_TEMPLATE_KEY";
}

/**
 * Read-only storage abstraction for source label templates (e.g.
 * `fixtures/label-templates/avery-5155/original.docx`). Unlike
 * LocalStorageService (which only ever addresses server-generated UUID
 * filenames for user uploads/artifacts), template keys are developer-chosen
 * relative paths under a configured root - e.g. LabelTemplate.templateStorageKey
 * = "avery-5155/original.docx". Path traversal is still rejected: the
 * resolved path must stay within templatesRootDir.
 */
export class TemplateFileStorage {
  private readonly templatesRootDir: string;

  constructor(templatesRootDir: string) {
    this.templatesRootDir = path.resolve(templatesRootDir);
  }

  private resolveKey(templateStorageKey: string): string {
    const resolved = path.resolve(this.templatesRootDir, templateStorageKey);
    const rootWithSep = this.templatesRootDir.endsWith(path.sep)
      ? this.templatesRootDir
      : `${this.templatesRootDir}${path.sep}`;
    if (resolved !== this.templatesRootDir && !resolved.startsWith(rootWithSep)) {
      throw new InvalidTemplateKeyError(
        `Rejected template key that escapes the templates root: "${templateStorageKey}"`,
      );
    }
    return resolved;
  }

  async read(templateStorageKey: string): Promise<Buffer> {
    return readFile(this.resolveKey(templateStorageKey));
  }

  getAbsolutePath(templateStorageKey: string): string {
    return this.resolveKey(templateStorageKey);
  }
}
