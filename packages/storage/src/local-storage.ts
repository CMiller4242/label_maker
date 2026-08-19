import { createHash, randomUUID } from "node:crypto";
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

export class InvalidStorageKeyError extends Error {
  readonly code = "INVALID_STORAGE_KEY";
}

/** Bucket subdirectories this storage service writes into. */
export type StorageBucket = "uploads" | "artifacts";

export interface SavedFile {
  /** Opaque key clients use to later read this file back (never a user-supplied path). */
  storageKey: string;
  byteSize: number;
  sha256: string;
}

export interface StorageRoots {
  uploadsDir: string;
  artifactsDir: string;
}

const SAFE_EXTENSION_PATTERN = /^[a-zA-Z0-9]{1,10}$/;
// storageKey format: "<bucket>/<uuid>.<ext>" - fully server-generated, never
// derived from a user-supplied filename, so there is nothing for a caller
// to path-traverse with.
const STORAGE_KEY_PATTERN = /^(uploads|artifacts)\/[0-9a-f-]{36}\.[a-zA-Z0-9]{1,10}$/;

/**
 * Local-disk implementation of file storage for development. Files are
 * never addressed by user-supplied filenames or paths: save() generates a
 * random UUID-based key server-side, and read()/getAbsolutePath() validate
 * that any incoming storageKey matches the exact generated shape before
 * touching the filesystem, which rules out path traversal (`../..`, absolute
 * paths, symlink-like tricks via unexpected characters) by construction.
 */
export class LocalStorageService {
  private readonly roots: StorageRoots;

  constructor(roots: StorageRoots) {
    this.roots = roots;
  }

  async ensureDirectories(): Promise<void> {
    await mkdir(this.roots.uploadsDir, { recursive: true });
    await mkdir(this.roots.artifactsDir, { recursive: true });
  }

  private bucketDir(bucket: StorageBucket): string {
    return bucket === "uploads" ? this.roots.uploadsDir : this.roots.artifactsDir;
  }

  private validateStorageKey(storageKey: string): { bucket: StorageBucket; filename: string } {
    if (!STORAGE_KEY_PATTERN.test(storageKey)) {
      throw new InvalidStorageKeyError(
        `Rejected storage key with unexpected shape: "${storageKey}"`,
      );
    }
    const [bucket, filename] = storageKey.split("/") as [StorageBucket, string];
    return { bucket, filename };
  }

  async save(bucket: StorageBucket, data: Buffer, extension: string): Promise<SavedFile> {
    if (!SAFE_EXTENSION_PATTERN.test(extension)) {
      throw new InvalidStorageKeyError(`Rejected unsafe extension: "${extension}"`);
    }

    const filename = `${randomUUID()}.${extension.toLowerCase()}`;
    const storageKey = `${bucket}/${filename}`;
    const absolutePath = path.join(this.bucketDir(bucket), filename);

    await mkdir(path.dirname(absolutePath), { recursive: true });
    await writeFile(absolutePath, data);

    const sha256 = createHash("sha256").update(data).digest("hex");

    return { storageKey, byteSize: data.byteLength, sha256 };
  }

  async read(storageKey: string): Promise<Buffer> {
    const { bucket, filename } = this.validateStorageKey(storageKey);
    const absolutePath = path.join(this.bucketDir(bucket), filename);
    return readFile(absolutePath);
  }

  getAbsolutePath(storageKey: string): string {
    const { bucket, filename } = this.validateStorageKey(storageKey);
    return path.join(this.bucketDir(bucket), filename);
  }
}
