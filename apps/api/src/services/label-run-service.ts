import { createHash } from "node:crypto";
import type { PrismaClient, Prisma } from "@label-maker/database";
import {
  AVERY_5155_GEOMETRY,
  buildPlacements,
  InvalidCopiesPerProductError,
  InvalidTemplateGeometryError,
  MissingRequiredProductDataError,
  validatePlacements,
  type LabelTemplateGeometry,
  type PlacementInputProduct,
} from "@label-maker/label-layout";
import {
  InvalidDocxPackageError,
  InvalidFixedGridTemplateError,
  MissingLabelTextStyleConfigError,
  UnsupportedFloatingTemplateError,
  renderDebugJsonArtifact,
  renderFixedGridDocx,
} from "@label-maker/docx-renderer";
import type { LabelRun, LabelTemplate as SharedLabelTemplate } from "@label-maker/shared";
import type { LocalStorageService, TemplateFileStorage } from "@label-maker/storage";
import {
  NotFoundError,
  TemplateUnavailableError,
  UnsupportedRenderingModeError,
  ValidationFailedError,
} from "../lib/errors.js";

export interface CreateLabelRunInput {
  sourceDocumentId: string;
  labelTemplateId: string;
  copiesPerProduct: number;
  includeDebugArtifact: boolean;
}

export interface GeneratedArtifactInfo {
  storageKey: string;
  byteSize: number;
  sha256: string;
  mimeType: string;
}

export interface CreateLabelRunResult {
  labelRun: LabelRun;
  artifact: GeneratedArtifactInfo;
  debugArtifact?: Omit<GeneratedArtifactInfo, "mimeType">;
}

function geometryFromTemplateRow(row: {
  id: string;
  columns: number;
  rows: number;
  labelsPerSheet: number;
}): LabelTemplateGeometry {
  if (row.id === AVERY_5155_GEOMETRY.id) return AVERY_5155_GEOMETRY;
  return { id: row.id, columns: row.columns, rows: row.rows, labelsPerSheet: row.labelsPerSheet };
}

function sha256Hex(buffer: Buffer): string {
  return createHash("sha256").update(buffer).digest("hex");
}

/**
 * Builds a placement plan and, for a validated FIXED_GRID template, a real
 * DOCX artifact. Placement generation is pure, deterministic, in-memory
 * computation over already-ingested Product rows (no file parsing), so
 * unlike document ingestion it does not need to be queued to the worker to
 * stay responsive - the GENERATE_LABEL_RUN ProcessingJob row is still
 * recorded (started+completed within this call) for a consistent audit
 * trail across job types.
 */
export async function createLabelRun(
  prisma: PrismaClient,
  storage: LocalStorageService,
  templates: TemplateFileStorage,
  input: CreateLabelRunInput,
): Promise<CreateLabelRunResult> {
  const sourceDocument = await prisma.sourceDocument.findUnique({
    where: { id: input.sourceDocumentId },
  });
  if (!sourceDocument) {
    throw new NotFoundError(`SourceDocument "${input.sourceDocumentId}" not found.`);
  }

  const labelTemplate = await prisma.labelTemplate.findUnique({
    where: { id: input.labelTemplateId },
  });
  if (!labelTemplate) {
    throw new NotFoundError(`LabelTemplate "${input.labelTemplateId}" not found.`);
  }

  const processingJob = await prisma.processingJob.create({
    data: {
      sourceDocumentId: input.sourceDocumentId,
      jobType: "GENERATE_LABEL_RUN",
      status: "PROCESSING",
      startedAt: new Date(),
    },
  });

  async function failJob(code: string, message: string): Promise<void> {
    await prisma.processingJob.update({
      where: { id: processingJob.id },
      data: { status: "FAILED", errorCode: code, errorMessage: message, completedAt: new Date() },
    });
  }

  if (labelTemplate.renderingMode === "FLOATING") {
    const error = new UnsupportedFloatingTemplateError(labelTemplate.id);
    await failJob(error.code, error.message);
    throw new UnsupportedRenderingModeError(error.message);
  }

  const products = await prisma.product.findMany({
    where: {
      sourceDocumentId: input.sourceDocumentId,
      include: true,
      status: { in: ["APPROVED", "AUTO_ACCEPTED"] },
    },
  });

  const placementInputs: PlacementInputProduct[] = products.map((p) => ({
    id: p.id,
    sku: p.sku,
    description: p.description,
    priceCents: p.priceCents,
    include: p.include,
  }));

  const geometry = geometryFromTemplateRow(labelTemplate);

  let placementResult;
  try {
    placementResult = buildPlacements(placementInputs, geometry, input.copiesPerProduct);
    validatePlacements(placementResult.placements, geometry);
  } catch (error) {
    const code =
      error instanceof InvalidTemplateGeometryError ||
      error instanceof InvalidCopiesPerProductError ||
      error instanceof MissingRequiredProductDataError
        ? error.code
        : "PLACEMENT_FAILED";
    await failJob(code, error instanceof Error ? error.message : String(error));
    if (
      error instanceof InvalidTemplateGeometryError ||
      error instanceof InvalidCopiesPerProductError ||
      error instanceof MissingRequiredProductDataError
    ) {
      throw new ValidationFailedError(error.message);
    }
    throw error;
  }

  const templateForRenderer: SharedLabelTemplate = {
    id: labelTemplate.id,
    displayName: labelTemplate.displayName,
    renderingMode: labelTemplate.renderingMode,
    columns: labelTemplate.columns,
    rows: labelTemplate.rows,
    labelsPerSheet: labelTemplate.labelsPerSheet,
    templateStorageKey: labelTemplate.templateStorageKey,
    templateVersion: labelTemplate.templateVersion,
    sourceTemplateSha256: labelTemplate.sourceTemplateSha256,
    configJson: labelTemplate.configJson as Record<string, unknown>,
    isPreset: labelTemplate.isPreset,
    createdAt: labelTemplate.createdAt,
    updatedAt: labelTemplate.updatedAt,
  };

  const placementPlanForRenderer = {
    labelTemplateId: placementResult.labelTemplateId,
    copiesPerProduct: placementResult.copiesPerProduct,
    totalSheets: placementResult.totalSheets,
    totalPlacements: placementResult.totalPlacements,
    placements: placementResult.placements,
  };

  if (!labelTemplate.templateStorageKey) {
    const message = `LabelTemplate "${labelTemplate.id}" has no templateStorageKey configured; no source DOCX to render from.`;
    await failJob("TEMPLATE_UNAVAILABLE", message);
    throw new TemplateUnavailableError(message);
  }

  let templateBuffer: Buffer;
  try {
    templateBuffer = await templates.read(labelTemplate.templateStorageKey);
  } catch (error) {
    const message = `LabelTemplate "${labelTemplate.id}" source file at "${labelTemplate.templateStorageKey}" could not be read: ${error instanceof Error ? error.message : String(error)}`;
    await failJob("TEMPLATE_UNAVAILABLE", message);
    throw new TemplateUnavailableError(message);
  }

  let docxResult;
  try {
    docxResult = await renderFixedGridDocx(
      templateBuffer,
      templateForRenderer,
      placementPlanForRenderer,
    );
  } catch (error) {
    if (
      error instanceof InvalidDocxPackageError ||
      error instanceof InvalidFixedGridTemplateError ||
      error instanceof MissingLabelTextStyleConfigError
    ) {
      await failJob(error.code, error.message);
      throw new TemplateUnavailableError(error.message);
    }
    await failJob("DOCX_RENDER_FAILED", error instanceof Error ? error.message : String(error));
    throw error;
  }

  const savedDocx = await storage.save("artifacts", docxResult.buffer, docxResult.fileExtension);
  const generatedArtifactSha256 = sha256Hex(docxResult.buffer);

  let debugArtifactInfo: Omit<GeneratedArtifactInfo, "mimeType"> | undefined;
  if (input.includeDebugArtifact) {
    const debugResult = renderDebugJsonArtifact(placementPlanForRenderer, templateForRenderer);
    const savedDebug = await storage.save(
      "artifacts",
      debugResult.buffer,
      debugResult.fileExtension,
    );
    debugArtifactInfo = {
      storageKey: savedDebug.storageKey,
      byteSize: savedDebug.byteSize,
      sha256: sha256Hex(debugResult.buffer),
    };
  }

  const sheetCount = placementResult.totalSheets;
  const filledSlotCount = placementResult.totalPlacements;
  const emptySlotCount = sheetCount * labelTemplate.labelsPerSheet - filledSlotCount;

  const labelRun = await prisma.labelRun.create({
    data: {
      sourceDocumentId: input.sourceDocumentId,
      labelTemplateId: input.labelTemplateId,
      copiesPerProduct: input.copiesPerProduct,
      status: "GENERATED",
      placementPlanJson: placementResult as unknown as Prisma.InputJsonValue,
      generatedArtifactStorageKey: savedDocx.storageKey,
      generatedArtifactSha256,
      templateVersion: labelTemplate.templateVersion,
      templateSha256: labelTemplate.sourceTemplateSha256,
      sheetCount,
      filledSlotCount,
      emptySlotCount,
      metadataJson: { rendererUsed: "FIXED_GRID_DOCX" } as unknown as Prisma.InputJsonValue,
    },
  });

  await prisma.processingJob.update({
    where: { id: processingJob.id },
    data: { status: "COMPLETED", progressCurrent: 1, progressTotal: 1, completedAt: new Date() },
  });

  return {
    labelRun: labelRun as unknown as LabelRun,
    artifact: {
      storageKey: savedDocx.storageKey,
      byteSize: savedDocx.byteSize,
      sha256: generatedArtifactSha256,
      mimeType: docxResult.mimeType,
    },
    ...(debugArtifactInfo ? { debugArtifact: debugArtifactInfo } : {}),
  };
}

export interface LabelRunArtifactDownload {
  buffer: Buffer;
  mimeType: string;
  filename: string;
}

const MIME_TYPE_BY_EXTENSION: Record<string, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  json: "application/json",
};

/** Loads a LabelRun's generated artifact bytes for download, never exposing the underlying filesystem path. */
export async function getLabelRunArtifactForDownload(
  prisma: PrismaClient,
  storage: LocalStorageService,
  labelRunId: string,
): Promise<LabelRunArtifactDownload> {
  const labelRun = await prisma.labelRun.findUnique({ where: { id: labelRunId } });
  if (!labelRun) {
    throw new NotFoundError(`LabelRun "${labelRunId}" not found.`);
  }
  if (!labelRun.generatedArtifactStorageKey) {
    throw new NotFoundError(`LabelRun "${labelRunId}" has no generated artifact yet.`);
  }

  const buffer = await storage.read(labelRun.generatedArtifactStorageKey);
  const extension = labelRun.generatedArtifactStorageKey.split(".").pop() ?? "";
  const mimeType = MIME_TYPE_BY_EXTENSION[extension] ?? "application/octet-stream";

  // Safe, server-derived filename - never the raw storage key/path.
  const filename = `label-run-${labelRun.id}.${extension || "bin"}`;

  return { buffer, mimeType, filename };
}
