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
import { renderDebugJsonArtifact } from "@label-maker/docx-renderer";
import type { LabelTemplate as SharedLabelTemplate } from "@label-maker/shared";
import type { LocalStorageService } from "@label-maker/storage";
import { NotFoundError, ValidationFailedError } from "../lib/errors.js";

export interface CreateLabelRunInput {
  sourceDocumentId: string;
  labelTemplateId: string;
  copiesPerProduct: number;
}

export interface CreateLabelRunResult {
  labelRun: {
    id: string;
    sourceDocumentId: string;
    labelTemplateId: string;
    copiesPerProduct: number;
    status: string;
    placementPlanJson: unknown;
    generatedArtifactStorageKey: string | null;
    createdAt: Date;
    updatedAt: Date;
  };
  debugArtifact: { storageKey: string; byteSize: number };
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

/**
 * Builds a placement plan and debug JSON artifact synchronously. Placement
 * generation is pure, deterministic, in-memory computation over already-
 * ingested Product rows (no file parsing), so unlike document ingestion it
 * does not need to be queued to the worker to stay responsive - the
 * GENERATE_LABEL_RUN ProcessingJob row is still recorded (started+completed
 * within this call) for a consistent audit trail across job types.
 */
export async function createLabelRun(
  prisma: PrismaClient,
  storage: LocalStorageService,
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

  const processingJob = await prisma.processingJob.create({
    data: {
      sourceDocumentId: input.sourceDocumentId,
      jobType: "GENERATE_LABEL_RUN",
      status: "PROCESSING",
      startedAt: new Date(),
    },
  });

  let placementResult;
  try {
    placementResult = buildPlacements(placementInputs, geometry, input.copiesPerProduct);
    validatePlacements(placementResult.placements, geometry);
  } catch (error) {
    await prisma.processingJob.update({
      where: { id: processingJob.id },
      data: {
        status: "FAILED",
        errorCode:
          error instanceof InvalidTemplateGeometryError ||
          error instanceof InvalidCopiesPerProductError ||
          error instanceof MissingRequiredProductDataError
            ? error.code
            : "PLACEMENT_FAILED",
        errorMessage: error instanceof Error ? error.message : String(error),
        completedAt: new Date(),
      },
    });
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

  const artifact = renderDebugJsonArtifact(placementPlanForRenderer, templateForRenderer);
  const saved = await storage.save("artifacts", artifact.buffer, artifact.fileExtension);

  const labelRun = await prisma.labelRun.create({
    data: {
      sourceDocumentId: input.sourceDocumentId,
      labelTemplateId: input.labelTemplateId,
      copiesPerProduct: input.copiesPerProduct,
      status: "GENERATED",
      placementPlanJson: placementResult as unknown as Prisma.InputJsonValue,
      generatedArtifactStorageKey: saved.storageKey,
    },
  });

  await prisma.processingJob.update({
    where: { id: processingJob.id },
    data: { status: "COMPLETED", progressCurrent: 1, progressTotal: 1, completedAt: new Date() },
  });

  return {
    labelRun,
    debugArtifact: { storageKey: saved.storageKey, byteSize: saved.byteSize },
  };
}
