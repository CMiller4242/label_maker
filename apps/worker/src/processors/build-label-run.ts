import { createHash } from "node:crypto";
import type { Prisma, PrismaClient } from "@label-maker/database";
import {
  AVERY_5155_GEOMETRY,
  buildPlacements,
  validatePlacements,
  type LabelTemplateGeometry,
  type PlacementInputProduct,
} from "@label-maker/label-layout";
import {
  InvalidDocxPackageError,
  InvalidFixedGridTemplateError,
  MissingLabelTextStyleConfigError,
  UnsupportedFloatingTemplateError,
  renderFixedGridDocx,
} from "@label-maker/docx-renderer";
import type { GenerateLabelRunJobPayload, LabelTemplate } from "@label-maker/shared";
import { storageService } from "../services/storage-service.js";
import { templateStorage } from "../services/template-storage.js";
import { logger } from "../logger.js";

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

class UnprocessableTemplateError extends Error {
  readonly code: string;

  constructor(code: string, message: string) {
    super(message);
    this.code = code;
  }
}

/**
 * Processes a GENERATE_LABEL_RUN job: builds a deterministic placement plan
 * from APPROVED/AUTO_ACCEPTED included products, validates it, and (for a
 * validated FIXED_GRID template) renders a real DOCX artifact. Mirrors
 * apps/api's label-run-service.ts synchronous path - this queued processor
 * exists for future async scaling (e.g. very large runs), not because
 * placement/DOCX generation is normally slow.
 *
 * Re-running this job for the same labelRunId is deterministic in its
 * placement output, but the storage layer generates a fresh artifact key
 * per run (see @label-maker/storage) - a prior artifact file is left
 * orphaned rather than overwritten. Acceptable for this milestone; a real
 * idempotency key (e.g. artifact key derived from labelRunId + plan hash)
 * is a TODO.
 */
export async function processBuildLabelRunJob(
  prisma: PrismaClient,
  payload: GenerateLabelRunJobPayload,
): Promise<void> {
  const startedAt = new Date();
  const job = await prisma.processingJob.update({
    where: { id: payload.processingJobId },
    data: { status: "PROCESSING", startedAt },
  });

  const labelRun = await prisma.labelRun.findUniqueOrThrow({
    where: { id: payload.labelRunId },
    include: { labelTemplate: true },
  });

  logger.info(
    { jobId: job.id, labelRunId: labelRun.id, sourceDocumentId: labelRun.sourceDocumentId },
    "build-label-run: starting",
  );

  try {
    if (labelRun.labelTemplate.renderingMode === "FLOATING") {
      throw new UnsupportedFloatingTemplateError(labelRun.labelTemplate.id);
    }

    const products = await prisma.product.findMany({
      where: {
        sourceDocumentId: labelRun.sourceDocumentId,
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

    const geometry = geometryFromTemplateRow(labelRun.labelTemplate);
    const placementResult = buildPlacements(placementInputs, geometry, labelRun.copiesPerProduct);
    validatePlacements(placementResult.placements, geometry);

    const templateForRenderer: LabelTemplate = {
      id: labelRun.labelTemplate.id,
      displayName: labelRun.labelTemplate.displayName,
      renderingMode: labelRun.labelTemplate.renderingMode,
      columns: labelRun.labelTemplate.columns,
      rows: labelRun.labelTemplate.rows,
      labelsPerSheet: labelRun.labelTemplate.labelsPerSheet,
      templateStorageKey: labelRun.labelTemplate.templateStorageKey,
      templateVersion: labelRun.labelTemplate.templateVersion,
      sourceTemplateSha256: labelRun.labelTemplate.sourceTemplateSha256,
      configJson: labelRun.labelTemplate.configJson as Record<string, unknown>,
      isPreset: labelRun.labelTemplate.isPreset,
      createdAt: labelRun.labelTemplate.createdAt,
      updatedAt: labelRun.labelTemplate.updatedAt,
    };

    const placementPlanForRenderer = {
      labelTemplateId: placementResult.labelTemplateId,
      copiesPerProduct: placementResult.copiesPerProduct,
      totalSheets: placementResult.totalSheets,
      totalPlacements: placementResult.totalPlacements,
      placements: placementResult.placements,
    };

    if (!labelRun.labelTemplate.templateStorageKey) {
      throw new UnprocessableTemplateError(
        "TEMPLATE_UNAVAILABLE",
        `LabelTemplate "${labelRun.labelTemplate.id}" has no templateStorageKey configured.`,
      );
    }

    let templateBuffer: Buffer;
    try {
      templateBuffer = await templateStorage.read(labelRun.labelTemplate.templateStorageKey);
    } catch (error) {
      throw new UnprocessableTemplateError(
        "TEMPLATE_UNAVAILABLE",
        `LabelTemplate "${labelRun.labelTemplate.id}" source file at "${labelRun.labelTemplate.templateStorageKey}" could not be read: ${error instanceof Error ? error.message : String(error)}`,
      );
    }

    const docxResult = await renderFixedGridDocx(
      templateBuffer,
      templateForRenderer,
      placementPlanForRenderer,
    );
    const saved = await storageService.save(
      "artifacts",
      docxResult.buffer,
      docxResult.fileExtension,
    );
    const generatedArtifactSha256 = sha256Hex(docxResult.buffer);

    const sheetCount = placementResult.totalSheets;
    const filledSlotCount = placementResult.totalPlacements;
    const emptySlotCount = sheetCount * labelRun.labelTemplate.labelsPerSheet - filledSlotCount;

    await prisma.labelRun.update({
      where: { id: labelRun.id },
      data: {
        status: "GENERATED",
        placementPlanJson: placementResult as unknown as Prisma.InputJsonValue,
        generatedArtifactStorageKey: saved.storageKey,
        generatedArtifactSha256,
        templateVersion: labelRun.labelTemplate.templateVersion,
        templateSha256: labelRun.labelTemplate.sourceTemplateSha256,
        sheetCount,
        filledSlotCount,
        emptySlotCount,
        metadataJson: { rendererUsed: "FIXED_GRID_DOCX" } as unknown as Prisma.InputJsonValue,
      },
    });

    const completedAt = new Date();
    await prisma.processingJob.update({
      where: { id: job.id },
      data: { status: "COMPLETED", progressCurrent: 1, progressTotal: 1, completedAt },
    });

    logger.info(
      {
        jobId: job.id,
        labelRunId: labelRun.id,
        durationMs: completedAt.getTime() - startedAt.getTime(),
        totalSheets: placementResult.totalSheets,
        totalPlacements: placementResult.totalPlacements,
        artifactStorageKey: saved.storageKey,
      },
      "build-label-run: completed",
    );
  } catch (error) {
    const code =
      error instanceof UnsupportedFloatingTemplateError
        ? error.code
        : error instanceof UnprocessableTemplateError
          ? error.code
          : error instanceof InvalidDocxPackageError ||
              error instanceof InvalidFixedGridTemplateError ||
              error instanceof MissingLabelTextStyleConfigError
            ? error.code
            : "LABEL_RUN_FAILED";
    const message = error instanceof Error ? error.message : String(error);
    logger.error(
      { jobId: job.id, labelRunId: labelRun.id, err: message },
      "build-label-run: failed",
    );
    await prisma.labelRun.update({ where: { id: labelRun.id }, data: { status: "FAILED" } });
    await prisma.processingJob.update({
      where: { id: job.id },
      data: { status: "FAILED", errorCode: code, errorMessage: message, completedAt: new Date() },
    });
    throw error;
  }
}
