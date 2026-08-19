import type { Prisma, PrismaClient } from "@label-maker/database";
import {
  AVERY_5155_GEOMETRY,
  buildPlacements,
  validatePlacements,
  type LabelTemplateGeometry,
  type PlacementInputProduct,
} from "@label-maker/label-layout";
import { renderDebugJsonArtifact } from "@label-maker/docx-renderer";
import type { GenerateLabelRunJobPayload, LabelTemplate } from "@label-maker/shared";
import { storageService } from "../services/storage-service.js";
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

/**
 * Processes a GENERATE_LABEL_RUN job: builds a deterministic placement plan
 * from APPROVED/AUTO_ACCEPTED included products, validates it, renders the
 * (stub) debug JSON artifact, and persists both onto the LabelRun row.
 *
 * Re-running this job for the same labelRunId is deterministic in its
 * placement output, but the storage layer generates a fresh artifact key
 * per run (see @label-maker/storage) - a prior artifact file is left
 * orphaned rather than overwritten. Acceptable for this scaffold; a real
 * idempotency key (e.g. artifact key derived from labelRunId + plan hash)
 * is a TODO once real DOCX rendering exists.
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

    const artifact = renderDebugJsonArtifact(placementPlanForRenderer, templateForRenderer);
    const saved = await storageService.save("artifacts", artifact.buffer, artifact.fileExtension);

    await prisma.labelRun.update({
      where: { id: labelRun.id },
      data: {
        status: "GENERATED",
        placementPlanJson: placementResult as unknown as Prisma.InputJsonValue,
        generatedArtifactStorageKey: saved.storageKey,
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
    const message = error instanceof Error ? error.message : String(error);
    logger.error(
      { jobId: job.id, labelRunId: labelRun.id, err: message },
      "build-label-run: failed",
    );
    await prisma.labelRun.update({ where: { id: labelRun.id }, data: { status: "FAILED" } });
    await prisma.processingJob.update({
      where: { id: job.id },
      data: {
        status: "FAILED",
        errorCode: "LABEL_RUN_FAILED",
        errorMessage: message,
        completedAt: new Date(),
      },
    });
    throw error;
  }
}
