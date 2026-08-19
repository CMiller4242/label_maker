import type { PrismaClient } from "@label-maker/database";
import { NEEDS_REVIEW_CONFIDENCE_THRESHOLD, type IngestionCandidate } from "@label-maker/ingestion";
import type { IngestDocumentJobPayload } from "@label-maker/shared";
import { routeDocumentForExtraction } from "../services/document-router.js";
import { storageService } from "../services/storage-service.js";
import { logger } from "../logger.js";

function boundingBoxJson(candidate: IngestionCandidate): Record<string, number> | undefined {
  return candidate.boundingBox
    ? {
        x: candidate.boundingBox.x,
        y: candidate.boundingBox.y,
        width: candidate.boundingBox.width,
        height: candidate.boundingBox.height,
      }
    : undefined;
}

/**
 * Processes an INGEST_DOCUMENT job: loads the source file, extracts
 * candidate product data (never finished Product-only-from-raw-text
 * records), and persists SourcePage/ExtractionCandidate/Product rows.
 *
 * Idempotent per (sourceDocumentId, jobType=INGEST_DOCUMENT): any
 * previously-persisted SourcePage/Product rows for this document are
 * cleared before writing fresh ones, so re-running the same job (e.g.
 * after a BullMQ retry) converges to the same end state rather than
 * duplicating rows.
 */
export async function processIngestDocumentJob(
  prisma: PrismaClient,
  payload: IngestDocumentJobPayload,
): Promise<void> {
  const startedAt = new Date();
  const job = await prisma.processingJob.update({
    where: { id: payload.processingJobId },
    data: { status: "PROCESSING", startedAt },
  });

  const sourceDocument = await prisma.sourceDocument.findUniqueOrThrow({
    where: { id: payload.sourceDocumentId },
  });

  logger.info(
    { jobId: job.id, documentId: sourceDocument.id, sourceType: sourceDocument.extension },
    "ingest-source-document: starting",
  );

  try {
    const buffer = await storageService.read(sourceDocument.storageKey);
    const routed = await routeDocumentForExtraction(buffer, sourceDocument.extension);
    const { result } = routed;

    // Idempotency: clear any prior extraction output for this document
    // before writing the new pass. Products cascade-delete their
    // ExtractionCandidates are on SourcePage, not Product, so both must be
    // cleared explicitly.
    await prisma.product.deleteMany({ where: { sourceDocumentId: sourceDocument.id } });
    await prisma.sourcePage.deleteMany({ where: { sourceDocumentId: sourceDocument.id } });

    let pageCount: number | undefined;
    let sheetCount: number | undefined;
    let productsCreated = 0;

    if (routed.kind === "pdf") {
      pageCount = routed.result.pageCount;
      for (const page of routed.result.pages) {
        const sourcePage = await prisma.sourcePage.create({
          data: {
            sourceDocumentId: sourceDocument.id,
            pageNumber: page.pageNumber,
            extractionMethod: "NATIVE_TEXT",
            rawText: page.rawText,
            extractionConfidence: page.needsReview ? 0.2 : 0.6,
            needsReview: page.needsReview,
          },
        });

        if (page.candidates.length > 0) {
          await prisma.extractionCandidate.createMany({
            data: page.candidates.map((candidate) => ({
              sourcePageId: sourcePage.id,
              field: candidate.field,
              rawValue: candidate.rawValue,
              normalizedValue: candidate.normalizedValue ?? null,
              confidence: candidate.confidence,
              boundingBoxJson: boundingBoxJson(candidate) ?? undefined,
              ruleName: candidate.ruleName,
            })),
          });
        }
      }
    } else {
      sheetCount = result.sheetNames?.length;

      // One SourcePage anchors this whole spreadsheet/CSV extraction pass;
      // ExtractionCandidate.rowNumber carries the per-row detail.
      const sourcePage =
        result.products.length > 0 || result.warnings.length > 0
          ? await prisma.sourcePage.create({
              data: {
                sourceDocumentId: sourceDocument.id,
                pageNumber: 1,
                extractionMethod: "SPREADSHEET",
                rawText: null,
                extractionConfidence: result.needsReview ? 0.5 : 0.9,
                needsReview: result.needsReview,
              },
            })
          : null;

      for (const product of result.products) {
        if (sourcePage && product.candidates.length > 0) {
          await prisma.extractionCandidate.createMany({
            data: product.candidates.map((candidate) => ({
              sourcePageId: sourcePage.id,
              rowNumber: product.sourceRowNumber ?? null,
              field: candidate.field,
              rawValue: candidate.rawValue,
              normalizedValue: candidate.normalizedValue ?? null,
              confidence: candidate.confidence,
              boundingBoxJson: boundingBoxJson(candidate) ?? undefined,
              ruleName: candidate.ruleName,
            })),
          });
        }

        const status =
          product.confidence >= NEEDS_REVIEW_CONFIDENCE_THRESHOLD && !result.needsReview
            ? "AUTO_ACCEPTED"
            : "NEEDS_REVIEW";

        await prisma.product.create({
          data: {
            sourceDocumentId: sourceDocument.id,
            sourcePageId: sourcePage?.id ?? null,
            sourceRowNumber: product.sourceRowNumber ?? null,
            sku: product.sku,
            description: product.description,
            priceCents: product.priceCents,
            include: product.include,
            status,
            confidence: product.confidence,
            extractionNotesJson:
              result.warnings.length > 0 ? { warnings: result.warnings } : undefined,
          },
        });
        productsCreated += 1;
      }
    }

    const completedAt = new Date();
    await prisma.processingJob.update({
      where: { id: job.id },
      data: {
        status: result.needsReview ? "NEEDS_REVIEW" : "COMPLETED",
        progressCurrent: 1,
        progressTotal: 1,
        completedAt,
      },
    });

    logger.info(
      {
        jobId: job.id,
        documentId: sourceDocument.id,
        sourceType: sourceDocument.extension,
        durationMs: completedAt.getTime() - startedAt.getTime(),
        pageCount,
        sheetCount,
        productsCreated,
        needsReview: result.needsReview,
      },
      "ingest-source-document: completed",
    );
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    logger.error(
      { jobId: job.id, documentId: sourceDocument.id, err: message },
      "ingest-source-document: failed",
    );
    await prisma.processingJob.update({
      where: { id: job.id },
      data: {
        status: "FAILED",
        errorCode: "INGESTION_FAILED",
        errorMessage: message,
        completedAt: new Date(),
      },
    });
    throw error;
  }
}
