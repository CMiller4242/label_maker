-- CreateEnum
CREATE TYPE "SourceDocumentExtension" AS ENUM ('PDF', 'CSV', 'XLSX', 'XLS');

-- CreateEnum
CREATE TYPE "ProcessingJobType" AS ENUM ('INGEST_DOCUMENT', 'GENERATE_LABEL_RUN');

-- CreateEnum
CREATE TYPE "ProcessingJobStatus" AS ENUM ('QUEUED', 'PROCESSING', 'COMPLETED', 'FAILED', 'NEEDS_REVIEW');

-- CreateEnum
CREATE TYPE "ExtractionMethod" AS ENUM ('NATIVE_TEXT', 'OCR', 'SPREADSHEET', 'MANUAL');

-- CreateEnum
CREATE TYPE "ExtractionField" AS ENUM ('SKU', 'DESCRIPTION', 'PRICE', 'INCLUDE_STATUS');

-- CreateEnum
CREATE TYPE "ProductStatus" AS ENUM ('AUTO_ACCEPTED', 'NEEDS_REVIEW', 'APPROVED', 'EXCLUDED');

-- CreateEnum
CREATE TYPE "LabelTemplateRenderingMode" AS ENUM ('FIXED_GRID', 'FLOATING');

-- CreateEnum
CREATE TYPE "LabelRunStatus" AS ENUM ('DRAFT', 'READY', 'GENERATED', 'FAILED');

-- CreateTable
CREATE TABLE "SourceDocument" (
    "id" TEXT NOT NULL,
    "originalFilename" TEXT NOT NULL,
    "mimeType" TEXT NOT NULL,
    "extension" "SourceDocumentExtension" NOT NULL,
    "storageKey" TEXT NOT NULL,
    "byteSize" INTEGER NOT NULL,
    "sha256" TEXT NOT NULL,
    "pageCount" INTEGER,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "SourceDocument_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "ProcessingJob" (
    "id" TEXT NOT NULL,
    "sourceDocumentId" TEXT NOT NULL,
    "jobType" "ProcessingJobType" NOT NULL,
    "status" "ProcessingJobStatus" NOT NULL DEFAULT 'QUEUED',
    "progressCurrent" INTEGER NOT NULL DEFAULT 0,
    "progressTotal" INTEGER NOT NULL DEFAULT 0,
    "errorCode" TEXT,
    "errorMessage" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "startedAt" TIMESTAMP(3),
    "completedAt" TIMESTAMP(3),

    CONSTRAINT "ProcessingJob_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "SourcePage" (
    "id" TEXT NOT NULL,
    "sourceDocumentId" TEXT NOT NULL,
    "pageNumber" INTEGER NOT NULL,
    "extractionMethod" "ExtractionMethod" NOT NULL,
    "rawText" TEXT,
    "extractionConfidence" DOUBLE PRECISION NOT NULL,
    "needsReview" BOOLEAN NOT NULL DEFAULT false,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "SourcePage_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "ExtractionCandidate" (
    "id" TEXT NOT NULL,
    "sourcePageId" TEXT,
    "rowNumber" INTEGER,
    "field" "ExtractionField" NOT NULL,
    "rawValue" TEXT NOT NULL,
    "normalizedValue" TEXT,
    "confidence" DOUBLE PRECISION NOT NULL,
    "boundingBoxJson" JSONB,
    "ruleName" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "ExtractionCandidate_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "Product" (
    "id" TEXT NOT NULL,
    "sourceDocumentId" TEXT NOT NULL,
    "sourcePageId" TEXT,
    "sourceRowNumber" INTEGER,
    "sku" TEXT,
    "description" TEXT,
    "priceCents" INTEGER,
    "include" BOOLEAN NOT NULL DEFAULT true,
    "status" "ProductStatus" NOT NULL DEFAULT 'NEEDS_REVIEW',
    "confidence" DOUBLE PRECISION NOT NULL,
    "extractionNotesJson" JSONB,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "Product_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "LabelTemplate" (
    "id" TEXT NOT NULL,
    "displayName" TEXT NOT NULL,
    "renderingMode" "LabelTemplateRenderingMode" NOT NULL,
    "columns" INTEGER NOT NULL,
    "rows" INTEGER NOT NULL,
    "labelsPerSheet" INTEGER NOT NULL,
    "templateStorageKey" TEXT,
    "configJson" JSONB NOT NULL,
    "isPreset" BOOLEAN NOT NULL DEFAULT false,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "LabelTemplate_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "LabelRun" (
    "id" TEXT NOT NULL,
    "sourceDocumentId" TEXT NOT NULL,
    "labelTemplateId" TEXT NOT NULL,
    "copiesPerProduct" INTEGER NOT NULL,
    "status" "LabelRunStatus" NOT NULL DEFAULT 'DRAFT',
    "placementPlanJson" JSONB NOT NULL,
    "generatedArtifactStorageKey" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "LabelRun_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE INDEX "SourceDocument_sha256_idx" ON "SourceDocument"("sha256");

-- CreateIndex
CREATE INDEX "ProcessingJob_status_idx" ON "ProcessingJob"("status");

-- CreateIndex
CREATE INDEX "ProcessingJob_sourceDocumentId_idx" ON "ProcessingJob"("sourceDocumentId");

-- CreateIndex
CREATE INDEX "SourcePage_sourceDocumentId_pageNumber_idx" ON "SourcePage"("sourceDocumentId", "pageNumber");

-- CreateIndex
CREATE INDEX "ExtractionCandidate_sourcePageId_idx" ON "ExtractionCandidate"("sourcePageId");

-- CreateIndex
CREATE INDEX "Product_sourceDocumentId_idx" ON "Product"("sourceDocumentId");

-- CreateIndex
CREATE INDEX "Product_status_idx" ON "Product"("status");

-- CreateIndex
CREATE INDEX "LabelRun_sourceDocumentId_idx" ON "LabelRun"("sourceDocumentId");

-- AddForeignKey
ALTER TABLE "ProcessingJob" ADD CONSTRAINT "ProcessingJob_sourceDocumentId_fkey" FOREIGN KEY ("sourceDocumentId") REFERENCES "SourceDocument"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "SourcePage" ADD CONSTRAINT "SourcePage_sourceDocumentId_fkey" FOREIGN KEY ("sourceDocumentId") REFERENCES "SourceDocument"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "ExtractionCandidate" ADD CONSTRAINT "ExtractionCandidate_sourcePageId_fkey" FOREIGN KEY ("sourcePageId") REFERENCES "SourcePage"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "Product" ADD CONSTRAINT "Product_sourceDocumentId_fkey" FOREIGN KEY ("sourceDocumentId") REFERENCES "SourceDocument"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "Product" ADD CONSTRAINT "Product_sourcePageId_fkey" FOREIGN KEY ("sourcePageId") REFERENCES "SourcePage"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "LabelRun" ADD CONSTRAINT "LabelRun_sourceDocumentId_fkey" FOREIGN KEY ("sourceDocumentId") REFERENCES "SourceDocument"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "LabelRun" ADD CONSTRAINT "LabelRun_labelTemplateId_fkey" FOREIGN KEY ("labelTemplateId") REFERENCES "LabelTemplate"("id") ON DELETE RESTRICT ON UPDATE CASCADE;
