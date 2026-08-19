-- AlterTable
ALTER TABLE "LabelRun" ADD COLUMN     "emptySlotCount" INTEGER,
ADD COLUMN     "filledSlotCount" INTEGER,
ADD COLUMN     "generatedArtifactSha256" TEXT,
ADD COLUMN     "metadataJson" JSONB,
ADD COLUMN     "sheetCount" INTEGER,
ADD COLUMN     "templateSha256" TEXT,
ADD COLUMN     "templateVersion" TEXT;

-- AlterTable
ALTER TABLE "LabelTemplate" ADD COLUMN     "sourceTemplateSha256" TEXT,
ADD COLUMN     "templateVersion" TEXT;
