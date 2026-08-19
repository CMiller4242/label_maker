import { mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import path from "node:path";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import type { FastifyInstance } from "fastify";
import JSZip from "jszip";
import { buildApp, loadConfig } from "@label-maker/api";
import { processIngestDocumentJob } from "@label-maker/worker";
import { buildMultipartUpload } from "./multipart.js";
import { buildSyntheticFixedGridDocx } from "./helpers/synthetic-fixed-grid-docx.js";

const fixturesDir = path.join(import.meta.dirname, "..", "..", "fixtures", "csv");
const templatesRoot = path.join(import.meta.dirname, "..", "..", "fixtures", "label-templates");

// A test-only LabelTemplate, backed by a small controlled synthetic
// fixed-grid .docx (see helpers/synthetic-fixed-grid-docx.ts) - NOT the
// real Avery 5155 template. It exists to exercise the real DOCX generation
// path end-to-end via the API, since the committed avery-5155 fixture is
// currently not a valid Word document (see
// packages/docx-renderer/src/README.md). Written under fixtures/label-templates
// (the API's configured templates root) only for the duration of this test
// file and removed in afterAll.
const SYNTHETIC_TEMPLATE_ID = "test-synthetic-fixed-grid";
const SYNTHETIC_TEMPLATE_DIR = path.join(templatesRoot, SYNTHETIC_TEMPLATE_ID);
const SYNTHETIC_TEMPLATE_STORAGE_KEY = `${SYNTHETIC_TEMPLATE_ID}/original.docx`;

describe("label-maker API integration", () => {
  let app: FastifyInstance;

  beforeAll(async () => {
    const config = loadConfig();
    app = await buildApp(config);
    await app.ready();

    // Clean slate: remove anything left over from a previous run, but keep
    // the seeded avery-5155/avery-22802 LabelTemplate presets.
    await app.prisma.labelRun.deleteMany({});
    await app.prisma.product.deleteMany({});
    await app.prisma.extractionCandidate.deleteMany({});
    await app.prisma.sourcePage.deleteMany({});
    await app.prisma.processingJob.deleteMany({});
    await app.prisma.sourceDocument.deleteMany({});
    await app.prisma.labelTemplate.deleteMany({ where: { id: SYNTHETIC_TEMPLATE_ID } });

    mkdirSync(SYNTHETIC_TEMPLATE_DIR, { recursive: true });
    const docxBuffer = await buildSyntheticFixedGridDocx({ columns: 4, rows: 15 });
    writeFileSync(path.join(SYNTHETIC_TEMPLATE_DIR, "original.docx"), docxBuffer);

    await app.prisma.labelTemplate.create({
      data: {
        id: SYNTHETIC_TEMPLATE_ID,
        displayName: "Test Synthetic Fixed Grid",
        renderingMode: "FIXED_GRID",
        columns: 4,
        rows: 15,
        labelsPerSheet: 60,
        templateStorageKey: SYNTHETIC_TEMPLATE_STORAGE_KEY,
        templateVersion: "0.0.0-test",
        sourceTemplateSha256: null,
        isPreset: false,
        configJson: {
          labelTextStyle: {
            fontFamily: "Calibri",
            skuFontSizeHalfPoints: 18,
            descriptionFontSizeHalfPoints: 16,
            priceFontSizeHalfPoints: 18,
            bold: false,
            colorHex: "000000",
            horizontalAlignment: "center",
            verticalAlignment: "center",
            lineSpacing: { lineRule: "auto", line: 240 },
            paragraphSpacingBeforeTwips: 0,
            paragraphSpacingAfterTwips: 0,
          },
        },
      },
    });
  });

  afterAll(async () => {
    await app.prisma.labelRun.deleteMany({ where: { labelTemplateId: SYNTHETIC_TEMPLATE_ID } });
    await app.prisma.labelTemplate.deleteMany({ where: { id: SYNTHETIC_TEMPLATE_ID } });
    rmSync(SYNTHETIC_TEMPLATE_DIR, { recursive: true, force: true });
    await app.close();
  });

  it("GET /health returns ok", async () => {
    const response = await app.inject({ method: "GET", url: "/health" });
    expect(response.statusCode).toBe(200);
    expect(response.json()).toEqual({ status: "ok" });
  });

  it("GET /label-templates includes the seeded presets", async () => {
    const response = await app.inject({ method: "GET", url: "/label-templates" });
    expect(response.statusCode).toBe(200);
    const body = response.json() as { templates: Array<{ id: string }> };
    expect(body.templates.some((t) => t.id === "avery-5155")).toBe(true);
    expect(body.templates.some((t) => t.id === "avery-22802")).toBe(true);
  });

  it("rejects unsupported file types with 415", async () => {
    const { body, contentTypeHeader } = buildMultipartUpload(
      "notes.txt",
      "text/plain",
      Buffer.from("just some text"),
    );
    const response = await app.inject({
      method: "POST",
      url: "/uploads",
      headers: { "content-type": contentTypeHeader },
      payload: body,
    });
    expect(response.statusCode).toBe(415);
    expect(response.json().error.code).toBe("UNSUPPORTED_MEDIA_TYPE");
  });

  it("returns 400 for an invalid label-run request body", async () => {
    const response = await app.inject({
      method: "POST",
      url: "/label-runs",
      payload: {
        sourceDocumentId: "not-a-uuid",
        labelTemplateId: "avery-5155",
        copiesPerProduct: 8,
      },
    });
    expect(response.statusCode).toBe(400);
  });

  describe("full flow", () => {
    let documentId: string;

    beforeAll(async () => {
      const csvBuffer = readFileSync(path.join(fixturesDir, "products-standard.csv"));
      const { body, contentTypeHeader } = buildMultipartUpload(
        "products-standard.csv",
        "text/csv",
        csvBuffer,
      );

      const uploadResponse = await app.inject({
        method: "POST",
        url: "/uploads",
        headers: { "content-type": contentTypeHeader },
        payload: body,
      });
      expect(uploadResponse.statusCode).toBe(201);
      const uploadBody = uploadResponse.json() as { documentId: string; jobId: string };
      documentId = uploadBody.documentId;

      // No live worker process in this test run: invoke the same processor
      // the BullMQ worker would run, directly, so ingestion is synchronous
      // and deterministic here.
      await processIngestDocumentJob(app.prisma, {
        processingJobId: uploadBody.jobId,
        sourceDocumentId: documentId,
      });
    });

    it("ingests the CSV into AUTO_ACCEPTED products", async () => {
      const productsResponse = await app.inject({
        method: "GET",
        url: `/documents/${documentId}/products`,
      });
      expect(productsResponse.statusCode).toBe(200);
      const productsBody = productsResponse.json() as {
        products: Array<{ sku: string; priceCents: number; status: string }>;
      };
      expect(productsBody.products).toHaveLength(5);
      expect(productsBody.products.every((p) => p.status === "AUTO_ACCEPTED")).toBe(true);
      expect(productsBody.products[0]?.sku).toBe("SKU-1001");
      expect(productsBody.products[0]?.priceCents).toBe(1449);
    });

    it("generates a real DOCX artifact for a validated FIXED_GRID template and downloads it", async () => {
      const labelRunResponse = await app.inject({
        method: "POST",
        url: "/label-runs",
        payload: {
          sourceDocumentId: documentId,
          labelTemplateId: SYNTHETIC_TEMPLATE_ID,
          copiesPerProduct: 8,
          includeDebugArtifact: true,
        },
      });
      expect(labelRunResponse.statusCode).toBe(201);
      const labelRunBody = labelRunResponse.json() as {
        labelRun: {
          id: string;
          status: string;
          sheetCount: number;
          filledSlotCount: number;
          emptySlotCount: number;
          generatedArtifactSha256: string;
          placementPlanJson: { totalPlacements: number; totalSheets: number };
        };
        artifact: { storageKey: string; byteSize: number; sha256: string; mimeType: string };
        debugArtifact?: { storageKey: string; byteSize: number; sha256: string };
      };

      expect(labelRunBody.labelRun.status).toBe("GENERATED");
      expect(labelRunBody.labelRun.placementPlanJson.totalPlacements).toBe(40); // 5 products x 8 copies
      expect(labelRunBody.labelRun.placementPlanJson.totalSheets).toBe(1);
      expect(labelRunBody.labelRun.sheetCount).toBe(1);
      expect(labelRunBody.labelRun.filledSlotCount).toBe(40);
      expect(labelRunBody.labelRun.emptySlotCount).toBe(20); // 60 - 40
      expect(labelRunBody.artifact.mimeType).toBe(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      );
      expect(labelRunBody.artifact.storageKey).toMatch(/^artifacts\/.+\.docx$/);
      expect(labelRunBody.debugArtifact?.storageKey).toMatch(/^artifacts\/.+\.json$/);

      // Re-open the generated artifact directly from storage and sanity check content.
      const docxBuffer = await app.storage.read(labelRunBody.artifact.storageKey);
      const zip = await JSZip.loadAsync(docxBuffer);
      const documentXml = await zip.file("word/document.xml")?.async("text");
      expect(documentXml).toContain("SKU-1001");
      expect(documentXml).toContain("$14.49");

      // GET /label-runs/:id/download returns the same bytes with correct headers.
      const downloadResponse = await app.inject({
        method: "GET",
        url: `/label-runs/${labelRunBody.labelRun.id}/download`,
      });
      expect(downloadResponse.statusCode).toBe(200);
      expect(downloadResponse.headers["content-type"]).toBe(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      );
      expect(downloadResponse.headers["content-disposition"]).toContain("attachment");
      expect(downloadResponse.headers["content-disposition"]).toMatch(
        new RegExp(`label-run-${labelRunBody.labelRun.id}\\.docx`),
      );
      expect(Buffer.from(downloadResponse.rawPayload)).toEqual(docxBuffer);
    });

    it("returns a typed TEMPLATE_UNAVAILABLE error for avery-5155 (committed fixture is not a valid Word document)", async () => {
      const response = await app.inject({
        method: "POST",
        url: "/label-runs",
        payload: {
          sourceDocumentId: documentId,
          labelTemplateId: "avery-5155",
          copiesPerProduct: 8,
        },
      });
      expect(response.statusCode).toBe(422);
      expect(response.json().error.code).toBe("TEMPLATE_UNAVAILABLE");
    });

    it("returns a typed UNSUPPORTED_RENDERING_MODE error for avery-22802 (FLOATING)", async () => {
      const response = await app.inject({
        method: "POST",
        url: "/label-runs",
        payload: {
          sourceDocumentId: documentId,
          labelTemplateId: "avery-22802",
          copiesPerProduct: 8,
        },
      });
      expect(response.statusCode).toBe(422);
      expect(response.json().error.code).toBe("UNSUPPORTED_RENDERING_MODE");
    });
  });
});
