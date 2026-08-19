import { readFileSync } from "node:fs";
import path from "node:path";
import { afterAll, beforeAll, describe, expect, it } from "vitest";
import type { FastifyInstance } from "fastify";
import { buildApp, loadConfig } from "@label-maker/api";
import { processIngestDocumentJob } from "@label-maker/worker";
import { buildMultipartUpload } from "./multipart.js";

const fixturesDir = path.join(import.meta.dirname, "..", "..", "fixtures", "csv");

describe("label-maker API integration", () => {
  let app: FastifyInstance;

  beforeAll(async () => {
    const config = loadConfig();
    app = await buildApp(config);
    await app.ready();

    // Clean slate: remove anything left over from a previous run, but keep
    // the seeded LabelTemplate presets.
    await app.prisma.labelRun.deleteMany({});
    await app.prisma.product.deleteMany({});
    await app.prisma.extractionCandidate.deleteMany({});
    await app.prisma.sourcePage.deleteMany({});
    await app.prisma.processingJob.deleteMany({});
    await app.prisma.sourceDocument.deleteMany({});
  });

  afterAll(async () => {
    await app.close();
  });

  it("GET /health returns ok", async () => {
    const response = await app.inject({ method: "GET", url: "/health" });
    expect(response.statusCode).toBe(200);
    expect(response.json()).toEqual({ status: "ok" });
  });

  it("GET /label-templates includes the seeded avery-5155 preset", async () => {
    const response = await app.inject({ method: "GET", url: "/label-templates" });
    expect(response.statusCode).toBe(200);
    const body = response.json() as { templates: Array<{ id: string }> };
    expect(body.templates.some((t) => t.id === "avery-5155")).toBe(true);
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

  it("full flow: upload CSV -> ingest -> review products -> generate label run", async () => {
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
    const { documentId, jobId } = uploadResponse.json() as { documentId: string; jobId: string };
    expect(documentId).toBeTruthy();
    expect(jobId).toBeTruthy();

    // No live worker process in this test run: invoke the same processor
    // the BullMQ worker would run, directly, so ingestion is synchronous
    // and deterministic here.
    await processIngestDocumentJob(app.prisma, {
      processingJobId: jobId,
      sourceDocumentId: documentId,
    });

    const jobResponse = await app.inject({ method: "GET", url: `/jobs/${jobId}` });
    expect(jobResponse.statusCode).toBe(200);
    expect(jobResponse.json().status).toBe("COMPLETED");

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

    const labelRunResponse = await app.inject({
      method: "POST",
      url: "/label-runs",
      payload: { sourceDocumentId: documentId, labelTemplateId: "avery-5155", copiesPerProduct: 8 },
    });
    expect(labelRunResponse.statusCode).toBe(201);
    const labelRunBody = labelRunResponse.json() as {
      labelRun: {
        status: string;
        placementPlanJson: { totalPlacements: number; totalSheets: number };
      };
      debugArtifact: { storageKey: string; byteSize: number };
    };
    expect(labelRunBody.labelRun.status).toBe("GENERATED");
    expect(labelRunBody.labelRun.placementPlanJson.totalPlacements).toBe(40); // 5 products x 8 copies
    expect(labelRunBody.labelRun.placementPlanJson.totalSheets).toBe(1);
    expect(labelRunBody.debugArtifact.storageKey).toMatch(/^artifacts\//);

    const artifactBuffer = await app.storage.read(labelRunBody.debugArtifact.storageKey);
    const artifact = JSON.parse(artifactBuffer.toString("utf-8")) as {
      labelTemplateId: string;
      placements: unknown[];
    };
    expect(artifact.labelTemplateId).toBe("avery-5155");
    expect(artifact.placements).toHaveLength(40);
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
});
