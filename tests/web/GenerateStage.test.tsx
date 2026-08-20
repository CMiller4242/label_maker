import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, fireEvent, render, screen } from "@testing-library/react";
import type { CreateLabelRunRequest, CreateLabelRunResponse } from "@label-maker/shared";

vi.mock("@label-maker/web/api/endpoints", () => ({
  createLabelRun: vi.fn(),
  downloadLabelRunArtifact: vi.fn(),
}));

const { createLabelRun, downloadLabelRunArtifact } = await import("@label-maker/web/api/endpoints");
const { ApiError } = await import("@label-maker/web/api/client");
const { GenerateStage } = await import("@label-maker/web/components/GenerateStage");

const createLabelRunMock = vi.mocked(createLabelRun);
const downloadArtifactMock = vi.mocked(downloadLabelRunArtifact);

const REQUEST: CreateLabelRunRequest = {
  sourceDocumentId: "11111111-1111-1111-1111-111111111111",
  labelTemplateId: "avery-5155",
  copiesPerProduct: 8,
  includeDebugArtifact: false,
};

function makeResponse(overrides: Partial<CreateLabelRunResponse["labelRun"]> = {}): CreateLabelRunResponse {
  return {
    labelRun: {
      id: "33333333-3333-3333-3333-333333333333",
      sourceDocumentId: REQUEST.sourceDocumentId,
      labelTemplateId: REQUEST.labelTemplateId,
      copiesPerProduct: REQUEST.copiesPerProduct,
      status: "GENERATED",
      placementPlanJson: {
        labelTemplateId: REQUEST.labelTemplateId,
        copiesPerProduct: REQUEST.copiesPerProduct,
        totalSheets: 1,
        totalPlacements: 24,
        placements: [],
      },
      generatedArtifactStorageKey: "runs/foo.docx",
      generatedArtifactSha256: "abc123",
      templateVersion: "1",
      templateSha256: "def456",
      sheetCount: 1,
      filledSlotCount: 24,
      emptySlotCount: 6,
      metadataJson: null,
      createdAt: new Date(),
      updatedAt: new Date(),
      ...overrides,
    },
    artifact: { storageKey: "runs/foo.docx", byteSize: 1024, sha256: "abc123", mimeType: "application/vnd.openxmlformats" },
  };
}

afterEach(() => {
  cleanup();
  vi.clearAllMocks();
});

describe("GenerateStage", () => {
  it("auto-triggers generation on mount and shows a pending state", () => {
    createLabelRunMock.mockReturnValue(new Promise(() => {})); // never resolves

    render(<GenerateStage request={REQUEST} onBack={vi.fn()} onStartOver={vi.fn()} />);

    expect(screen.getByRole("status")).toHaveTextContent(/generating/i);
    expect(createLabelRunMock).toHaveBeenCalledWith(REQUEST);
    expect(createLabelRunMock).toHaveBeenCalledTimes(1);
  });

  it("shows sheet/filled/blank counts, print guidance, and a download control on success", async () => {
    createLabelRunMock.mockResolvedValue(makeResponse());

    render(<GenerateStage request={REQUEST} onBack={vi.fn()} onStartOver={vi.fn()} />);

    expect(await screen.findByText("DOCX generated successfully.")).toBeInTheDocument();
    expect(screen.getByText("1")).toBeInTheDocument(); // sheets
    expect(screen.getByText("24")).toBeInTheDocument(); // filled
    expect(screen.getByText("6")).toBeInTheDocument(); // blank
    expect(screen.getByText(/actual size \/ 100%/i)).toBeInTheDocument();
    expect(screen.getByText(/does not confirm physical print alignment/i)).toBeInTheDocument();
    expect(screen.getByRole("button", { name: /download docx/i })).toBeInTheDocument();
  });

  it("shows an actionable error and lets the user retry generation", async () => {
    createLabelRunMock.mockRejectedValueOnce(
      new ApiError(500, { error: { code: "RENDER_FAILED", message: "The DOCX renderer crashed." } }),
    );

    render(<GenerateStage request={REQUEST} onBack={vi.fn()} onStartOver={vi.fn()} />);

    expect(await screen.findByRole("alert")).toHaveTextContent("The DOCX renderer crashed.");

    createLabelRunMock.mockResolvedValueOnce(makeResponse());
    fireEvent.click(screen.getByRole("button", { name: /retry/i }));

    expect(await screen.findByText("DOCX generated successfully.")).toBeInTheDocument();
    expect(createLabelRunMock).toHaveBeenCalledTimes(2);
  });

  it("adds a server-template hint when the failure code is TEMPLATE_UNAVAILABLE", async () => {
    createLabelRunMock.mockRejectedValue(
      new ApiError(409, { error: { code: "TEMPLATE_UNAVAILABLE", message: "Template could not be used." } }),
    );

    render(<GenerateStage request={REQUEST} onBack={vi.fn()} onStartOver={vi.fn()} />);

    expect(await screen.findByRole("alert")).toHaveTextContent("Template could not be used.");
    expect(screen.getByText(/server-side template issue/i)).toBeInTheDocument();
  });

  it("surfaces a download failure without discarding the successful generation", async () => {
    createLabelRunMock.mockResolvedValue(makeResponse());
    downloadArtifactMock.mockRejectedValue(new Error("download failed"));

    render(<GenerateStage request={REQUEST} onBack={vi.fn()} onStartOver={vi.fn()} />);

    fireEvent.click(await screen.findByRole("button", { name: /download docx/i }));

    expect(await screen.findByRole("alert")).toHaveTextContent("download failed");
    expect(screen.getByText("DOCX generated successfully.")).toBeInTheDocument();
  });
});
