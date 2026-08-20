import { afterEach, describe, expect, it, vi } from "vitest";
import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import type { ProcessingJob } from "@label-maker/shared";

vi.mock("@label-maker/web/api/endpoints", () => ({
  uploadDocument: vi.fn(),
  getJob: vi.fn(),
}));

const { uploadDocument, getJob } = await import("@label-maker/web/api/endpoints");
const { ApiError, NetworkError } = await import("@label-maker/web/api/client");
const { UploadStage } = await import("@label-maker/web/components/UploadStage");

const uploadDocumentMock = vi.mocked(uploadDocument);
const getJobMock = vi.mocked(getJob);

const DOCUMENT_ID = "11111111-1111-1111-1111-111111111111";
const JOB_ID = "22222222-2222-2222-2222-222222222222";

function makeJob(overrides: Partial<ProcessingJob> = {}): ProcessingJob {
  return {
    id: JOB_ID,
    sourceDocumentId: DOCUMENT_ID,
    jobType: "INGEST_DOCUMENT",
    status: "COMPLETED",
    progressCurrent: 1,
    progressTotal: 1,
    errorCode: null,
    errorMessage: null,
    createdAt: new Date(),
    startedAt: new Date(),
    completedAt: new Date(),
    ...overrides,
  };
}

function selectFile(name = "products.csv", type = "text/csv") {
  const input = screen.getByLabelText(/choose a file to upload/i) as HTMLInputElement;
  const file = new File(["sku,description,price"], name, { type });
  fireEvent.change(input, { target: { files: [file] } });
  return file;
}

afterEach(() => {
  cleanup();
  vi.clearAllMocks();
});

describe("UploadStage", () => {
  it("uploads a file, shows the filename/type, and lets the user continue once processing completes", async () => {
    uploadDocumentMock.mockResolvedValue({ documentId: DOCUMENT_ID, jobId: JOB_ID });
    getJobMock.mockResolvedValue(makeJob({ status: "COMPLETED" }));
    const onDocumentReady = vi.fn();

    render(<UploadStage onDocumentReady={onDocumentReady} />);
    selectFile("products.csv");

    expect(await screen.findByText("products.csv")).toBeInTheDocument();
    expect(screen.getByText("CSV")).toBeInTheDocument();

    const continueButton = await screen.findByRole("button", { name: /continue to product review/i });
    fireEvent.click(continueButton);
    expect(onDocumentReady).toHaveBeenCalledWith(DOCUMENT_ID);
  });

  it("surfaces a NEEDS_REVIEW completion as a note rather than an error", async () => {
    uploadDocumentMock.mockResolvedValue({ documentId: DOCUMENT_ID, jobId: JOB_ID });
    getJobMock.mockResolvedValue(makeJob({ status: "NEEDS_REVIEW" }));

    render(<UploadStage onDocumentReady={vi.fn()} />);
    selectFile();

    expect(await screen.findByText(/some rows need review/i)).toBeInTheDocument();
  });

  it("shows a server-reported error and lets the user retry the same file", async () => {
    uploadDocumentMock
      .mockRejectedValueOnce(
        new ApiError(422, { error: { code: "UNSUPPORTED_FILE_TYPE", message: "That file type isn't supported." } }),
      )
      .mockResolvedValueOnce({ documentId: DOCUMENT_ID, jobId: JOB_ID });
    getJobMock.mockResolvedValue(makeJob({ status: "COMPLETED" }));

    render(<UploadStage onDocumentReady={vi.fn()} />);
    selectFile("products.pdf", "application/pdf");

    expect(await screen.findByRole("alert")).toHaveTextContent("That file type isn't supported.");

    fireEvent.click(screen.getByRole("button", { name: /retry/i }));

    expect(await screen.findByRole("button", { name: /continue to product review/i })).toBeInTheDocument();
    expect(uploadDocumentMock).toHaveBeenCalledTimes(2);
  });

  it("shows a generic connection message when the request never reaches the server", async () => {
    uploadDocumentMock.mockRejectedValue(new NetworkError(new Error("offline")));

    render(<UploadStage onDocumentReady={vi.fn()} />);
    selectFile();

    expect(await screen.findByRole("alert")).toHaveTextContent(/check your connection/i);
  });

  it("reports a FAILED processing job with the job's own error message", async () => {
    uploadDocumentMock.mockResolvedValue({ documentId: DOCUMENT_ID, jobId: JOB_ID });
    getJobMock.mockResolvedValue(makeJob({ status: "FAILED", errorMessage: "Could not parse the PDF." }));

    render(<UploadStage onDocumentReady={vi.fn()} />);
    selectFile("products.pdf", "application/pdf");

    await waitFor(() => {
      expect(getJobMock).toHaveBeenCalled();
    });
    expect(await screen.findByRole("alert")).toHaveTextContent("Could not parse the PDF.");
  });
});
