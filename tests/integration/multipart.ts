const BOUNDARY = "----labelMakerTestBoundary";

/** Builds a minimal single-file multipart/form-data body for fastify.inject(). */
export function buildMultipartUpload(
  filename: string,
  contentType: string,
  fileBuffer: Buffer,
): { body: Buffer; contentTypeHeader: string } {
  const preamble = Buffer.from(
    `--${BOUNDARY}\r\nContent-Disposition: form-data; name="file"; filename="${filename}"\r\nContent-Type: ${contentType}\r\n\r\n`,
    "utf-8",
  );
  const epilogue = Buffer.from(`\r\n--${BOUNDARY}--\r\n`, "utf-8");

  return {
    body: Buffer.concat([preamble, fileBuffer, epilogue]),
    contentTypeHeader: `multipart/form-data; boundary=${BOUNDARY}`,
  };
}
