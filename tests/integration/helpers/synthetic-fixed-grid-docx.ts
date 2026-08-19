import JSZip from "jszip";

/**
 * Builds a minimal, structurally-valid .docx containing ONE normal in-flow,
 * fixed-layout table with `columns x rows` blank cells (no placeholder
 * text), plus a trailing <w:sectPr>. Controlled TEST FIXTURE ONLY - not the
 * real Avery 5155 template (see packages/docx-renderer/src/README.md for
 * why the committed one is currently unusable).
 */
export function buildSyntheticFixedGridDocxXml(options: { columns: number; rows: number }): string {
  const columnWidthTwips = 2610;
  const rowHeightTwips = 1440;

  const gridCols = Array.from(
    { length: options.columns },
    () => `<w:gridCol w:w="${columnWidthTwips}"/>`,
  ).join("");

  const cell = `<w:tc><w:tcPr><w:tcW w:w="${columnWidthTwips}" w:type="dxa"/><w:tcMar><w:top w:w="100" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar></w:tcPr><w:p/></w:tc>`;
  const row = `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="${rowHeightTwips}" w:hRule="exact"/></w:trPr>${cell.repeat(options.columns)}</w:tr>`;
  const rows = row.repeat(options.rows);

  const table =
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/>` +
    `<w:tblCellMar><w:top w:w="100" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>` +
    `<w:bottom w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar>` +
    `</w:tblPr><w:tblGrid>${gridCols}</w:tblGrid>${rows}</w:tbl>`;

  const sectPr =
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
    `<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="720" w:footer="720" w:gutter="0"/>` +
    `</w:sectPr>`;

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:body>${table}${sectPr}</w:body></w:document>`
  );
}

const CONTENT_TYPES_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
  `</Types>`;

const ROOT_RELS_XML =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
  `</Relationships>`;

export async function buildSyntheticFixedGridDocx(options: {
  columns: number;
  rows: number;
}): Promise<Buffer> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES_XML);
  zip.file("_rels/.rels", ROOT_RELS_XML);
  zip.file("word/document.xml", buildSyntheticFixedGridDocxXml(options));
  return zip.generateAsync({ type: "nodebuffer" });
}
