import { describe, expect, it } from "vitest";
import { XMLValidator } from "fast-xml-parser";
import JSZip from "jszip";
import {
  InvalidGeneratedDocxPackageError,
  assertGeneratedDocxPackageIsValid,
  buildXml,
  parseXml,
} from "@label-maker/docx-renderer";

/**
 * Regression coverage for the exact bug that made real Avery 5155 sample
 * DOCX artifacts unopenable in Microsoft Word ("Word experienced an error
 * trying to open the file"): buildXml() always prepended its own
 * `<?xml ...?>` declaration WITHOUT stripping the declaration node
 * parseXml() captures from a source document that already had one (via
 * fast-xml-parser's `preserveOrder: true`) - producing two declarations in
 * word/document.xml. That is invalid XML per spec, but neither `unzip -t`
 * (zip/CRC integrity only) nor re-parsing the output with the same
 * lenient parser used to build it ever caught it - only a strict
 * validator (or Word itself) does.
 */
describe("buildXml", () => {
  it("produces exactly one XML declaration when rebuilding a tree parsed from a document that already had one", () => {
    const source =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<root><child>text</child></root>';
    const tree = parseXml(source);
    const rebuilt = buildXml(tree);

    expect((rebuilt.match(/<\?xml/g) ?? []).length).toBe(1);
    expect(XMLValidator.validate(rebuilt)).toBe(true);
  });

  it("still produces exactly one declaration when rebuilding a tree that never had one", () => {
    const tree = parseXml("<root><child>text</child></root>");
    const rebuilt = buildXml(tree);

    expect((rebuilt.match(/<\?xml/g) ?? []).length).toBe(1);
    expect(XMLValidator.validate(rebuilt)).toBe(true);
  });

  it("round-trips real Word document content (namespaced elements/attributes) as well-formed XML", () => {
    const source =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:body><w:p><w:r><w:t xml:space="preserve">Hello</w:t></w:r></w:p></w:body>' +
      "</w:document>";
    const rebuilt = buildXml(parseXml(source));

    expect((rebuilt.match(/<\?xml/g) ?? []).length).toBe(1);
    const result = XMLValidator.validate(rebuilt);
    expect(result).toBe(true);
  });
});

describe("assertGeneratedDocxPackageIsValid", () => {
  function zipWith(entries: Record<string, string>): JSZip {
    const zip = new JSZip();
    // createFolders: false - same as fixed-grid-renderer.ts's real usage -
    // so JSZip doesn't auto-vivify a spurious parent folder entry (e.g.
    // "word/") that these tests don't expect and never asked for.
    for (const [name, content] of Object.entries(entries)) {
      zip.file(name, content, { createFolders: false });
    }
    return zip;
  }

  it("resolves when every expected part is present and every XML/rels part is well-formed", async () => {
    const zip = zipWith({
      "[Content_Types].xml": '<?xml version="1.0"?><Types/>',
      "word/document.xml": '<?xml version="1.0"?><w:document/>',
    });

    await expect(
      assertGeneratedDocxPackageIsValid(zip, new Set(["[Content_Types].xml", "word/document.xml"])),
    ).resolves.toBeUndefined();
  });

  it("throws InvalidGeneratedDocxPackageError when an XML part has more than one declaration (the exact regression)", async () => {
    const zip = zipWith({
      "word/document.xml": '<?xml version="1.0"?>\r\n<?xml version="1.0"?><w:document/>',
    });

    await expect(assertGeneratedDocxPackageIsValid(zip, new Set(["word/document.xml"]))).rejects.toThrow(
      InvalidGeneratedDocxPackageError,
    );
    try {
      await assertGeneratedDocxPackageIsValid(zip, new Set(["word/document.xml"]));
      expect.unreachable();
    } catch (error) {
      const issues = (error as InvalidGeneratedDocxPackageError).issues;
      expect(
        issues.some((i) => i.includes("word/document.xml") && i.includes("not well-formed")),
      ).toBe(true);
    }
  });

  it("throws when an expected package part is missing", async () => {
    const zip = zipWith({ "word/document.xml": "<root/>" });

    try {
      await assertGeneratedDocxPackageIsValid(zip, new Set(["word/document.xml", "word/styles.xml"]));
      expect.unreachable();
    } catch (error) {
      const issues = (error as InvalidGeneratedDocxPackageError).issues;
      expect(issues.some((i) => i.includes('Missing expected package part: "word/styles.xml"'))).toBe(
        true,
      );
    }
  });

  it("throws when an unexpected package part was added", async () => {
    const zip = zipWith({ "word/document.xml": "<root/>", "word/extra.xml": "<x/>" });

    try {
      await assertGeneratedDocxPackageIsValid(zip, new Set(["word/document.xml"]));
      expect.unreachable();
    } catch (error) {
      const issues = (error as InvalidGeneratedDocxPackageError).issues;
      expect(issues.some((i) => i.includes('"word/extra.xml"'))).toBe(true);
    }
  });
});
