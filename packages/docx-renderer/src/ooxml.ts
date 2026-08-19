import { createHash } from "node:crypto";
import JSZip from "jszip";
import { XMLBuilder, XMLParser, XMLValidator } from "fast-xml-parser";

/**
 * Low-level, label-agnostic helpers for reading and writing WordprocessingML
 * (.docx) packages directly via JSZip + XML parsing. Nothing in this file
 * knows about labels, Avery templates, or placement plans - see
 * inspect-template.ts / template-validation.ts / fixed-grid-renderer.ts for
 * that. This module intentionally never shells out to Word, LibreOffice, or
 * any conversion service: it manipulates the OOXML package bytes directly.
 */

export const WORD_DOCUMENT_XML_PATH = "word/document.xml";

export class InvalidDocxPackageError extends Error {
  readonly code = "INVALID_DOCX_PACKAGE";
}

/**
 * Thrown when a .docx package this module just generated fails structural
 * or XML well-formedness validation - i.e. the output would not actually
 * open in Microsoft Word, even though it is a structurally valid zip and
 * re-parses fine with lenient tools. A renderer must never return bytes
 * that fail this check.
 */
export class InvalidGeneratedDocxPackageError extends Error {
  readonly code = "INVALID_GENERATED_DOCX_PACKAGE";
  readonly issues: string[];

  constructor(issues: string[]) {
    super(`Generated .docx package failed validation:\n${issues.map((issue) => `  - ${issue}`).join("\n")}`);
    this.name = "InvalidGeneratedDocxPackageError";
    this.issues = issues;
  }
}

/** Attribute bag for one XML element, as produced by fast-xml-parser's `preserveOrder` mode. */
export type XmlAttributes = Record<string, string>;

/**
 * One XML node in fast-xml-parser's `preserveOrder` shape: an object with
 * exactly one "tag name" key (or "#text") holding the children array (or
 * text value), plus an optional ":@" key holding this element's attributes.
 * Using `unknown` for values (never `any`) and narrow accessor functions
 * below keeps this safe without fighting TypeScript's structural typing for
 * a shape that is inherently dynamic (the key IS the tag name).
 */
export type XmlNode = Record<string, unknown> & { ":@"?: XmlAttributes };

const ATTRIBUTE_KEY = ":@";
const TEXT_KEY = "#text";
const ATTRIBUTE_NAME_PREFIX = "@_";

function xmlParser(): XMLParser {
  return new XMLParser({
    preserveOrder: true,
    ignoreAttributes: false,
    attributeNamePrefix: ATTRIBUTE_NAME_PREFIX,
    trimValues: false,
    parseTagValue: false,
    parseAttributeValue: false,
  });
}

function xmlBuilder(): XMLBuilder {
  return new XMLBuilder({
    preserveOrder: true,
    ignoreAttributes: false,
    attributeNamePrefix: ATTRIBUTE_NAME_PREFIX,
    suppressBooleanAttributes: false,
    unpairedTags: [],
  });
}

/** Parses an XML document into an array of root-level nodes (`preserveOrder` shape). */
export function parseXml(xml: string): XmlNode[] {
  return xmlParser().parse(xml) as XmlNode[];
}

/**
 * Serializes a `preserveOrder` node array back to an XML string, with a
 * standard XML declaration - always exactly one. `parseXml()` (via
 * fast-xml-parser's `preserveOrder: true`) captures a source document's own
 * `<?xml ...?>` prolog as a literal node at the front of its returned
 * array, so building a tree derived from a parsed document (e.g.
 * `templatePackage.documentTree`) without stripping that node first would
 * re-emit it verbatim *in addition to* the declaration prepended below -
 * two declarations, which is invalid XML. `unzip`/a lenient re-parse never
 * catch this, but Microsoft Word's strict OOXML parser rejects the whole
 * file outright ("Word experienced an error trying to open the file").
 */
export function buildXml(nodes: XmlNode[]): string {
  const withoutDeclaration = nodes.filter((node) => tagNameOf(node) !== "?xml");
  const body = xmlBuilder().build(withoutDeclaration) as string;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n${body}`;
}

/**
 * Validates that a just-generated .docx package is actually openable by
 * Microsoft Word - not just a structurally valid zip. Checks:
 *
 * 1. Every part present in `expectedEntryNames` (captured from the source
 *    template at load time) is present in `zip`, and no unexpected part
 *    was added - a renderer must only ever rewrite existing parts.
 * 2. Every `.xml`/`.rels` part is well-formed XML per fast-xml-parser's
 *    `XMLValidator` - a *strict* check (catches things like more than one
 *    XML declaration, unclosed tags, mismatched namespaces) distinct from
 *    just re-parsing with the same lenient parser used to build the tree,
 *    or checking zip/CRC integrity with a tool like `unzip -t` - neither
 *    of which would have caught the duplicate-declaration bug this
 *    function exists to prevent from recurring.
 *
 * Throws InvalidGeneratedDocxPackageError (with every issue found, not
 * just the first) rather than returning a boolean, so a renderer can never
 * silently return bytes that fail this check.
 */
export async function assertGeneratedDocxPackageIsValid(
  zip: JSZip,
  expectedEntryNames: ReadonlySet<string>,
): Promise<void> {
  const issues: string[] = [];

  const actualEntryNames = new Set(Object.keys(zip.files));
  for (const name of expectedEntryNames) {
    if (!actualEntryNames.has(name)) issues.push(`Missing expected package part: "${name}".`);
  }
  for (const name of actualEntryNames) {
    if (!expectedEntryNames.has(name)) {
      issues.push(`Unexpected package part not present in the source template: "${name}".`);
    }
  }

  const xmlEntries = Object.entries(zip.files).filter(
    ([name, entry]) => !entry.dir && (name.endsWith(".xml") || name.endsWith(".rels")),
  );
  for (const [name, entry] of xmlEntries) {
    const text = await entry.async("text");
    const result = XMLValidator.validate(text);
    if (result !== true) {
      issues.push(
        `"${name}" is not well-formed XML: ${result.err.msg} (line ${result.err.line}, col ${result.err.col}).`,
      );
    }
  }

  if (issues.length > 0) {
    throw new InvalidGeneratedDocxPackageError(issues);
  }
}

/** The tag name of a node (e.g. "w:tbl"), or undefined if this is a text node. */
export function tagNameOf(node: XmlNode): string | undefined {
  return Object.keys(node).find((key) => key !== ATTRIBUTE_KEY);
}

export function isTextNode(node: XmlNode): boolean {
  return TEXT_KEY in node;
}

export function textOf(node: XmlNode): string {
  const value = node[TEXT_KEY];
  return typeof value === "string" ? value : "";
}

/** This element's own attributes (empty object if none). */
export function attributesOf(node: XmlNode): XmlAttributes {
  return node[ATTRIBUTE_KEY] ?? {};
}

export function getAttr(node: XmlNode, name: string): string | undefined {
  return attributesOf(node)[`${ATTRIBUTE_NAME_PREFIX}${name}`];
}

/** Sets (or overwrites) a single attribute on an element node, creating the attribute bag if absent. */
export function setAttr(node: XmlNode, name: string, value: string): void {
  const existing = node[ATTRIBUTE_KEY];
  const attrs: XmlAttributes = existing ?? {};
  attrs[`${ATTRIBUTE_NAME_PREFIX}${name}`] = value;
  node[ATTRIBUTE_KEY] = attrs;
}

/** This element's children (the array under its tag key), or [] for a text node / empty element. */
export function childrenOf(node: XmlNode): XmlNode[] {
  const tag = tagNameOf(node);
  if (!tag) return [];
  const value = node[tag];
  return Array.isArray(value) ? (value as XmlNode[]) : [];
}

/** Replaces an element's children array in place. */
export function setChildren(node: XmlNode, children: XmlNode[]): void {
  const tag = tagNameOf(node);
  if (!tag) throw new InvalidDocxPackageError("Cannot set children on a text node.");
  node[tag] = children;
}

/** Finds all descendant elements with the given tag name, depth-first, at any depth. */
export function findAllByTag(nodes: XmlNode[], tag: string): XmlNode[] {
  const results: XmlNode[] = [];
  for (const node of nodes) {
    if (tagNameOf(node) === tag) results.push(node);
    results.push(...findAllByTag(childrenOf(node), tag));
  }
  return results;
}

/** Finds the first descendant element with the given tag name, depth-first, at any depth. */
export function findFirstByTag(nodes: XmlNode[], tag: string): XmlNode | undefined {
  for (const node of nodes) {
    const nodeTag = tagNameOf(node);
    if (nodeTag === tag) return node;
    const found = findFirstByTag(childrenOf(node), tag);
    if (found) return found;
  }
  return undefined;
}

/** Finds only the direct children (not descendants) of `nodes` matching `tag`. */
export function findDirectChildren(nodes: XmlNode[], tag: string): XmlNode[] {
  return nodes.filter((node) => tagNameOf(node) === tag);
}

/** Concatenates all `w:t` text descendants of a node, in document order. */
export function extractText(node: XmlNode): string {
  return findAllByTag(childrenOf(node), "w:t")
    .map((t) => textOf(childrenOf(t)[0] ?? { "#text": "" }))
    .join("");
}

/** Deep-clones a node (or node array) so mutations never affect the source template tree. */
export function cloneNode<T>(value: T): T {
  return structuredClone(value);
}

/**
 * Converts a plain `{ "w:val": "x" }` attribute bag into the
 * `{ "@_w:val": "x" }` shape fast-xml-parser's `preserveOrder` builder
 * requires (see ATTRIBUTE_NAME_PREFIX) - every make*Element() helper below
 * must apply this before storing attributes under ":@", or attributes
 * silently lose their namespace prefix (or are dropped) on serialization.
 */
function prefixAttrs(attrs: XmlAttributes): XmlAttributes {
  const prefixed: XmlAttributes = {};
  for (const [name, value] of Object.entries(attrs)) {
    prefixed[`${ATTRIBUTE_NAME_PREFIX}${name}`] = value;
  }
  return prefixed;
}

/** Builds a text-node-wrapped element, e.g. makeElement("w:t", {}, "Hello") -> <w:t>Hello</w:t>. */
export function makeTextElement(tag: string, attrs: XmlAttributes, text: string): XmlNode {
  const node: XmlNode = { [tag]: [{ [TEXT_KEY]: text }] };
  if (Object.keys(attrs).length > 0) node[ATTRIBUTE_KEY] = prefixAttrs(attrs);
  return node;
}

/** Builds an element with no text content, only attributes (e.g. a self-closing tag like <w:sz w:val="20"/>). */
export function makeEmptyElement(tag: string, attrs: XmlAttributes = {}): XmlNode {
  const node: XmlNode = { [tag]: [] };
  if (Object.keys(attrs).length > 0) node[ATTRIBUTE_KEY] = prefixAttrs(attrs);
  return node;
}

/** Builds an element wrapping other elements, e.g. makeElement("w:rPr", {}, [sz, color]). */
export function makeElement(tag: string, attrs: XmlAttributes, children: XmlNode[]): XmlNode {
  const node: XmlNode = { [tag]: children };
  if (Object.keys(attrs).length > 0) node[ATTRIBUTE_KEY] = prefixAttrs(attrs);
  return node;
}

// --- Unit conversions -------------------------------------------------
// OOXML measures most WordprocessingML lengths in twentieths of a point
// ("twips"): 1 inch = 1440 twips = 72 points. Font sizes (w:sz) are in
// half-points instead.

export function twipsToInches(twips: number): number {
  return twips / 1440;
}

export function twipsToPoints(twips: number): number {
  return twips / 20;
}

export function pointsToTwips(points: number): number {
  return Math.round(points * 20);
}

export function halfPointsToPoints(halfPoints: number): number {
  return halfPoints / 2;
}

// --- Package loading ----------------------------------------------------

/** An opened .docx/.dotx template package: the zip plus the parsed word/document.xml tree. */
export interface TemplatePackage {
  /** A human-readable label for error messages (a storage key, file path, or fixture name). */
  templateStorageKey: string;
  zip: JSZip;
  /** Raw document.xml text, preserved for reference/debugging. */
  documentXml: string;
  /** Parsed root nodes of document.xml (should contain exactly one `w:document` element). */
  documentTree: XmlNode[];
}

/**
 * Opens a .docx package from bytes and parses word/document.xml. Throws
 * InvalidDocxPackageError if the zip does not contain a word/document.xml
 * part - this is the one required part of any real Word document, and its
 * absence (e.g. because the file is actually some other Office package
 * type, or was corrupted/truncated) is caught here rather than surfacing as
 * a confusing downstream XML parse failure.
 */
export async function loadTemplateDocx(
  buffer: Buffer,
  templateStorageKey: string,
): Promise<TemplatePackage> {
  const zip = await JSZip.loadAsync(buffer);
  const documentXmlFile = zip.file(WORD_DOCUMENT_XML_PATH);
  if (!documentXmlFile) {
    const partNames = Object.keys(zip.files).slice(0, 20);
    throw new InvalidDocxPackageError(
      `"${templateStorageKey}" is not a usable Word document: the package has no ` +
        `${WORD_DOCUMENT_XML_PATH} part. This zip may be a different Office package ` +
        `type (e.g. a theme/.thmx file), corrupted, or truncated. Parts found: ` +
        `${partNames.length > 0 ? partNames.join(", ") : "(none)"}.`,
    );
  }
  const documentXml = await documentXmlFile.async("text");
  const documentTree = parseXml(documentXml);
  return { templateStorageKey, zip, documentXml, documentTree };
}

export function sha256Hex(buffer: Buffer): string {
  return createHash("sha256").update(buffer).digest("hex");
}
