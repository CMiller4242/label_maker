#!/usr/bin/env -S pnpm exec tsx
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { inspectDocxTemplate } from "@label-maker/docx-renderer";

/**
 * Reusable DOCX template inspection CLI. Opens a .docx package directly via
 * JSZip + XML parsing (no Microsoft Word or any installed office suite
 * required) and prints a structured JSON report: page size, section
 * margins, every table's layout/grid/row/cell/font/paragraph properties,
 * total writable cell count, and a best-effort classification against the
 * two known preset shapes (Avery 5155 fixed-grid, Avery 22802 floating).
 *
 * Usage:
 *   pnpm exec tsx scripts/inspect-docx-template.ts <file.docx> [<file2.docx> ...]
 *   pnpm exec tsx scripts/inspect-docx-template.ts   # inspects both committed fixtures
 */

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.join(__dirname, "..");

const DEFAULT_TARGETS = [
  path.join(repoRoot, "fixtures", "label-templates", "avery-5155", "original.docx"),
  path.join(repoRoot, "fixtures", "label-templates", "avery-22802", "original.docx"),
];

async function inspectOne(filePath: string): Promise<void> {
  const buffer = readFileSync(filePath);
  const relativePath = path.relative(repoRoot, filePath);
  const report = await inspectDocxTemplate({ buffer, filePath: relativePath });

  console.log(`\n${"=".repeat(80)}`);
  console.log(`Inspecting: ${relativePath}`);
  console.log("=".repeat(80));
  console.log(JSON.stringify(report, null, 2));
}

async function main(): Promise<void> {
  const args = process.argv.slice(2);
  const targets = args.length > 0 ? args.map((p) => path.resolve(p)) : DEFAULT_TARGETS;

  for (const target of targets) {
    await inspectOne(target);
  }
}

main().catch((error: unknown) => {
  console.error("Inspection failed:", error);
  process.exitCode = 1;
});
