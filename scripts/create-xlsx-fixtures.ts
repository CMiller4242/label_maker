import { fileURLToPath } from "node:url";
import path from "node:path";
import * as XLSX from "xlsx";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const outputPath = path.join(__dirname, "..", "fixtures", "xlsx", "products-two-sheets.xlsx");

const productsSheetRows: (string | number)[][] = [
  ["Item Number", "Description", "As Low As"],
  ["XLS-1001", "Stainless Steel Water Bottle", "$14.49"],
  ["XLS-1002", "Ceramic Coffee Mug 12oz", "$9.99"],
  ["XLS-1003", "Bamboo Cutting Board", "$24.00"],
];

// A second sheet with an equally viable-looking header row, on purpose:
// this is what forces parseWorkbook() to refuse to guess and return
// needsReview instead of silently picking one sheet.
const archivedSheetRows: (string | number)[][] = [
  ["Item Number", "Description", "As Low As"],
  ["ARC-9001", "Discontinued Travel Mug", "$7.49"],
  ["ARC-9002", "Old Style Tote Bag", "$3.99"],
];

function main(): void {
  const workbook = XLSX.utils.book_new();

  const productsSheet = XLSX.utils.aoa_to_sheet(productsSheetRows);
  XLSX.utils.book_append_sheet(workbook, productsSheet, "Products");

  const archivedSheet = XLSX.utils.aoa_to_sheet(archivedSheetRows);
  XLSX.utils.book_append_sheet(workbook, archivedSheet, "Archived");

  XLSX.writeFile(workbook, outputPath);
  console.log(`Wrote ${outputPath}`);
}

main();
