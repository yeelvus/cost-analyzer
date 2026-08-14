import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const [officialPath, adjustedCopyPath, outputPath] = process.argv.slice(2);
if (!officialPath || !adjustedCopyPath || !outputPath) {
  throw new Error("Expected official workbook, adjusted copy, and output paths");
}

const normalize = (value) => (value === null || value === undefined ? "" : String(value).trim());
const money = (value) => {
  const n = Number(value);
  return Number.isFinite(n) ? n.toFixed(2) : "0.00";
};
const identity = (row) => [row[1], row[2], row[6], row[7], money(row[8])]
  .map(normalize)
  .join("\u001F");

const official = await SpreadsheetFile.importXlsx(await FileBlob.load(officialPath));
const adjusted = await SpreadsheetFile.importXlsx(await FileBlob.load(adjustedCopyPath));
const officialSheet = official.worksheets.getItem("学习池");
const adjustedSheet = adjusted.worksheets.getItem("学习池");

const officialRows = officialSheet.getRange("A1:M19514").values;
const adjustedRows = adjustedSheet.getRange("A1:Q19514").values;
if (officialRows.length !== adjustedRows.length) {
  throw new Error(`Row count differs: official=${officialRows.length}, adjusted=${adjustedRows.length}`);
}

const updatedCategories = [];
const changed = [];
for (let i = 1; i < officialRows.length; i += 1) {
  const officialRow = officialRows[i];
  const adjustedRow = adjustedRows[i];
  if (identity(officialRow) !== identity(adjustedRow)) {
    throw new Error(`Immutable fields differ at Excel row ${i + 1}`);
  }
  const modified = normalize(adjustedRow[16]) === "已修改";
  if (modified) {
    const previous = [adjustedRow[13], adjustedRow[14], adjustedRow[15]].map(normalize);
    const current = [officialRow[3], officialRow[4], officialRow[5]].map(normalize);
    if (previous.join("\u001F") !== current.join("\u001F")) {
      throw new Error(`Baseline categories differ at Excel row ${i + 1}`);
    }
    updatedCategories.push([adjustedRow[3], adjustedRow[4], adjustedRow[5]]);
    changed.push(i + 1);
  } else {
    updatedCategories.push([officialRow[3], officialRow[4], officialRow[5]]);
  }
}

officialSheet.getRange("D2:F19514").values = updatedCategories;
officialSheet.freezePanes.freezeRows(1);
await fs.mkdir(new URL(".", `file://${outputPath}`).pathname, { recursive: true }).catch(() => {});
const output = await SpreadsheetFile.exportXlsx(official);
await output.save(outputPath);

console.log(JSON.stringify({
  rows: officialRows.length - 1,
  updatedRows: changed.length,
  firstUpdatedExcelRows: changed.slice(0, 10),
  lastUpdatedExcelRows: changed.slice(-10),
  outputPath,
}, null, 2));
