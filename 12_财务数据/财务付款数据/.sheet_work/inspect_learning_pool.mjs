import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = process.argv[2];
if (!inputPath) throw new Error("Expected input .xlsx path");

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const summary = await workbook.inspect({
  kind: "workbook,sheet,table,computedStyle",
  maxChars: 6000,
  tableMaxRows: 5,
  tableMaxCols: 18,
  tableMaxCellChars: 80,
});
process.stdout.write(summary.ndjson + "\n");

const preview = await workbook.render({
  sheetName: "学习池",
  range: "A1:Q14",
  scale: 1.5,
  format: "png",
});
const previewPath = process.argv[3];
if (!previewPath) throw new Error("Expected preview PNG path");
await (await import("node:fs/promises")).writeFile(
  previewPath,
  new Uint8Array(await preview.arrayBuffer()),
);
