import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "D:/Woodpecker工具/Result/Woodpecker_Workstation4_20260327T111400Z_20260327T111800Z.xlsx";
const outputDir = "D:/Woodpecker工具/outputs/01a06089-ec10-7173-84c1-9c50f5d5c6";
const outputPath = `${outputDir}/Woodpecker_Workstation4_20260327T111400Z_20260327T111800Z.xlsx`;
const previewPath = `${outputDir}/Behavior Logs.png`;

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const sheet = workbook.worksheets.getItem("Behavior Logs");
const used = sheet.getUsedRange();
const sourceValues = used.values;

if (sourceValues.length !== 18 || sourceValues[0].length !== 6) {
  throw new Error(`Unexpected workbook shape: ${sourceValues.length}x${sourceValues[0]?.length ?? 0}`);
}

const expectedHeaders = [
  "winlog.event_data.CommandLine",
  "@timestamp",
  "detection.information.name",
  "detection.information.tags",
  "detection.information.detail",
  "winlog_event_id",
];
if (JSON.stringify(sourceValues[0]) !== JSON.stringify(expectedHeaders)) {
  throw new Error("Workbook headers do not match the expected Woodpecker schema.");
}

// Retain direct or distinct corroborating evidence for the four YAML abilities.
// Excel rows kept: 2-7 and 9-11. Row 8 is a generic duplicate of row 7;
// rows 12-18 are cleanup/reversal/summary records not defined in the supplied payload.
const keepExcelRows = [2, 3, 4, 5, 6, 7, 9, 10, 11];
const filteredValues = [sourceValues[0], ...keepExcelRows.map((row) => sourceValues[row - 1])];

const sourceTable = sheet.tables.items[0];
const tableProperties = {
  name: sourceTable.name,
  style: sourceTable.style,
  showHeaders: sourceTable.showHeaders,
  showTotals: sourceTable.showTotals,
  showBandedColumns: sourceTable.showBandedColumns,
  showFilterButton: sourceTable.showFilterButton,
};

sourceTable.delete();
sheet.getRange("A1:F18").clear({ applyTo: "all" });
sheet.getRange("A1:F10").values = filteredValues;

const table = sheet.tables.add("A1:F10", true, tableProperties.name);
table.style = tableProperties.style;
table.showHeaders = tableProperties.showHeaders;
table.showTotals = tableProperties.showTotals;
table.showBandedColumns = tableProperties.showBandedColumns;
table.showFilterButton = tableProperties.showFilterButton;

sheet.getRange("A2:A10").format.wrapText = true;
sheet.getRange("E2:E10").format.wrapText = true;
sheet.getRange("F2:F10").format.wrapText = true;
sheet.getRange("B2:B10").format.numberFormat = 'm"月" d, yyyy hh:mm:ss';
sheet.getRange("A1:F10").format.autofitRows();

await fs.mkdir(outputDir, { recursive: true });
const rendered = await workbook.render({ sheetName: "Behavior Logs", autoCrop: "all", scale: 1.5, format: "png" });
await fs.writeFile(previewPath, new Uint8Array(await rendered.arrayBuffer()));

const exported = await SpreadsheetFile.exportXlsx(workbook);
await exported.save(outputPath);

const checkWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(outputPath));
const checkSheet = checkWorkbook.worksheets.getItem("Behavior Logs");
const checkValues = checkSheet.getUsedRange(true).values;
if (checkValues.length !== 10 || checkValues[0].length !== 6) {
  throw new Error(`Verification failed: output shape is ${checkValues.length}x${checkValues[0]?.length ?? 0}`);
}

const removedNames = checkValues.slice(1).map((row) => row[2]);
if (removedNames.includes("Suspicious Device") || checkValues.some((row) => String(row[0] ?? "").includes("Remove-Item"))) {
  throw new Error("Verification failed: an excluded cleanup or summary record remains.");
}

const requiredEvidence = ["RTCore64", "EDRSandBlast.exe", "mimikatz.exe"];
const flattened = JSON.stringify(checkValues);
for (const token of requiredEvidence) {
  if (!flattened.includes(token)) throw new Error(`Verification failed: required evidence ${token} is missing.`);
}

const errors = await checkWorkbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
  maxChars: 4000,
});

const inspectSidecar = `${outputPath}.inspect.ndjson`;
await fs.rm(inspectSidecar, { force: true });

console.log(JSON.stringify({
  outputPath,
  retainedDataRows: checkValues.length - 1,
  removedDataRows: sourceValues.length - checkValues.length,
  retainedDetections: checkValues.slice(1).map((row) => row[2]),
  usedRange: "A1:F10",
  formulaErrors: errors.ndjson,
}, null, 2));
