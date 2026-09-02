import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = "D:/Woodpecker工具/Result/Woodpecker_Workstation4_20260327T111400Z_20260327T111800Z.xlsx";
const previewDir = "D:/Woodpecker工具/.codex-temp/01a06089-ec10-7173-84c1-9c50f5d5c6/previews-before";

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const overview = await workbook.inspect({
  kind: "workbook,sheet,table",
  maxChars: 12000,
  tableMaxRows: 10,
  tableMaxCols: 20,
  tableMaxCellChars: 240,
});
console.log("OVERVIEW\n" + overview.ndjson);

await fs.mkdir(previewDir, { recursive: true });
const sheetInfo = await workbook.inspect({ kind: "sheet", include: "id,name", maxChars: 8000 });
console.log("SHEETS\n" + sheetInfo.ndjson);

for (let i = 0; i < workbook.worksheets.items.length; i += 1) {
  const sheet = workbook.worksheets.getItemAt(i);
  console.log("TABLES\n" + JSON.stringify(sheet.tables.items.map((table) => ({
    name: table.name,
    style: table.style,
    showHeaders: table.showHeaders,
    showTotals: table.showTotals,
    showBandedColumns: table.showBandedColumns,
    showFilterButton: table.showFilterButton,
  }))));
  const used = sheet.getUsedRange();
  console.log(`SHEET ${i + 1}: ${sheet.name}`);
  if (!used) {
    console.log("EMPTY");
    continue;
  }
  console.log("USED VALUES\n" + JSON.stringify(used.values));
  console.log("USED FORMULAS\n" + JSON.stringify(used.formulas));
  const style = await workbook.inspect({
    kind: "computedStyle",
    sheetId: sheet.name,
    range: used.address,
    maxChars: 6000,
  });
  console.log("STYLES\n" + style.ndjson);
  const preview = await workbook.render({ sheetName: sheet.name, autoCrop: "all", scale: 1.5, format: "png" });
  const safeName = sheet.name.replace(/[\\/:*?"<>|]/g, "_");
  await fs.writeFile(`${previewDir}/${i + 1}-${safeName}.png`, new Uint8Array(await preview.arrayBuffer()));
}

const deleteHelp = workbook.help("range.delete", { include: "index,examples,notes", maxChars: 5000 });
console.log("DELETE HELP\n" + deleteHelp.ndjson);
