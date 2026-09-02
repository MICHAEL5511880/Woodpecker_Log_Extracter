import path from "node:path";
import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

export const OUTPUT_COLUMNS = [
  "winlog.event_data.CommandLine",
  "@timestamp",
  "detection.information.name",
  "detection.information.tags",
  "detection.information.detail",
  "winlog_event_id",
];

const EXCEL_MAX_DATA_ROWS = 1_048_575;
const WRITE_BATCH_SIZE = 5_000;

function excelCell(value, separator = " | ") {
  if (value === null || value === undefined) return "";
  if (Array.isArray(value)) return value.map((item) => excelCell(item, separator)).filter(Boolean).join(separator);
  if (typeof value === "object") return JSON.stringify(value);
  return String(value);
}

function offsetMinutes(timezoneOffset) {
  const match = /^([+-])(\d{2}):(\d{2})$/.exec(timezoneOffset);
  if (!match) throw new Error(`timezoneOffset 格式錯誤：${timezoneOffset}`);
  const minutes = Number(match[2]) * 60 + Number(match[3]);
  return match[1] === "-" ? -minutes : minutes;
}

export function toExcelZonedDate(value, timezoneOffset = "+08:00") {
  if (!value) return "";
  const instant = new Date(value);
  if (Number.isNaN(instant.getTime())) return String(value);
  return new Date(instant.getTime() + offsetMinutes(timezoneOffset) * 60_000);
}

function sheetName(index) {
  return index === 0 ? "Behavior Logs" : `Behavior Logs ${index + 1}`;
}

export async function writeBehaviorWorkbook(rows, outputPath, { timezoneOffset = "+08:00" } = {}) {
  const workbook = Workbook.create();
  const sheetCount = Math.max(1, Math.ceil(rows.length / EXCEL_MAX_DATA_ROWS));

  for (let sheetIndex = 0; sheetIndex < sheetCount; sheetIndex += 1) {
    const sheet = workbook.worksheets.add(sheetName(sheetIndex));
    const start = sheetIndex * EXCEL_MAX_DATA_ROWS;
    const chunk = rows.slice(start, start + EXCEL_MAX_DATA_ROWS);

    sheet.showGridLines = false;
    sheet.freezePanes.freezeRows(1);
    sheet.getRange("A1:F1").values = [OUTPUT_COLUMNS];
    sheet.getRange("A1:F1").format = {
      fill: "#153B5B",
      font: { bold: true, color: "#FFFFFF" },
      rowHeight: 26,
      verticalAlignment: "center",
    };

    for (let offset = 0; offset < chunk.length; offset += WRITE_BATCH_SIZE) {
      const batch = chunk.slice(offset, offset + WRITE_BATCH_SIZE).map((row) =>
        OUTPUT_COLUMNS.map((column) => {
          if (column === "@timestamp") return toExcelZonedDate(row[column], timezoneOffset);
          const separator = column === "winlog.event_data.CommandLine" || column === "winlog_event_id" ? "\n" : " | ";
          return excelCell(row[column], separator);
        }),
      );
      sheet.getRangeByIndexes(offset + 1, 0, batch.length, OUTPUT_COLUMNS.length).values = batch;
    }

    const usedRows = chunk.length + 1;
    const usedRange = sheet.getRangeByIndexes(0, 0, usedRows, OUTPUT_COLUMNS.length);
    usedRange.format.font = { name: "Aptos", size: 10 };
    sheet.getRange("A1:F1").format.font = { name: "Aptos Display", size: 10, bold: true, color: "#FFFFFF" };

    const widths = [62, 25, 38, 42, 72, 34];
    widths.forEach((width, columnIndex) => {
      sheet.getRangeByIndexes(0, columnIndex, usedRows, 1).format.columnWidth = width;
    });
    if (chunk.length > 0) {
      sheet.getRangeByIndexes(1, 1, chunk.length, 1).format.numberFormat = 'm"月" d, yyyy hh:mm:ss';
      sheet.getRangeByIndexes(1, 0, chunk.length, 1).format.wrapText = true;
      sheet.getRangeByIndexes(1, 4, chunk.length, 2).format.wrapText = true;
      sheet.getRangeByIndexes(1, 0, chunk.length, OUTPUT_COLUMNS.length).format.verticalAlignment = "top";
      sheet.getRangeByIndexes(1, 0, chunk.length, OUTPUT_COLUMNS.length).format.rowHeight = 54;
    }

    if (chunk.length > 0) {
      const lastRow = chunk.length + 1;
      const table = sheet.tables.add(`A1:F${lastRow}`, true, `BehaviorLogs${sheetIndex + 1}`);
      table.style = "TableStyleMedium2";
      table.showBandedRows = true;
      table.showFilterButton = true;
    }
  }

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  const file = await SpreadsheetFile.exportXlsx(workbook);
  await file.save(outputPath);
  return { outputPath, rowCount: rows.length, sheetCount };
}
