#!/usr/bin/env node
import fs from "node:fs/promises";
import path from "node:path";
import http from "node:http";
import https from "node:https";
import readline from "node:readline/promises";
import { stdin as input, stdout as output } from "node:process";
import { pathToFileURL } from "node:url";
import { OUTPUT_COLUMNS, writeBehaviorWorkbook } from "./xlsx-export.mjs";

export const API_PATH = "/config/v1/public/detection/search";
export const RAW_API_PATH = "/config/v1/public/raw/search";
export const DETECTION_FIELD_PATHS = {
  detection_id: "_id",
  "@timestamp": "_source.@timestamp",
  "detection.information.name": "_source.detection.information.name",
  "detection.information.tags": "_source.detection.information.tags",
  "detection.information.detail": "_source.detection.information.detail",
  related_event_index: "_source.enrichment.related_event_index",
};
export const RAW_FIELD_PATHS = {
  raw_id: "_id",
  raw_timestamp: "_source.@timestamp",
  command_line: "_source.winlog.event_data.CommandLine",
  winlog_event_id: "_source.winlog.event_id",
};

// This server accepts the related-detection field as a filter but does not expose
// it as a projected field. Query one detection at a time so every raw row can be
// assigned without guessing.
const RAW_BATCH_SIZE = 1;
const EXCEL_CELL_SAFE_LENGTH = 32_000;
const TRUNCATION_MARKER = "\n… [內容過長，已截斷]";

function parseArgs(argv) {
  const result = {};
  const booleanKeys = new Set(["insecure", "non-interactive", "help"]);
  for (let index = 0; index < argv.length; index += 1) {
    const item = argv[index];
    if (!item.startsWith("--")) throw new Error(`不支援的參數：${item}`);
    const key = item.slice(2);
    if (booleanKeys.has(key)) {
      result[key] = true;
      continue;
    }
    const value = argv[index + 1];
    if (!value || value.startsWith("--")) throw new Error(`參數 --${key} 缺少值。`);
    result[key] = value;
    index += 1;
  }
  return result;
}

function showHelp() {
  console.log(`Woodpecker XVR 行為分析 Log 匯出工具

使用方式：
  .\\run-export.ps1
  .\\run-export.ps1 --start "2026-09-01 00:00:00" --end "2026-09-01 23:59:59" --machine "PC-001"

參數：
  --start <時間>          起始時間（含）
  --end <時間>            結束時間（含）
  --machine <名稱>        覆寫 config.json 的 defaultMachine
  --url <URL>             XVR Server URL
  --token <Token>         Public API Bearer Token
  --output <路徑>         輸出 .xlsx 路徑
  --config <路徑>         設定檔，預設 config.json
  --timezone <+08:00>     未含時區之輸入時間所使用的 offset
  --page-size <1-10000>   單頁筆數，預設 1000
  --insecure              允許自簽憑證（會降低 TLS 安全性）
  --non-interactive       缺少參數時直接報錯，不顯示互動提示
  --help                  顯示說明`);
}

async function loadConfig(configPath) {
  try {
    return JSON.parse(await fs.readFile(configPath, "utf8"));
  } catch (error) {
    if (error.code === "ENOENT") return {};
    throw new Error(`無法讀取設定檔 ${configPath}：${error.message}`);
  }
}

function hasTimeZone(value) {
  return /(?:Z|[+-]\d{2}:?\d{2})$/i.test(value.trim());
}

function validateOffset(offset) {
  if (!/^[+-](?:0\d|1\d|2[0-3]):[0-5]\d$/.test(offset)) {
    throw new Error(`timezoneOffset 格式錯誤：${offset}，應為 +08:00 這類格式。`);
  }
  return offset;
}

export function normalizeTime(value, timezoneOffset = "+08:00", isEnd = false) {
  const raw = value.trim();
  const offset = validateOffset(timezoneOffset);
  let candidate = raw;
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    candidate = `${raw}T${isEnd ? "23:59:59.999" : "00:00:00"}${offset}`;
  } else if (!hasTimeZone(raw)) {
    candidate = `${raw.replace(" ", "T")}${offset}`;
  }
  const parsed = new Date(candidate);
  if (Number.isNaN(parsed.getTime())) throw new Error(`無法解析時間：${value}`);
  return parsed.toISOString();
}

function safeFilename(value) {
  return value.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_").replace(/\s+/g, "_").slice(0, 80);
}

function defaultOutputPath(outputDir, machine, startIso, endIso) {
  const compact = (iso) => iso.replace(/[-:]/g, "").replace(/\.\d{3}Z$/, "Z");
  return path.resolve(outputDir, `Woodpecker_${safeFilename(machine)}_${compact(startIso)}_${compact(endIso)}.xlsx`);
}

function requestJsonOnce(url, token, body, options) {
  const transport = url.protocol === "https:" ? https : http;
  const payload = JSON.stringify(body);
  return new Promise((resolve, reject) => {
    const request = transport.request(
      url,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json",
          "Content-Type": "application/json",
          "Content-Length": Buffer.byteLength(payload),
        },
        rejectUnauthorized: options.rejectUnauthorized,
        timeout: options.timeoutMs,
      },
      (response) => {
        const chunks = [];
        response.on("data", (chunk) => chunks.push(chunk));
        response.on("end", () => {
          const text = Buffer.concat(chunks).toString("utf8");
          let json;
          try {
            json = text ? JSON.parse(text) : {};
          } catch {
            reject(new Error(`API 回傳非 JSON（HTTP ${response.statusCode}）：${text.slice(0, 300)}`));
            return;
          }
          resolve({ statusCode: response.statusCode ?? 0, json });
        });
      },
    );
    request.on("timeout", () => request.destroy(new Error(`API 請求超過 ${options.timeoutMs / 1000} 秒`)));
    request.on("error", reject);
    request.end(payload);
  });
}

const RETRY_STATUS = new Set([429, 502, 503, 504]);
const wait = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds));

async function requestJson(url, token, body, options) {
  let lastError;
  for (let attempt = 0; attempt < 4; attempt += 1) {
    try {
      const response = await requestJsonOnce(url, token, body, options);
      if (!RETRY_STATUS.has(response.statusCode) || attempt === 3) return response;
      lastError = new Error(`API 暫時無法使用（HTTP ${response.statusCode}）`);
    } catch (error) {
      lastError = error;
      if (attempt === 3) throw error;
    }
    await wait(1000 * 2 ** attempt);
  }
  throw lastError;
}

function apiError(statusCode, json) {
  const details = json.error || json.message || JSON.stringify(json).slice(0, 500);
  return new Error(`Woodpecker API 查詢失敗（HTTP ${statusCode}）：${details}`);
}

function normalizeDetectionRow(row) {
  return {
    detection_id: row?.detection_id ?? "",
    "@timestamp": row?.["@timestamp"] ?? "",
    "detection.information.name": row?.["detection.information.name"] ?? "",
    "detection.information.tags": row?.["detection.information.tags"] ?? "",
    "detection.information.detail": row?.["detection.information.detail"] ?? "",
    related_event_index: row?.related_event_index ?? "",
  };
}

function scalarStrings(value) {
  if (value === null || value === undefined || value === "") return [];
  if (Array.isArray(value)) return value.flatMap(scalarStrings);
  if (typeof value === "object") return [JSON.stringify(value)];
  return [String(value)];
}

function normalizeCellValue(value) {
  return String(value).replace(/\r\n?/g, "\n").replace(/[ \t]+$/gm, "").trim();
}

function hasWindowsEventsIndex(value) {
  return scalarStrings(value).some((indexName) => indexName.startsWith("xvrwineventcollector-fc"));
}

function aggregateCell(values, detectionId, field, warnings) {
  const joined = values.join("\n");
  if (joined.length <= EXCEL_CELL_SAFE_LENGTH) return joined;
  warnings.add(`行為分析 ${detectionId} 的 ${field} 超過 Excel 儲存格限制，內容已截斷。`);
  return `${joined.slice(0, EXCEL_CELL_SAFE_LENGTH - TRUNCATION_MARKER.length)}${TRUNCATION_MARKER}`;
}

export async function fetchBehaviorLogs({
  baseUrl,
  token,
  startIso,
  endIso,
  machine,
  pageSize = 1000,
  rejectUnauthorized = true,
  onProgress = () => {},
}) {
  const normalizedBaseUrl = baseUrl.endsWith("/") ? baseUrl : `${baseUrl}/`;
  const detectionUrl = new URL(API_PATH, normalizedBaseUrl);
  const rawUrl = new URL(RAW_API_PATH, normalizedBaseUrl);
  const detections = [];
  const warnings = new Set();
  const cursors = new Set();
  let cursor = null;
  let page = 0;
  let expectedTotal = null;

  while (true) {
    const body = cursor
      ? { pagination: { cursor, keep_alive: "5m" } }
      : {
          query: {
            filters: [
              { range: { "@timestamp": { gte: startIso, lte: endIso } } },
              { term: { "host.hostname.keyword": machine } },
            ],
            sort: [{ "@timestamp": { order: "asc" } }],
            size: pageSize,
            fields: DETECTION_FIELD_PATHS,
          },
          pagination: { keep_alive: "5m" },
        };

    const { statusCode, json } = await requestJson(detectionUrl, token, body, {
      rejectUnauthorized,
      timeoutMs: 120_000,
    });
    if (statusCode < 200 || statusCode >= 300 || json.success !== true) throw apiError(statusCode, json);

    const data = json.data;
    if (!data || !Array.isArray(data.data)) throw new Error("API 回傳格式不符手冊：缺少 data.data 陣列。");
    expectedTotal ??= Number.isFinite(data.total) ? data.total : null;
    data.warnings
      ?.map((warning) => String(warning).trim())
      .filter(Boolean)
      .forEach((warning) => warnings.add(warning));
    detections.push(...data.data.map(normalizeDetectionRow));
    page += 1;
    onProgress({ stage: "detections", page, fetched: detections.length, total: expectedTotal });

    if (!data.pagination?.has_more) break;
    const nextCursor = data.pagination.cursor;
    if (!nextCursor) throw new Error("API 表示仍有後續資料，但沒有回傳 pagination.cursor。");
    if (cursors.has(nextCursor)) throw new Error("API 重複回傳相同 cursor，已停止以避免無限迴圈。");
    cursors.add(nextCursor);
    cursor = nextCursor;
  }

  detections.sort((left, right) => {
    const leftTime = Date.parse(left["@timestamp"]);
    const rightTime = Date.parse(right["@timestamp"]);
    if (Number.isNaN(leftTime) && Number.isNaN(rightTime)) return 0;
    if (Number.isNaN(leftTime)) return 1;
    if (Number.isNaN(rightTime)) return -1;
    return leftTime - rightTime;
  });

  if (expectedTotal !== null && detections.length !== expectedTotal) {
    warnings.add(`API 宣告 total=${expectedTotal}，實際取得 ${detections.length} 筆；請確認查詢期間資料是否持續變動。`);
  }

  const detectionIds = detections
    .filter((row) => row.detection_id && hasWindowsEventsIndex(row.related_event_index))
    .map((row) => String(row.detection_id));
  const uniqueDetectionIds = [...new Set(detectionIds)];
  const rawByDetection = new Map(
    detections
      .filter((row) => row.detection_id)
      .map((row) => [String(row.detection_id), { commands: [], commandSet: new Set(), eventIds: [], eventIdSet: new Set() }]),
  );
  const rawRows = [];
  let rawPages = 0;
  let rawMatchedCount = 0;
  const batchCount = Math.ceil(uniqueDetectionIds.length / RAW_BATCH_SIZE);

  for (let batchIndex = 0; batchIndex < batchCount; batchIndex += 1) {
    const ids = uniqueDetectionIds.slice(batchIndex * RAW_BATCH_SIZE, (batchIndex + 1) * RAW_BATCH_SIZE);
    let rawCursor = null;
    const rawCursors = new Set();
    const seenRawIdsInBatch = new Set();
    let rawPageInBatch = 0;

    while (true) {
      const body = rawCursor
        ? { pagination: { cursor: rawCursor, keep_alive: "5m" } }
        : {
            query: {
              category: "Windows Events",
              filters: [
                { terms: { "enrichment.related_detection.id.keyword": ids } },
                { term: { "host.hostname.keyword": machine } },
              ],
              sort: [{ "@timestamp": { order: "asc" } }, { _id: { order: "asc" } }],
              size: pageSize,
              fields: RAW_FIELD_PATHS,
            },
            pagination: { keep_alive: "5m" },
          };

      const { statusCode, json } = await requestJson(rawUrl, token, body, {
        rejectUnauthorized,
        timeoutMs: 120_000,
      });
      if (statusCode < 200 || statusCode >= 300 || json.success !== true) throw apiError(statusCode, json);
      const data = json.data;
      if (!data || !Array.isArray(data.data)) throw new Error("原始日誌 API 回傳格式不符手冊：缺少 data.data 陣列。");
      data.warnings
        ?.map((warning) => String(warning).trim())
        .filter(Boolean)
        .forEach((warning) => warnings.add(warning));

      for (const rawRow of data.data) {
        const rawId = rawRow?.raw_id ? String(rawRow.raw_id) : "";
        if (rawId && seenRawIdsInBatch.has(rawId)) continue;
        if (rawId) seenRawIdsInBatch.add(rawId);
        rawRows.push({ ...rawRow, queried_detection_ids: ids });
        rawMatchedCount += 1;
      }
      rawPages += 1;
      rawPageInBatch += 1;
      onProgress({
        stage: "raw",
        batch: batchIndex + 1,
        batches: batchCount,
        page: rawPageInBatch,
        fetched: rawMatchedCount,
      });

      if (!data.pagination?.has_more) break;
      const nextCursor = data.pagination.cursor;
      if (!nextCursor) throw new Error("原始日誌 API 表示仍有後續資料，但沒有回傳 pagination.cursor。");
      if (rawCursors.has(nextCursor)) throw new Error("原始日誌 API 重複回傳相同 cursor，已停止以避免無限迴圈。");
      rawCursors.add(nextCursor);
      rawCursor = nextCursor;
    }
  }

  rawRows.sort((left, right) => {
    const timeDifference = (Date.parse(left?.raw_timestamp) || 0) - (Date.parse(right?.raw_timestamp) || 0);
    return timeDifference || String(left?.raw_id ?? "").localeCompare(String(right?.raw_id ?? ""));
  });

  for (const rawRow of rawRows) {
    const commands = scalarStrings(rawRow?.command_line).map(normalizeCellValue).filter(Boolean);
    const eventIds = scalarStrings(rawRow?.winlog_event_id).map(normalizeCellValue).filter(Boolean);
    for (const detectionId of rawRow.queried_detection_ids) {
      const aggregate = rawByDetection.get(detectionId);
      if (!aggregate) continue;
      for (const command of commands) {
        if (aggregate.commandSet.has(command)) continue;
        aggregate.commandSet.add(command);
        aggregate.commands.push(command);
      }
      for (const eventId of eventIds) {
        if (aggregate.eventIdSet.has(eventId)) continue;
        aggregate.eventIdSet.add(eventId);
        aggregate.eventIds.push(eventId);
      }
    }
  }

  const rows = detections.map((detection) => {
    const detectionId = String(detection.detection_id || "");
    const aggregate = rawByDetection.get(detectionId);
    return {
      "winlog.event_data.CommandLine": aggregate
        ? aggregateCell(aggregate.commands, detectionId, "winlog.event_data.CommandLine", warnings)
        : "",
      "@timestamp": detection["@timestamp"],
      "detection.information.name": detection["detection.information.name"],
      "detection.information.tags": detection["detection.information.tags"],
      "detection.information.detail": detection["detection.information.detail"],
      winlog_event_id: aggregate ? aggregateCell(aggregate.eventIds, detectionId, "winlog_event_id", warnings) : "",
    };
  });

  return {
    rows,
    warnings: [...warnings],
    expectedTotal,
    pages: page,
    detectionPages: page,
    rawPages,
    rawMatchedCount,
  };
}

async function promptMissing(settings) {
  const terminal = readline.createInterface({ input, output });
  try {
    settings.baseUrl ||= (await terminal.question("Woodpecker XVR Server URL：")).trim();
    settings.token ||= (await terminal.question("Public API Token（輸入時會顯示於畫面）：")).trim();
    settings.start ||= (await terminal.question("起始時間（含，例如 2026-09-01 00:00:00）：")).trim();
    settings.end ||= (await terminal.question("結束時間（含，例如 2026-09-01 23:59:59）：")).trim();
    settings.machine ||= (await terminal.question("機器名稱（精確 Hostname）：")).trim();
  } finally {
    terminal.close();
  }
}

function required(settings, key, label) {
  if (!settings[key]) throw new Error(`缺少 ${label}；請使用 --${key} 或設定檔／環境變數提供。`);
}

export async function main(argv = process.argv.slice(2)) {
  const args = parseArgs(argv);
  if (args.help) {
    showHelp();
    return;
  }

  const configPath = path.resolve(args.config || "config.json");
  const config = await loadConfig(configPath);
  const settings = {
    baseUrl: args.url || process.env.WOODPECKER_URL || config.baseUrl,
    token: args.token || process.env.WOODPECKER_TOKEN || config.token,
    start: args.start,
    end: args.end,
    machine: args.machine || process.env.WOODPECKER_MACHINE || config.defaultMachine || "Workstation4",
    timezoneOffset: args.timezone || process.env.WOODPECKER_TIMEZONE || config.timezoneOffset || "+08:00",
    pageSize: Number(args["page-size"] || process.env.WOODPECKER_PAGE_SIZE || config.pageSize || 1000),
    rejectUnauthorized: args.insecure ? false : config.rejectUnauthorized !== false,
    outputDir: config.outputDir || "Result",
    output: args.output,
  };

  if (!args["non-interactive"]) await promptMissing(settings);
  required(settings, "baseUrl", "XVR Server URL");
  required(settings, "token", "Public API Token");
  required(settings, "start", "起始時間");
  required(settings, "end", "結束時間");
  required(settings, "machine", "機器名稱");
  if (!Number.isInteger(settings.pageSize) || settings.pageSize < 1 || settings.pageSize > 10_000) {
    throw new Error("pageSize 必須是 1 到 10000 的整數。");
  }

  const startIso = normalizeTime(settings.start, settings.timezoneOffset, false);
  const endIso = normalizeTime(settings.end, settings.timezoneOffset, true);
  if (Date.parse(startIso) > Date.parse(endIso)) throw new Error("起始時間不可晚於結束時間。");
  const outputPath = path.resolve(
    settings.output || defaultOutputPath(settings.outputDir, settings.machine, startIso, endIso),
  );

  console.log(`查詢機器：${settings.machine}`);
  console.log(`查詢區間：${startIso} ～ ${endIso}（含）`);
  if (!settings.rejectUnauthorized) console.warn("警告：已允許未受信任的 TLS 憑證。僅應用於可信任的內部 XVR Server。");

  const result = await fetchBehaviorLogs({
    ...settings,
    startIso,
    endIso,
    onProgress: (progress) => {
      if (progress.stage === "detections") {
        console.log(
          `行為分析第 ${progress.page} 頁完成，累計 ${progress.fetched}${progress.total === null ? "" : ` / ${progress.total}`} 筆`,
        );
      } else {
        console.log(
          `詳細日誌批次 ${progress.batch} / ${progress.batches}、第 ${progress.page} 頁完成，累計 ${progress.fetched} 筆`,
        );
      }
    },
  });
  const workbook = await writeBehaviorWorkbook(result.rows, outputPath, {
    timezoneOffset: settings.timezoneOffset,
  });

  result.warnings.forEach((warning) => console.warn(`API 警告：${warning}`));
  console.log(`完成：${workbook.rowCount} 筆、${workbook.sheetCount} 個工作表`);
  console.log(`Excel：${workbook.outputPath}`);
}

const invokedDirectly = process.argv[1] && import.meta.url === pathToFileURL(path.resolve(process.argv[1])).href;
if (invokedDirectly) {
  main().catch((error) => {
    console.error(`錯誤：${error.message}`);
    process.exitCode = 1;
  });
}
