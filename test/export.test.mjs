import assert from "node:assert/strict";
import fs from "node:fs/promises";
import http from "node:http";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";
import {
  DETECTION_FIELD_PATHS,
  RAW_FIELD_PATHS,
  fetchBehaviorLogs,
  normalizeTime,
} from "../src/woodpecker-export.mjs";
import { OUTPUT_COLUMNS, toExcelZonedDate, writeBehaviorWorkbook } from "../src/xlsx-export.mjs";

test("normalizeTime interprets local Taiwan time and date-only end", () => {
  assert.equal(normalizeTime("2026-09-02 08:30:00", "+08:00"), "2026-09-02T00:30:00.000Z");
  assert.equal(normalizeTime("2026-09-02", "+08:00", true), "2026-09-02T15:59:59.999Z");
  assert.equal(toExcelZonedDate("2026-03-27T08:00:21.000Z", "+08:00").toISOString(), "2026-03-27T16:00:21.000Z");
});

test("fetches related detail logs, deduplicates values, and writes Taiwan-time Excel", async (context) => {
  const detectionBodies = [];
  const rawBodies = [];
  const server = http.createServer((request, response) => {
    const chunks = [];
    request.on("data", (chunk) => chunks.push(chunk));
    request.on("end", () => {
      assert.equal(request.headers.authorization, "Bearer test-token");
      const body = JSON.parse(Buffer.concat(chunks).toString("utf8"));
      let json;

      if (request.url === "/config/v1/public/detection/search") {
        detectionBodies.push(body);
        json = !body.pagination?.cursor
          ? {
              success: true,
              data: {
                total: 3,
                pagination: { cursor: "detection-cursor-2", has_more: true },
                warnings: [],
                data: [
                  {
                    detection_id: "det-2",
                    "@timestamp": "2026-03-27T09:00:00.000Z",
                    "detection.information.name": "Second",
                    "detection.information.tags": ["Execution"],
                    "detection.information.detail": "detail 2",
                    related_event_index: ["xvrwineventcollector-fc-2026-03-000001"],
                  },
                  {
                    detection_id: "det-1",
                    "@timestamp": "2026-03-27T08:00:21.000Z",
                    "detection.information.name": "First",
                    "detection.information.tags": ["Discovery", "Execution"],
                    "detection.information.detail": { source: "mock" },
                    related_event_index: "xvrwineventcollector-fc-2026-03-000001",
                  },
                ],
              },
            }
          : {
              success: true,
              data: {
                total: 3,
                pagination: { cursor: "detection-cursor-2", has_more: false },
                warnings: [],
                data: [
                  {
                    detection_id: "det-3",
                    "@timestamp": "2026-03-27T10:00:00.000Z",
                    "detection.information.name": "No Windows details",
                    "detection.information.tags": [],
                    "detection.information.detail": "detail 3",
                    related_event_index: ["xvrnetworkevent-fc-2026-03-000001"],
                  },
                ],
              },
            };
      } else if (request.url === "/config/v1/public/raw/search") {
        rawBodies.push(body);
        const queriedId = body.query?.filters?.[0]?.terms?.["enrichment.related_detection.id.keyword"]?.[0];
        if (queriedId === "det-1" && !body.pagination?.cursor) {
          json = {
              success: true,
              data: {
                total: 3,
                pagination: { cursor: "raw-cursor-2", has_more: true },
                warnings: [],
                data: [
                  {
                    raw_id: "raw-2",
                    raw_timestamp: "2026-03-27T08:00:02.000Z",
                    command_line: "powershell -enc AAA   ",
                    winlog_event_id: 4103,
                  },
                  {
                    raw_id: "raw-1",
                    raw_timestamp: "2026-03-27T08:00:01.000Z",
                    command_line: "powershell -enc AAA",
                    winlog_event_id: "4103",
                  },
                ],
              },
            };
        } else if (body.pagination?.cursor === "raw-cursor-2") {
          json = {
              success: true,
              data: {
                total: 3,
                pagination: { cursor: "raw-cursor-2", has_more: false },
                warnings: [],
                data: [
                  {
                    raw_id: "raw-3",
                    raw_timestamp: "2026-03-27T08:00:03.000Z",
                    command_line: ["cmd /c whoami"],
                    winlog_event_id: 4104,
                  },
                ],
              },
            };
        } else if (queriedId === "det-2") {
          json = {
              success: true,
              data: {
                total: 2,
                pagination: { has_more: false },
                warnings: [],
                data: [
                  {
                    raw_id: "raw-2",
                    raw_timestamp: "2026-03-27T08:00:02.000Z",
                    command_line: "powershell -enc AAA",
                    winlog_event_id: 4103,
                  },
                  {
                    raw_id: "raw-4",
                    raw_timestamp: "2026-03-27T08:00:04.000Z",
                    command_line: "cmd /c hostname",
                    winlog_event_id: 4688,
                  },
                ],
              },
            };
        } else {
          throw new Error(`Unexpected raw query: ${JSON.stringify(body)}`);
        }
      } else {
        response.writeHead(404);
        response.end();
        return;
      }

      response.writeHead(200, { "Content-Type": "application/json" });
      response.end(JSON.stringify(json));
    });
  });
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  context.after(() => server.close());

  const address = server.address();
  const progress = [];
  const result = await fetchBehaviorLogs({
    baseUrl: `http://127.0.0.1:${address.port}`,
    token: "test-token",
    startIso: "2026-03-27T00:00:00.000Z",
    endIso: "2026-03-27T23:59:59.999Z",
    machine: "TEST-PC",
    pageSize: 2,
    onProgress: (item) => progress.push(item),
  });

  assert.equal(result.detectionPages, 2);
  assert.equal(result.rawPages, 3);
  assert.equal(result.rawMatchedCount, 5);
  assert.deepEqual(result.rows.map((row) => row["detection.information.name"]), ["First", "Second", "No Windows details"]);
  assert.equal(result.rows[0]["winlog.event_data.CommandLine"], "powershell -enc AAA\ncmd /c whoami");
  assert.equal(result.rows[0].winlog_event_id, "4103\n4104");
  assert.equal(result.rows[1]["winlog.event_data.CommandLine"], "powershell -enc AAA\ncmd /c hostname");
  assert.equal(result.rows[1].winlog_event_id, "4103\n4688");
  assert.equal(result.rows[2]["winlog.event_data.CommandLine"], "");
  assert.equal(result.rows[2].winlog_event_id, "");

  assert.deepEqual(detectionBodies[0].query.fields, DETECTION_FIELD_PATHS);
  assert.equal(detectionBodies[0].query.filters[1].term["host.hostname.keyword"], "TEST-PC");
  assert.equal(detectionBodies[1].pagination.cursor, "detection-cursor-2");
  assert.equal(detectionBodies[1].query, undefined);
  assert.equal(rawBodies[0].query.category, "Windows Events");
  assert.deepEqual(rawBodies[0].query.fields, RAW_FIELD_PATHS);
  assert.deepEqual(rawBodies[0].query.filters[0].terms["enrichment.related_detection.id.keyword"], ["det-1"]);
  assert.equal(rawBodies[0].query.filters[1].term["host.hostname.keyword"], "TEST-PC");
  assert.equal(rawBodies[1].pagination.cursor, "raw-cursor-2");
  assert.equal(rawBodies[1].query, undefined);
  assert.deepEqual(rawBodies[2].query.filters[0].terms["enrichment.related_detection.id.keyword"], ["det-2"]);
  assert.ok(progress.some((item) => item.stage === "detections"));
  assert.ok(progress.some((item) => item.stage === "raw"));

  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "woodpecker-export-"));
  const outputPath = path.join(tempDir, "verified.xlsx");
  const preserveQa = process.env.WOODPECKER_KEEP_QA === "1";
  if (!preserveQa) context.after(() => fs.rm(tempDir, { recursive: true, force: true }));
  await writeBehaviorWorkbook(result.rows, outputPath, { timezoneOffset: "+08:00" });

  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(outputPath));
  const sheet = workbook.worksheets.getItem("Behavior Logs");
  const values = sheet.getRange("A1:F4").values;
  assert.deepEqual(values[0], OUTPUT_COLUMNS);
  assert.equal(values[1][0], "powershell -enc AAA\ncmd /c whoami");
  assert.equal(values[1][3], "Discovery | Execution");
  assert.equal(values[1][5], "4103\n4104");

  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 30 },
    summary: "final formula error scan",
  });
  assert.doesNotMatch(errors.ndjson, /#REF!|#DIV\/0!|#VALUE!|#NAME\?|#N\/A/);

  const preview = await workbook.render({
    sheetName: "Behavior Logs",
    range: "A1:F4",
    scale: 1.5,
    format: "png",
  });
  await fs.writeFile(path.join(tempDir, "verified-preview.png"), new Uint8Array(await preview.arrayBuffer()));
  if (preserveQa) console.log(`QA_DIR=${tempDir}`);
});
