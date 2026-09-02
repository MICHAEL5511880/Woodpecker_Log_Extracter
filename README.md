# Woodpecker XVR 行為分析 Log 匯出工具

此工具依據《Woodpecker-XVR 產品使用說明書 v5.5.10 v3》第 6.4 節 Public API 實作，會：

- 使用 `POST /config/v1/public/detection/search` 查詢行為分析資料。
- 再使用 `POST /config/v1/public/raw/search`，依 `enrichment.related_detection.id.keyword` 精確取得該行為分析「詳細資訊」底下的 Windows 日誌。
- 以 Bearer Token 驗證。
- 依 `@timestamp`（起訖皆包含）及 `host.hostname.keyword` 精確篩選。
- 自動循著 pagination cursor 抓完所有頁面。
- 依行為分析的 `@timestamp` 由舊到新排序，並以台灣時區顯示為 `3月 27, 2026 16:00:21` 這類格式。
- 每筆行為分析維持一列；其詳細日誌若有多個 CommandLine 或 Event ID，會去重並在同一儲存格內換行列出。
- 輸出 Excel；若超過單一工作表上限，會自動拆成多個工作表。
- 預設查詢 `config.json` 的 `defaultMachine`；可用 `--machine` 臨時覆寫。

Excel 欄位順序固定為：

1. `winlog.event_data.CommandLine`（取自「詳細資訊」關聯 Windows 日誌的 `_source.winlog.event_data.CommandLine`）
2. `@timestamp`
3. `detection.information.name`
4. `detection.information.tags`
5. `detection.information.detail`
6. `winlog_event_id`（取自關聯 Windows 日誌的 `_source.winlog.event_id`）

若某筆行為分析沒有關聯的 Windows 日誌，仍會保留該列，CommandLine 與 `winlog_event_id` 留白。

## 第一次設定

1. 複製 `config.example.json` 為 `config.json`。
2. 將 `baseUrl` 改成 XVR Server 網址，例如 `https://192.168.1.10`。
3. 將 `defaultMachine` 設成平常查詢的機器名稱，例如 `Workstation4`。
4. 建議不要把 Token 寫進檔案；可在 PowerShell 設定目前視窗專用的環境變數：

```powershell
$env:WOODPECKER_TOKEN = "pat#你的Token"
```

也可以把 Token 填入 `config.json` 的 `token`，但請妥善保護該檔案。

Token 的建立方式在手冊 6.4.1；可使用的角色包含 `admin`、`auditor`、`manager`、`help_desk`、`user`。

## 使用方式

互動輸入時間與機器名稱：

```powershell
.\run-export.ps1
```

一次帶入全部查詢值：

```powershell
.\run-export.ps1 --start "2026-09-01 00:00:00" --end "2026-09-01 23:59:59" --machine "DESKTOP-EXAMPLE"
```

未指定 `--machine` 時會使用 `config.json` 的 `defaultMachine`。偶爾需要查別台機器時，只要用 `--machine` 臨時覆寫，不會更改預設值：

```powershell
.\run-export.ps1 --machine "Workstation5"
```

日期若未含時區，預設依 `config.json` 的 `timezoneOffset`（預設 `+08:00`）解讀。只輸入日期時，起始日從 00:00:00 算起，結束日到 23:59:59.999 為止。Excel 內的 `@timestamp` 也會依同一個 offset 轉換；目前設定 `+08:00` 即為台灣時間。

如內部 XVR Server 使用自簽憑證，確認連線目標可信後可加 `--insecure`：

```powershell
.\run-export.ps1 --insecure
```

預設輸出到 `Result`，也可指定完整檔名：

```powershell
.\run-export.ps1 --start "2026-09-01" --end "2026-09-02" --machine "PC-001" --output "D:\Reports\PC-001.xlsx"
```

## 安全注意事項

- `config.json`、`Result` 與 `node_modules` 已加入 `.gitignore`。
- 不要將 Token 貼進程式碼或分享至版本控制。
- `--insecure` 會停用 TLS 憑證驗證，只能用於你確定可信的內部伺服器。

## 測試

測試會啟動本機模擬 API，驗證 Bearer Token、兩階段 API、cursor 分頁、詳細日誌關聯、去重、台灣時間、欄位順序與 Excel 內容，不會連線到正式 XVR Server：

```powershell
.\run-export.ps1 --help
& "C:\Users\$env:USERNAME\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe" --test
```
