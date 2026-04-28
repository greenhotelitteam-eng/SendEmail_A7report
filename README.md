# SendEmail_A7report
 
## Wise SQL Server config for H04/H06

H04/H06 monthly occupancy before `2026/04/08` is read from Wise SQL Server `daytot`.
On every host that runs this program, set one of these sources:

1. Environment variables: `Conn_SQLSERVER_H04`, `Conn_SQLSERVER_H06`
2. `appsettings.local.json`:

```json
"Wise": {
  "ConnectionStrings": {
    "H04": "Data Source=...;Initial Catalog=HotelMis;User ID=...;Password=...;Integrated Security=False",
    "H06": "Data Source=...;Initial Catalog=HotelMis;User ID=...;Password=...;Integrated Security=False"
  },
  "LegacyConfigFile": "D:\\00.VSCode\\Sendmail_WiseReportTot\\bin\\Debug\\SendmailReportTot.local.config"
}
```

3. Legacy fallback file: `D:\00.VSCode\Sendmail_WiseReportTot\bin\Debug\SendmailReportTot.local.config`

If Wise connection strings are missing on another host, the `2026/03` monthly value for H04/H06 can show as `0%`.
Athena 報表 Email 發送工具。

這個專案不直接抓 Athena 網頁，而是讀取 `athena_report01` 與 `athena_setting`，組成 HTML Email 後寄出。

## 正確流程

一定要先抓資料，再發信：

1. 先執行 `GetAthenaA7`
2. 確認最新資料已寫入 `athena_report01`
3. 再執行 `SendEmail_A7report`

如果先寄信，會寄到舊資料。

## 資料來源

主要來自 `athena_report01`：

- `data_date`
- `hotel_id`
- `room_total`
- `room_count`
- `room_ooo`
- `room_revenue`
- `room_occupancy`
- `room_adr`
- `room_data`
- `data_lmtime`

輔助資料：

- `athena_info`
  - 讀飯店名稱等資料
- `athena_setting`
  - 讀寄件者與收件者設定
- Wise SQL Server `daytot`
  - 僅 H04 / H06 月住房率在 `2026/04/07` 以前使用
  - 連線字串優先讀環境變數 `Conn_SQLSERVER_H04` / `Conn_SQLSERVER_H06`
  - 若環境變數沒有值，會讀 `D:\00.VSCode\Sendmail_WiseReportTot\bin\Debug\SendmailReportTot.local.config`

## athena_setting 規則

### 寄件者

全館共用一筆：

- `func_name = 'email_sender'`
- `is_use = 'Y'`
- `hotel_id` 可固定填 `ALL`
- `value` 放 JSON

範例：

```json
{"Host":"smtp.example.com","Port":587,"EnableSsl":true,"Username":"sender@example.com","Password":"***","FromAddress":"sender@example.com","FromDisplayName":"Athena A7 Report"}
```

### 收件者

每館一筆：

- `hotel_id = 'H04'`、`'H06'`...
- `is_use = 'Y'`
- `value` 可放純字串或 JSON

預設群組：

- `func_name = 'email_recipient'`

早上 08:00 群組：

- `func_name = 'email_recipient_morning'`

純字串範例：

```text
a@example.com;b@example.com
```

JSON 範例：

```json
{"ToAddresses":["a@example.com"],"CcAddresses":["b@example.com"]}
```

## recipient-group 參數

目前支援用同一支程式切不同收件群組。

對應規則：

- `--recipient-group default`
  - 讀 `email_recipient`
- `--recipient-group morning`
  - 讀 `email_recipient_morning`

若不帶參數，預設就是 `default`。

## Email 規則

### 主旨

格式：

```text
{hotel_name}*本日:{今日入住}({今日住房率})*昨日:{昨日入住}({昨日住房率}) - {現在發信時間}
```

說明：

- 最後時間是發信當下時間
- 不是資料抓取時間
- 住房率顯示為四捨五入後的整數 `%`

### 明細區間

Email 下方明細目前顯示：

- 昨日
- 今日起算未來 120 天

### 月住房率摘要

上方「年月」摘要固定顯示基準日上個月起算 4 個月。

例如基準日是 `2026/04/28` 時，會顯示：

- `2026/03`
- `2026/04`
- `2026/05`
- `2026/06`

H04 / H06 在 `2026/04/08` 切到 A7：

- `2026/04/07` 以前從 Wise SQL Server `daytot` 計算
- `2026/04/08` 以後從 MySQL `athena_report01` 計算
- 跨系統月份會合併 Wise 與 A7 的住房數 / 總房數後重新計算住房率

### 明細欄位

表格目前使用：

- 日期
- 星期
- 入住
- 庫存
- 修理參觀
- 住房率
- 各房型
- 金額
- ADR

其中：

- `庫存 = room_total - room_count`
- `修理參觀 = room_ooo`

### 房型欄位說明

Email 內顯示：

```text
☆ 房型6數字: 用房/修參/鎖控/官網/空房/總數
```

`room_data` 會完整顯示 6 個數字，不會裁掉。

### 住房率顏色規則

目前規則：

- `0~29%`：紫色
- `30~59%`：紅色
- `60~79%`：黃色
- `80~89%`：綠色
- `90%以上`：白色

Email 最下方也會顯示：

```text
☆ 住房率0~29%為紫色、30~59%為紅色、60~79%為黃色、80~89%為綠色、90%以上為白色
```

## 執行方式

預覽，不寄信：

```powershell
dotnet run -- --preview-only --hotel H06 --date 2026/04/22
```

正常寄信：

```powershell
dotnet run
```

寄給早上群組：

```powershell
dotnet run -- --recipient-group morning
```

直接跑編譯檔：

```powershell
dotnet .\bin\Debug\net5.0\SendEmail_A7report.dll
```

## 設定檔

本機仍可保留：

- `appsettings.local.json`
- `appsettings.local.json.example`

但目前寄件者與收件者主要以資料表 `athena_setting` 為準。

## 預覽輸出

HTML 預覽會輸出到：

- `bin\Debug\net5.0\output`

例如：

- `H04_20260422.html`
- `H06_20260422.html`

## 可攜版

可攜版資料夾：

- `dist\SendEmail_A7report_portable`

重新產生可攜版：

```powershell
dotnet publish -c Release -r win-x64 --self-contained true -o .\dist\SendEmail_A7report_portable
```

這一包比 `GetAthenaA7` 簡單：

- 不需要 Playwright
- 不需要 Chromium
- 不需要 Python

只要該主機可以：

- 連 MySQL
- H04 / H06 若要計算 2026/04/07 以前的月住房率，需能連 Wise SQL Server
- 連 SMTP

就能執行。

## 排程建議

工作排程器建議拆成兩個步驟：

1. 先跑 `GetAthenaA7`
2. 再跑 `SendEmail_A7report`

不要同時跑。

若要每天早上 `08:00` 寄給另一組收件者，可另外建一個排程：

```powershell
SendEmail_A7report.exe --recipient-group morning
```

## 主要檔案

- `Program.vb`
  - 主流程與 Email HTML 組裝
- `appsettings.local.json.example`
  - 本機設定範例
