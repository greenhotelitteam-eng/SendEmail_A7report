# SendEmail_A7report Project Notes

這份文件只放下次接手時最重要的已確認規則。

## 專案目的

從 `athena_report01` 讀取 Athena 報表資料，組成 HTML Email 並寄出。

## 最重要流程

寄信前一定要先完成抓資料：

1. 先跑 `GetAthenaA7`
2. 等 `athena_report01` 更新完成
3. 再跑 `SendEmail_A7report`

如果順序錯，寄出去的會是舊資料。

## 資料來源

### athena_report01

主要使用：

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

### athena_info

- 用來取 `hotel_name`

### athena_setting

- 用來取 SMTP 寄件者
- 用來取每館不同收件者

### Wise SQL Server daytot

- 只用在 H04 / H06 的月住房率摘要
- H04 / H06 在 `2026/04/08` 切到 A7
- `2026/04/07` 以前的月資料需從 Wise SQL Server `daytot` 補回
- 連線字串優先讀環境變數 `Conn_SQLSERVER_H04` / `Conn_SQLSERVER_H06`
- 若環境變數沒有值，會讀 `D:\00.VSCode\Sendmail_WiseReportTot\bin\Debug\SendmailReportTot.local.config`

## athena_setting 規則

### 寄件者

- `func_name = 'email_sender'`
- 全館共用一筆
- `is_use = 'Y'`
- `value` 放 JSON

### 收件者

預設群組：

- `func_name = 'email_recipient'`

早上群組：

- `func_name = 'email_recipient_morning'`

規則：

- 每館一筆
- `hotel_id` 區分不同飯店
- `is_use = 'Y'`
- `value` 可放純收件名單或 JSON

## recipient-group 參數

同一支程式可用不同參數切換收件群組。

規則：

- `default`
  - 對應 `email_recipient`
- `morning`
  - 對應 `email_recipient_morning`

若之後還要新增：

- `night`
- `manager`

就延伸成：

- `email_recipient_night`
- `email_recipient_manager`

## 主旨規則

格式：

`{hotel_name}*本日:{今日入住}({今日住房率})*昨日:{昨日入住}({昨日住房率}) - {現在發信時間}`

注意：

- 最後時間要用發信當下時間
- 不要用資料抓取時間
- 住房率顯示為四捨五入後整數 `%`

## 明細區間規則

Email 下方明細要顯示：

- 昨日
- 今日起算未來 120 天

## 月摘要規則

上方「年月」摘要顯示基準日上個月起算 4 個月。

例：基準日 `2026/04/28` 要顯示：

- `2026/03`
- `2026/04`
- `2026/05`
- `2026/06`

H04 / H06 月住房率要注意 Wise / A7 切換：

- `2026/04/07` 以前：Wise SQL Server `daytot`
- `2026/04/08` 以後：MySQL `athena_report01`
- 跨系統月份要合併住房數與總房數後重新計算，不可直接平均兩邊百分比

## 表格欄位規則

目前下方明細使用：

- 日期
- 星期
- 入住
- 庫存
- 修理參觀
- 住房率
- 房型欄位
- 金額
- ADR

## 欄位計算規則

- `庫存 = room_total - room_count`
- `修理參觀 = room_ooo`

## 房型說明規則

Email 內文要顯示：

`☆ 房型6數字: 用房/修參/鎖控/官網/空房/總數`

`room_data` 顯示規則：

- 每個房型保留完整 6 個數字
- 不裁掉

## 住房率顏色規則

- `0~29%`：紫色
- `30~59%`：紅色
- `60~79%`：黃色
- `80~89%`：綠色
- `90%以上`：白色

底部說明文字要和實際規則一致。

## 版面規則

- 內容字體已調小
- 上方兩個摘要表格欄寬一致
- HTML 預覽輸出到 `bin\Debug\net5.0\output`

## 可攜版

- `dist\SendEmail_A7report_portable`
- 不需要 Python / Playwright
- 只要能連 MySQL 與 SMTP 即可
- H04 / H06 若要補 `2026/04/07` 以前月住房率，還需要 Wise SQL Server 連線
- 更新可攜版使用：
  `dotnet publish -c Release -r win-x64 --self-contained true -o .\dist\SendEmail_A7report_portable`

## 下次修改時優先確認

1. `athena_report01` 是否已有最新資料
2. `athena_setting` 的 sender / recipient 是否啟用
3. `recipient-group` 對應的 `func_name` 是否正確
4. 主旨時間是否仍使用發信當下時間
5. 顏色規則與底部說明文字是否一致
6. H04 / H06 月摘要是否仍符合 Wise / A7 切換日 `2026/04/08`
