# SendEmail_A7report Project Notes

這份文件只放「下次接手時最重要的已確認規則」。

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

## athena_setting 規則

### 寄件者

- `func_name = 'email_sender'`
- 全館共用一筆
- `is_use = 'Y'`
- `value` 放 JSON

### 收件者

- `func_name = 'email_recipient'`
- 每館一筆
- `hotel_id` 區分不同飯店
- `is_use = 'Y'`
- `value` 可放純收件名單或 JSON

## 主旨規則

格式：

`{hotel_name}*本日:{今日入住}({今日住房率})*昨日:{昨日入住}({昨日住房率}) - {現在發信時間}`

注意：

- 最後時間要用發信當下時間
- 不要用資料抓取時間

## 明細區間規則

Email 下方明細要顯示：

- 昨日
- 今日起算未來 120 天

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
- 資料格式來自 `GetAthenaA7`

## 住房率顏色規則

- `0~29%`：紫色
- `30~59%`：紅色
- `60~79%`：黃色
- `80~89%`：綠色
- `90%以上`：白色

Email 最下方也要顯示說明文字：

`☆ 住房率0~29%為紫色、30~59%為紅色、60~79%為黃色、80~89%為綠色、90%以上為白色`

## 版面規則

- 內容字體已調小
- 風格以實用、可快速閱讀為主
- HTML 預覽輸出到 `bin\Debug\net5.0\output`

## 可攜版

- `dist\SendEmail_A7report_portable`
- 這包不需要 Python / Playwright
- 只要能連 MySQL 與 SMTP 即可

## 下次修改時優先確認

1. `athena_report01` 是否已有最新資料
2. `athena_setting` 的 sender / recipient 是否啟用
3. 主旨時間是否仍使用發信當下時間
4. 顏色規則與底部說明文字是否一致
