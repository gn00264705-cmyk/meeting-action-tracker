# 會議行事曆與改善事項追蹤系統

本專案以 Google Apps Script（GAS）與 Google Sheets 建置私人使用的會議行事曆 Web 應用程式。前端採 Vue 3 Composition API、Tailwind CSS 與 Lucide Icons CDN；首次開啟時會自動建立 `Meetings` 與 `ActionItems` 分頁及其欄位。

## 部署步驟

1. 在 Google 雲端硬碟新增 Google 試算表，作為本系統資料庫。
2. 在該試算表選擇「擴充功能」→「Apps Script」，建立綁定式 GAS 專案。
3. 將 [Code.gs](Code.gs) 與 [Index.html](Index.html) 內容分別貼入同名檔案，並以 [appsscript.json](appsscript.json) 覆蓋專案資訊清單；儲存專案。
4. 若使用獨立 GAS 專案而非試算表綁定專案，請在「專案設定」→「指令碼屬性」新增 `SPREADSHEET_ID`，值為資料庫試算表網址中的 ID。
5. 選擇「部署」→「新增部署作業」→類型選「網頁應用程式」。執行身分建議選「我」，存取權選擇僅自己可存取；完成 Google 授權後開啟部署網址。
6. 首次開啟會建立資料表。首次使用「同步至 Google 日曆」時，依畫面要求授權 Calendar 權限。

## 寫入管線自我測試

部署前或發生無法儲存的問題時，在 GAS 編輯器上方的函式選單選擇 `run_verification_test` 後按「執行」。此完整驗收會測試特殊符號標題、空標題拒絕、HH:mm 時間格式、0／1／多筆改善事項關聯，以及前端送出合約；所有 Mock 資料會自動移除。請在「執行記錄」確認全部 `[PASS]`。若出現 `[FAIL]`，可直接依錯誤訊息檢查試算表綁定、`SPREADSHEET_ID`、帳號授權或前端通訊實作。

`test_backendPipeline` 仍保留作為快速單筆寫入／查詢測試。

## 資料與權限注意事項

- 此版本設計為私人部署；資料存於部署帳號可存取的 Google 試算表與主日曆。
- 請勿手動變更兩個分頁的第一列欄位名稱或順序。系統偵測到有資料的欄位不符時會停止，避免破壞既有資料。
- 重複同步同一會議時，系統會嘗試更新先前建立的日曆事件；事件對應 ID 儲存在 GAS 指令碼屬性中。
- 獨立 GAS 專案可直接在 `Code.gs` 頂端填入 `SPREADSHEET_ID`；綁定式專案則優先使用其目前的試算表。
