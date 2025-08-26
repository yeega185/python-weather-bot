# Weather Summary Auto-Split (程式自動輸出)

## 目的
你的程式會自動產生 `weather_summary.xlsx` 並寫入 OneDrive；當該檔被覆蓋時，Power Automate 會自動備份並把資料依縣市拆成多個工作表，且在每列最左側加入當日日期。

## 架構
1. **資料產生（程式）**：內部/排程程式將資料整理成 `weather_summary.xlsx`，存放至 OneDrive `/weather`（程式負責格式與欄位）。  
2. **自動化處理（Power Automate）**：Trigger = When a file is created or modified (properties only)  
   - Delay 15s → Get file metadata & content → 判斷是否為覆蓋（Created ≠ Modified）  
   - 若為覆蓋：Create file（備份到 `/backup`）→ Excel Online → Run Office Script (`SplitByCityAddDate`)
3. **Office Script**：`SplitByCityAddDate` 會偵測縣/市欄，依縣市分組並在每列前加入當日日期（yyyy/MM/dd）。

## 部署條件
- OneDrive for Business（支援 Office Scripts）  
- Power Automate 權限（OneDrive、Excel Online）  
- 程式能把 `weather_summary.xlsx` 寫到 OneDrive 指定資料夾（或使用 Microsoft Graph 上傳）

## 測試
1. 由程式產生並上傳 `weather_summary.xlsx` 到 `/weather`（或手動複製測試檔）。  
2. 覆蓋該檔 → 等 30 秒 → 檢查 Power Automate Run history、/backup 與 Excel 裡的各縣市分頁。

