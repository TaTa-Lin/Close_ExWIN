# Close_ExWin

自動關閉 Excel OLE Timeout 等異常彈窗。

使用 **WinEvent Hook + 定期掃描** 雙重偵測模式，視窗一出現立即排程處理。

---

## 使用方式

1. 執行 `Close_ExWin.exe`
2. 右下角系統匣出現藍色圖示（X）
3. 右鍵可選：**暫停 / 設定 / 檢視 Log / 查看截圖 / 結束**

---

## 偵測機制

| 機制 | 說明 |
|------|------|
| WinEvent Hook | 監聽 `EVENT_OBJECT_SHOW` / `EVENT_SYSTEM_FOREGROUND`，視窗一出現即觸發 |
| 啟動掃描 | 程式啟動時用 `EnumWindows` 掃描所有已開啟的視窗 |
| 定期掃描 | 每 2 秒執行一次 `EnumWindows`，補抓 Hook 漏掉的特殊視窗（如 Excel OLE 等待對話框） |

> **注意**：定期掃描對同一個視窗有 10 秒冷卻，不會重複觸發。

---

## 設定說明

| 欄位 | 說明 |
|------|------|
| 記錄 Log | 動作記錄寫入 `logs/` 資料夾，每次啟動建立新檔（時間戳記命名） |
| 截圖開關 | 全域截圖開關；關閉時全不截圖，開啟時依各規則的「截圖」欄位決定是否截圖 |
| 動作延遲（秒） | 偵測到視窗後等待幾秒，再重新確認視窗仍存在且標題未變，才執行動作。預設 15 秒。設為 0 = 立即執行 |
| 視窗標題 | 要監控的視窗標題文字 |
| 比對 | `完全符合`：標題完全相同；`包含`：標題含有指定關鍵字 |
| 關鍵字 | 選填。對話框內文必須包含此字串，規則才生效（用於區分同標題的不同對話框） |
| 動作 | 按 Enter / Tab+Tab+Enter / 強制關閉 / 按結束(E) |
| 按鈕 | 選填。NUIDialog（Office DirectUI 對話框）專用，依優先順序列出按鈕名稱（逗號分隔） |
| 啟用 | 勾選才生效 |
| 截圖 | 勾選後，此規則觸發時會在動作前截圖（需全域截圖開關同時開啟） |

設定存於 `Close_ExWin_config.json`（與 EXE 同目錄）

---

## 檔案結構

```
Close_ExWin.exe
Close_ExWin_config.json
logs/
  Close_ExWin_20260604_153012.log   ← 每次啟動一個新檔
  screenshots/
    20260604_153045_Microsoft Visual Basic.png
```

---

## 動作延遲說明

當 Windows 排程（Task Scheduler）定期啟動 Excel 檔案時，Excel 啟動過程中視窗標題會短暫出現 `Microsoft Excel` 或 `活頁簿1 - Excel`，若立即動作會誤觸規則。

**動作延遲**的運作方式：
1. 偵測到符合規則的視窗
2. 等待 N 秒
3. 重新確認：視窗仍可見 **且** 標題未改變 → 執行動作
4. 若標題已變（Excel 已載入實際檔案）→ 自動跳過，不動作

建議依 Excel 檔案的開啟速度設定，通常 3～5 秒已足夠。

---

## NUIDialog（Office DirectUI 對話框）說明

Excel 的部分對話框（如 OLE 等待、儲存提示）使用 Office 自訂的 `NUIDialog` 視窗類別，無法用標準 Win32 訊息（`WM_COMMAND`）操作按鈕。

本程式改用 **UI Automation（IUIAutomation）** 讀取對話框內容並點擊按鈕：

1. 讀取對話框內文文字 → 寫入 Log
2. 依內文比對規則的「關鍵字」欄位 → 決定套用哪條規則的按鈕順序
3. 找到對應按鈕 → 呼叫 `IUIAutomationInvokePattern.Invoke()` 點擊
4. UIA 失敗時，依序嘗試 `WM_COMMAND(IDNO)` → `WM_COMMAND(IDOK)` → `WM_CLOSE` 作為 fallback

---

## 預設規則

| 視窗標題 | 比對 | 關鍵字 | 動作 | 按鈕 | 截圖 | 觸發情境 |
|---------|------|--------|------|------|------|---------|
| Microsoft Excel | 完全符合 | 正在等候 | 按 Enter | 確定 | ✘ | OLE Timeout 等待對話框（「正在等候...完成 OLE 動作」），點「確定」繼續等待，避免按取消觸發 VBA 錯誤 |
| Microsoft Excel | 完全符合 | （空） | 按 Enter | 不儲存, 取消 | ✘ | 一般儲存提示，優先點「不儲存」；唯讀檔案出現不同對話框則點「取消」 |
| 檔案使用中 | 完全符合 | （空） | 按 Tab+Tab+Enter | （空） | ✘ | 檔案被其他人開啟時的共用提示，Tab 兩次移至「唯讀」按鈕後確認 |
| Microsoft Visual Basic | 完全符合 | （空） | 按結束(E) | （空） | ✔ | VBA 執行錯誤對話框，動作前截圖保留錯誤訊息 |
| Excel | 完全符合 | （空） | 強制關閉 | （空） | ✘ | 標題剛好只有「Excel」的異常視窗 |
| 活頁簿1 - Excel | 完全符合 | （空） | 強制關閉 | （空） | ✘ | Excel 啟動後自動建立的空白活頁簿（排程啟動特定檔案時不需要） |

> **規則比對優先順序**：有填「關鍵字」的規則優先比對；同標題有多條規則時，含關鍵字且符合內文的規則優先套用。

---

## 打包 EXE

需要 Python 3.8+，執行 `build.bat` 自動打包。
產出的 EXE 不需安裝 Python，直接執行。

手動打包指令：

```
pyinstaller --onefile --windowed --name=Close_ExWin ^
  --hidden-import comtypes.client ^
  --hidden-import comtypes.gen.UIAutomationClient ^
  --collect-submodules comtypes ^
  Close_ExWIN.py
```
