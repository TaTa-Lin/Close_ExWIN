# Close_ExWin

自動關閉 Excel OLE Timeout 等異常彈窗。

使用 **WinEvent Hook + 定期掃描** 雙重偵測模式，視窗一出現立即排程處理。

---

## 使用方式

1. 執行 `Close_ExWin.exe`
2. 右下角系統匣出現藍色圖示（X）
3. 右鍵 → **設定** 或 **結束**

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
| 記錄 Log | 動作記錄寫入 `Close_ExWin.log` |
| 動作延遲（秒） | 偵測到視窗後等待幾秒，再重新確認視窗仍存在且標題未變，才執行動作。預設 3 秒。設為 0 = 立即執行 |
| 視窗標題 | 要監控的視窗標題文字 |
| 比對 | `完全符合`：標題完全相同；`包含`：標題含有指定關鍵字 |
| 動作 | 按 Enter / Tab+Tab+Enter / 強制關閉 |
| 啟用 | 勾選才生效 |

設定存於 `Close_ExWin_config.json`（與 EXE 同目錄）
日誌存於 `Close_ExWin.log`（與 EXE 同目錄）

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

## 預設規則

| 視窗標題 | 比對 | 動作 | 觸發情境 |
|---------|------|------|---------|
| Microsoft Excel | 完全符合 | 按 Enter | OLE Timeout 等待對話框（「正在等候...完成 OLE 動作」） |
| 檔案使用中 | 完全符合 | 按 Tab+Tab+Enter | 檔案被其他人開啟時的共用提示，Tab 兩次移至「唯讀」按鈕後確認 |
| Microsoft Visual Basic | 完全符合 | 按 Enter | VBA 執行錯誤對話框 |
| Excel | 完全符合 | 強制關閉 | 標題剛好只有「Excel」的異常視窗 |
| 活頁簿1 - Excel | 完全符合 | 強制關閉 | Excel 啟動後自動建立的空白活頁簿（排程啟動特定檔案時不需要） |

---

## 打包 EXE

需要 Python 3.8+，執行 `build.bat` 自動打包。
產出的 EXE 不需安裝 Python，直接執行。
