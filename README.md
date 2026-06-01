# Close_ExWin

自動關閉 Excel OLE Timeout 等異常彈窗。

使用 **WinEvent Hook** 模式：視窗一出現立即處理，不輪詢、不閃爍。

---

## 使用方式

1. 執行 `Close_ExWin.exe`
2. 右下角系統匣出現藍色圖示（X）
3. 右鍵 → **設定** 或 **結束**

---

## 設定說明

| 欄位 | 說明 |
|------|------|
| 視窗標題 | 要監控的視窗標題文字 |
| 比對 | `exact`=完全符合、`contains`=包含關鍵字 |
| 動作 | 按 Enter / Tab+Tab+Enter / 強制關閉 |
| 啟用 | 勾選才生效 |

設定存於 `Close_ExWin_config.json`（與 EXE 同目錄）
日誌存於 `Close_ExWin.log`（與 EXE 同目錄）

---

## 預設規則

| 視窗標題 | 動作 |
|---------|------|
| Microsoft Excel | 按 Enter |
| 檔案使用中 | 按 Tab+Tab+Enter |
| Microsoft Visual Basic | 按 Enter |
| Excel | 強制關閉 |
| 活頁簿1 - Excel | 強制關閉 |

---

## 打包 EXE

需要 Python 3.8+，執行 `build.bat` 自動打包。
產出的 EXE 不需安裝 Python，直接執行。
