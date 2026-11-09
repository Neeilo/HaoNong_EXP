# HaoNong_EXP

## 訂單彙總／撿貨清單工具（PWA 版）

`order_tool.exe`（Python／tkinter）移植成 **PWA**：手機、平板、電腦的瀏覽器都能用，可「安裝」到主畫面，**完全離線可用**。
所有處理都在瀏覽器本機完成，**客戶的姓名、電話、地址不會上傳到任何伺服器**。

### 用法
1. 開啟網頁（或從主畫面開啟 App）。
2. 選擇一個或多個 CRM 匯出檔（`.xlsx` / `.xlsm`），也可以直接拖放。
3. 按「產生收件人清單＋撿貨清單」。
4. 沒問題 → 預覽結果並下載 `收件人清單_撿貨清單_<時間>.xlsx`（含 `收件人清單`／`撿貨清單`／`處理紀錄` 三張工作表，手機可直接「分享／傳送」）。
   有問題 → 畫面列出所有問題，並可下載 `error_report_<時間>.txt`；與原工具一樣，**有任何問題就不產生輸出檔**。

### 與原 exe 的行為差異
行為以原程式為準逐項移植，並用原程式實際跑出來的結果驗證（見「測試」）。刻意的差異只有：
- 不支援舊版 `.xls`（原工具用 openpyxl，本來就讀不了 `.xls`；現在會明確提示「請另存為 .xlsx」）。
- 數量／單價若是 `nan`、`inf` 這類文字，視為「無法辨識」（Python 的 `float()` 會當成數字）。
- 檔案損毀或無法開啟時，以錯誤報告的一項呈現，而不是整個程式當掉。

### 部署
`docs/` 是純靜態網站（沒有建置步驟），任何靜態空間都能放；路徑皆為相對路徑，可放在子目錄。
GitHub Pages：Settings → Pages → Build and deployment → Deploy from a branch → 選 `main`、資料夾 `/docs`。
（PWA 需要 HTTPS，GitHub Pages 預設就是。）

本機預覽：

```bash
npm run serve   # http://localhost:8080/
```

### 開發
| 檔案 | 說明 |
|---|---|
| `docs/core.js` | 核心邏輯（欄位對應、訂單彙總、撿貨清單、錯誤報告），與畫面和 Excel 函式庫無關 |
| `docs/xlsx-io.js` | 以 ExcelJS 讀取輸入檔、產生輸出 `.xlsx` |
| `docs/app.js`, `index.html`, `styles.css` | 畫面 |
| `docs/sw.js`, `manifest.webmanifest`, `icons/` | PWA（離線快取、安裝） |
| `docs/vendor/` | ExcelJS 4.4.0（MIT），已內含，不需 npm install |

**發佈新版**：修改 `docs/` 內任何檔案後，把 `docs/sw.js` 的 `VERSION` 與 `docs/app.js` 的 `APP_VERSION` 一起加一；
使用者的畫面會出現「立即更新」。若新增或刪除檔案，也要同步更新 `sw.js` 的 `ASSETS`（測試會檢查）。

### 測試
```bash
npm test
```
- `tests/fixtures/golden.json`：用**原 order_tool.exe 的 Python 邏輯**，對 52 組隨機生成的 CRM 檔案（含全形數字、千分位、重複表頭、缺欄位、各種異常資料）實際跑出來的輸入與輸出；JS 版必須逐項一致。
- 另有 PWA 檔案完整性測試（`sw.js` 預先快取清單、版本號、manifest 圖示）。
