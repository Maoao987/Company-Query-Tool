# Company Query Tool

台灣公司資料與股票資訊查詢工具。支援以統一編號、公司名稱、股票代號查詢公司登記資料，並整合上市、上櫃、興櫃、ETF 與基金常用資訊。

## 功能

- 單筆查詢：統一編號、公司名稱、股票代號。
- 批次查詢：上傳 CSV / Excel，一次查多筆資料。
- 公司資料：登記現況、公司名稱、資本額、所在地、董監事、營業項目。
- 股市資訊：市場別、商品類型、發行地、ISIN Code、指定日期前最近收盤價。
- 除權息資訊：整理近兩年除權息紀錄與股利資料。
- 匯出格式：Excel、CSV、PDF。
- 來源快照：可產生 findbiz 與股市來源頁面的 PDF 快照。
- 自動更新：透過 GitHub Releases 檢查新版並下載安裝。

## 使用方式

### Windows 安裝版

從 GitHub Releases 下載最新版 `CompanyQueryToolSetup.exe`，安裝後開啟桌面捷徑即可使用。

### 原始碼執行

需要 Python 3.12 以上。

```powershell
pip install -r requirements.txt
streamlit run app.py
```

## 批次查詢格式

CSV 或 Excel 可包含以下任一欄位：

- `統一編號`
- `股票代號`
- `公司名稱`

系統會依序優先使用統一編號、股票代號、公司名稱。如果沒有明確欄名，會嘗試從第一欄自動判斷。

## 資料來源

- 經濟部商工登記公示資料查詢服務 findbiz
- 臺灣證券交易所 TWSE
- 證券櫃檯買賣中心 TPEX
- ISIN 公開資料
- Yahoo Finance / yfinance
- 公開資訊觀測站 MOPS

## 發布流程

版本號以 `version.txt` 為準。一般使用者可見更新請升 patch 版本：

```powershell
powershell -ExecutionPolicy Bypass -File .\publish_release.ps1 -Mode release
```

同版 hotfix 可覆蓋既有 release asset：

```powershell
powershell -ExecutionPolicy Bypass -File .\publish_release.ps1 -Mode fix
```

## 注意事項

- 外部資料來源若改版、限流或暫時不可用，查詢結果可能失敗，稍後重試通常可恢復。
- 興櫃歷史行情使用官方提供的成交均價，與上市櫃收盤價定義不同。
- ETF、基金與部分特殊商品可能無法唯一對應公司統編，但仍會盡量提供市場、ISIN、發行地與股利資料。
