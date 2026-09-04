# XLSX_to_XLS V6.0 產線簡易穩定版

V6.0 不再把目標設定成「完整 Office 文件格式轉換器」，而是針對產線資料表使用情境：
讀取 XLSX 的工作表與儲存格，重新建立真正 Excel 97-2003 BIFF8 XLS。

## 主要功能

- 一次選擇一個或多個 XLSX
- 多工作表
- 文字、數字、日期、布林值
- 一般 Excel 公式（可被 xlwt BIFF8 parser 接受者）
- 欄寬、列高
- 合併儲存格
- 常用字型 / 粗體 / 斜體 / 對齊 / 框線 / 部分底色
- 同名輸出自動加流水號
- 自動建立 `XLS_轉換完成`
- 輸出後驗證 OLE/BIFF8 Header
- 強制 BIFF8 65536 列 / 256 欄限制
- Style cache + style_compression，避免過去 4094 XF 樣式爆量問題

## 不保證完整保留

V6.0 是「產線資料表優先」，以下項目不屬於主要目標：

- VBA / 巨集
- ActiveX
- Power Query
- Pivot Cache / 樞紐分析快取
- 圖片與圖表
- 外部資料連線
- 複雜條件格式
- 部分新式 Excel 公式

## 為什麼這版不是單一 EXE

V5.0 把大型執行引擎全部塞進單一 EXE，在企業防毒環境中容易觸發啟發式偵測。
V6.0 改成 PyInstaller `--onedir` Portable 資料夾，避免自解壓式單檔架構。

正式電腦只需要把整個：
`XLSX_to_XLS_V6.0_Portable`
資料夾複製過去，再執行：
`XLSX_to_XLS_V6.0.exe`

不需要安裝 Python、Excel 或 LibreOffice。

## GitHub 建置

把本專案加入既有 Repository 後：
Actions → Build V6.0 Production Stable Portable → Run workflow

下載 Artifact：
`XLSX_to_XLS_V6.0_Windows_Portable`
