# XLSX_to_XLS V4.0 Native Windows Portable

這版是重新開發的 Windows 原生架構。

## 與 V3.x 的差異

V4.0 不使用：

- Python / PyInstaller
- .NET / NPOI
- SharpZipLib
- Excel COM / ActiveX
- 產線端 BUILD.cmd / BAT
- 產線端 PowerShell
- 產線端 NuGet

V4.0 的前端是標準 Win32 C++ EXE，使用 Windows API 建立視窗。
實際 Excel 格式轉換交由 Portable 套件中一起附帶的 LibreOffice 引擎執行。

## 為什麼這個方式較適合企業電腦

正式使用電腦只會看到：

XLSX_to_XLS_V4.0.exe
LibreOffice\...

EXE 不會下載任何東西、不會呼叫編譯器、不會建立另一個 EXE、
不會啟動 PowerShell，也不會使用 PyInstaller 自解壓。

LibreOffice 本身是成熟的桌面辦公軟體。
這通常比「腳本下載 DLL 後即時編譯」或「Python one-file 自解壓」更容易被企業資安接受。

但任何沒有數位簽章的自製 EXE，都無法保證 100% 不被所有 EDR 誤判。
公司正式部署仍建議加 Authenticode 簽章或交由 IT 建立白名單。

## GitHub 產生正式 Windows EXE

將本 ZIP 的內容完整上傳 GitHub Repository。

GitHub：
Actions
→ Build V4.0 Windows Portable Release
→ Run workflow

GitHub 的 Windows Runner 會：

1. 使用 Microsoft MSVC 編譯 Win32 C++ Release EXE
2. 安裝 LibreOffice
3. 將 LibreOffice 引擎一起封裝
4. 驗證 Windows PE EXE
5. 產生：

XLSX_to_XLS_V4.0_Windows_Portable.zip

正式產線端只下載這一個 ZIP。

## 使用方式

1. 解壓縮正式 Portable ZIP
2. 保持 EXE 與 LibreOffice 資料夾在一起
3. 雙擊 XLSX_to_XLS_V4.0.exe
4. 可一次選擇多個 XLSX
5. 輸出位於：

來源資料夾\XLS_轉換完成\

## 轉換輸出

程式呼叫 LibreOffice：

--headless --convert-to "xls:MS Excel 97"

輸出為真正 Excel 97-2003 XLS，
並在轉換完成後檢查 OLE Compound File Header：

D0 CF 11 E0 A1 B1 1A E1

若不是這個簽章，程式會判定轉換失敗。

## 目前版本定位

V4.0 是 Windows x64 Portable Release。

若公司仍有真正 32-bit Windows 電腦，
需要另外製作 V4.0 Legacy 32-bit 引擎版。
