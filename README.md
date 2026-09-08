# WinCleaner v3.8 - Windows 系統安全清理工具

一款安全、有感、小白也能上手的 Windows 系統清理工具。
清理完會告訴你釋放了多少空間，讓你真的感受到差異。

## 下載使用

1. 點右上角綠色 **Code** → **Download ZIP**
2. 解壓縮
3. 雙擊 `啟動清理工具.bat`
   > 若出現藍色「Windows 已保護您的電腦」視窗，屬正常現象（非知名發布者程式都會有這提示），
   > 點「其他資訊」→「仍要執行」即可。
4. 出現 UAC 視窗點「是」（若不小心點了「否」，重新雙擊即可，不會有任何檔案被誤刪）
5. 選 `[1]` 一鍵全部清理

## 功能

- 使用者 & 系統暫存檔
- Windows Update 快取
- 瀏覽器快取（Chrome / Edge / Firefox / Brave / IE）
- **App 快取**（LINE / Discord / Teams / Slack / Steam / Zoom）
- Delivery Optimization 快取（常常 5~10GB）
- 預先擷取 & 縮圖快取
- DNS 快取、字型快取重建
- 錯誤報告 & 系統日誌
- 記憶體最佳化
- **啟動項目管理**（讓開機變快的關鍵）
- 舊版 Windows 安裝（Windows.old）

## 支援系統

Windows 7 / 8 / 8.1 / 10 / 11（所有版本）

## 安全說明

- 不會刪除個人檔案（文件、圖片、影片）
- 不會刪除 LINE 對話紀錄、Discord 設定、瀏覽器書籤密碼
- 只清快取垃圾，所有操作有錯誤保護
- Windows.old 刪除需手動確認
- 啟動項目停用可隨時從工作管理員還原

## 更新日誌

### v3.8（2026-09-08，暫時診斷版）
- v3.7 也沒解決中文亂碼（症狀跟 v3.6 完全相同），代表問題不在 chcp/編碼設定這一層。加了一段暫時的純 ASCII 診斷輸出，印出主控台實際的 codepage 與編碼狀態以找出真正原因；確認後會移除這段診斷代碼

### v3.7（2026-09-08，放棄 Win7 強制 UTF-8）
- 修正：v3.6 把 `chcp` 跟 `Console.OutputEncoding` 都設成 UTF-8 後，中文在原生 Win7 上還是亂碼。查證後確認這是 Windows 7 舊版主控台（點陣字型 Raster Fonts）對 codepage 65001 支援不完整的已知限制（見 [microsoft/terminal#16701](https://github.com/microsoft/terminal/issues/16701)），不是編碼沒對齊。改為只在 PowerShell 3+／Windows 8 以後才切換 UTF-8；PS2／原生 Win7 維持系統原生 codepage（例如繁體中文的 950），本來就能正確顯示中文

### v3.6（2026-09-08，中文亂碼修正）
- 修正：v3.5 解決當機問題後，實測發現中文整個顯示亂碼。只設定 `[Console]::OutputEncoding = UTF8` 並不會改變主控台本身實際使用的 codepage，兩邊編碼對不上導致誤解讀；加上 `chcp 65001` 讓兩邊一致

### v3.5（2026-09-08，第三次實機測試後修正）
- 修正：v3.4 只處理了套色失敗的情況，但實測發現部分遠端桌面環境連無色的 `Write-Host` 都會失敗；徹底繞過 `Write-Host`，失敗時改用 `[Console]::WriteLine()`/`[Console]::Write()` 直接寫入輸出資料流
- 修正：全檔案 12 處裸的 `Write-Host ""` 空白行輸出沒有套用保護，一併納入相同容錯機制

### v3.4（2026-09-08，實機測試後再修正）
- 修正：透過遠端桌面工具（AnyDesk 等）或部分虛擬機主控台執行時，`Write-Host -ForegroundColor` 讀取主控台緩衝區失敗，跳出大量 `device...not functioning` 錯誤洗版；`Write-C`/elevation 階段的彩色輸出現在讀取失敗時自動退回無色輸出

### v3.3（2026-09-08，實機 Windows 7 測試後修正）
- 修正：主選單用了 `-in` 運算子（PowerShell 3.0+ 才有語法），原生 Win7 的 PowerShell 2.0 連整支腳本都無法解析，一啟動就報 parse error 閃退
- 修正：啟動項目管理用 `[PSCustomObject]@{}` 轉換 hashtable 是 PS3+ 語意，在 PowerShell 2.0 上不會產生正確的具名屬性，改用 `New-Object PSObject -Property @{}`

### v3.2（2026-09-04）
- 修正：原生 Windows 7（PowerShell 2.0）讀不到系統資訊、記憶體最佳化完全失效（`Get-CimInstance` 需要 PS3+，改為自動 fallback 回 `Get-WmiObject`）
- 修正：UAC 視窗點「否」時會顯示未處理例外後閃退，改為清楚提示並可安全重試
- 修正：啟動器 `.bat` 誤用 Unix 語法 `/dev/null`，改回 Windows 的 `nul`
- 新增：使用說明補上 SmartScreen 警示、UAC 誤按「否」的 FAQ

### v3.1（2026-03-20）
- 修正：bat 檔在繁體中文 Windows 出現亂碼錯誤（移除 bat 內中文，改為純英文）
- 修正：PS1 加入 UTF-8 BOM，解決 PowerShell 在中文系統語法錯誤
- 修正：中文介面文字重疊顯示（將 `chcp 65001` 改為 `[Console]::OutputEncoding`）
- 修正：Delivery Optimization 清理指令名稱錯誤（`Delete-` → `Clear-`）
- 改善：清縮圖快取前加提示，避免使用者誤以為桌面當機
- 改善：RAM 最佳化加注意說明（HDD 電腦可能短暫變慢）
- 改善：`Get-WmiObject` 全面改為 `Get-CimInstance`，相容 PowerShell 7+

### v3.0
- 初始發布
- 支援 15 項清理功能
- 雙語介面（中/英）
- 啟動項目管理
