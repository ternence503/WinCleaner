# WinCleaner v3.1 - Windows 系統安全清理工具

一款安全、有感、小白也能上手的 Windows 系統清理工具。
清理完會告訴你釋放了多少空間，讓你真的感受到差異。

## 下載使用

1. 點右上角綠色 **Code** → **Download ZIP**
2. 解壓縮
3. 雙擊 `啟動清理工具.bat`
4. 出現 UAC 視窗點「是」
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
