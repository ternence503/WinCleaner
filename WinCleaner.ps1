#Requires -Version 2.0
<#
.SYNOPSIS
    WinCleaner - Windows 系統安全清理工具 / Windows System Cleaner
.DESCRIPTION
    安全、雙語 (中/英) 系統清理工具，支援 Windows 7 ~ 11
    Safe bilingual (Chinese/English) system cleaner for Windows 7-11
.VERSION
    3.1
.NOTES
    以系統管理員身份執行 / Run as Administrator
#>

# ============================================================
# SECTION 1: AUTO-ELEVATION / 自動取得管理員權限
# ============================================================

function Test-IsAdmin {
    $identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdmin)) {
    Write-Host "需要管理員權限，正在請求... / Requesting administrator rights..." -ForegroundColor Yellow
    $scriptPath = $MyInvocation.MyCommand.Path
    $psArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    Start-Process powershell -ArgumentList $psArgs -Verb RunAs
    exit
}

try {
    $Host.UI.RawUI.WindowTitle = "WinCleaner v3.1 - Windows 系統清理工具"
    $Host.UI.RawUI.BufferSize  = New-Object System.Management.Automation.Host.Size(120, 3000)
    $Host.UI.RawUI.WindowSize  = New-Object System.Management.Automation.Host.Size(82, 42)
} catch { }

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# ============================================================
# SECTION 2: GLOBAL STATE / 全域變數
# ============================================================

$script:TotalFreed  = [long]0
$script:TaskResults = [System.Collections.ArrayList]@()
$script:WinVersion  = 0
$script:PSMajor     = $PSVersionTable.PSVersion.Major

# ============================================================
# SECTION 3: UTILITY FUNCTIONS / 工具函式
# ============================================================

function Get-WindowsVersion {
    $osVer = [System.Environment]::OSVersion.Version
    switch ($osVer.Major) {
        6 {
            switch ($osVer.Minor) {
                1 { return 7 }
                2 { return 8 }
                3 { return 8 }
                default { return 8 }
            }
        }
        10 {
            if ($osVer.Build -ge 22000) { return 11 }
            return 10
        }
        default { return 10 }
    }
}

function Get-FolderSize {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return [long]0 }
    $size = [long]0
    try {
        $files = [System.IO.Directory]::GetFiles($Path, "*", [System.IO.SearchOption]::AllDirectories)
        foreach ($f in $files) {
            try { $size += (New-Object System.IO.FileInfo($f)).Length } catch { }
        }
    } catch { }
    return $size
}

function Format-Bytes {
    param([long]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N1} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N0} KB" -f ($Bytes / 1KB) }
    if ($Bytes -gt 0)   { return "$Bytes B" }
    return "0 B"
}

function Write-C {
    param([string]$Text, [ConsoleColor]$Color = [ConsoleColor]::White, [switch]$NoNewLine)
    if ($NoNewLine) { Write-Host $Text -ForegroundColor $Color -NoNewline }
    else            { Write-Host $Text -ForegroundColor $Color }
}

function Record-Result {
    param([string]$NameEN, [string]$NameCN, [long]$Freed)
    $null = $script:TaskResults.Add(@{EN=$NameEN; CN=$NameCN; Freed=$Freed})
}

function Show-TaskStart {
    param([string]$NameEN, [string]$NameCN)
    Write-C "  [ ] " -Color DarkGray -NoNewLine
    Write-C $NameCN -Color Cyan -NoNewLine
    Write-C " / $NameEN ..." -Color DarkGray -NoNewLine
}

function Show-TaskDone {
    param([long]$Freed, [string]$Label = "")
    Write-Host ""
    if ($Label -ne "") {
        Write-C "  [v] $Label" -Color Green
    } elseif ($Freed -gt 0) {
        Write-C "  [v] " -Color Green -NoNewLine
        Write-C "釋放 " -Color White -NoNewLine
        Write-C (Format-Bytes $Freed) -Color Yellow
    } else {
        Write-C "  [v] 已完成 / Done" -Color DarkGreen
    }
}

function Show-TaskSkip {
    param([string]$Reason)
    Write-Host ""
    Write-C "  [-] 略過 / Skipped: $Reason" -Color DarkGray
}

function Remove-Contents {
    # 安全刪除資料夾內容，回傳已釋放位元組
    param([string]$Path, [string]$Filter = "*")
    if (-not (Test-Path $Path)) { return [long]0 }
    $freed = [long]0
    try {
        $items = Get-ChildItem -LiteralPath $Path -Filter $Filter -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            try {
                $size = if ($item.PSIsContainer) { Get-FolderSize $item.FullName } else { $item.Length }
                Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
                $freed += $size
            } catch { }
        }
    } catch { }
    return $freed
}

# ============================================================
# SECTION 4: UI / 使用者介面
# ============================================================

function Show-Header {
    Clear-Host
    Write-C ""
    Write-C "  +========================================================+" -Color Cyan
    Write-C "  |                                                        |" -Color Cyan
    Write-C "  |        Windows 系統安全清理工具  v3.1                 |" -Color Yellow
    Write-C "  |        WinCleaner - Windows Security Cleaner v3.1     |" -Color White
    Write-C "  |                                                        |" -Color Cyan
    Write-C "  +========================================================+" -Color Cyan
    Write-C ""
    $verLabel = switch ($script:WinVersion) {
        7 {"Windows 7"}; 8 {"Windows 8/8.1"}; 10 {"Windows 10"}; 11 {"Windows 11"}
        default {"Windows"}
    }
    Write-C "  系統 / OS: $verLabel  |  PowerShell $($PSVersionTable.PSVersion)" -Color DarkGray
    Write-C ""
}

function Show-MainMenu {
    Write-C "  +---------------------------------------------+" -Color Green
    Write-C "  |           請選擇功能 / Choose Function      |" -Color Green
    Write-C "  +---------------------------------------------+" -Color Green
    Write-C "  |                                             |" -Color Green
    Write-C "  |  [1]  一鍵全部清理 Full Clean  (推薦)      |" -Color White
    Write-C "  |                                             |" -Color Green
    Write-C "  |  [2]  自選清理項目 Selective Clean          |" -Color White
    Write-C "  |                                             |" -Color Green
    Write-C "  |  [3]  啟動項目管理 Startup Manager  (加速) |" -Color White
    Write-C "  |                                             |" -Color Green
    Write-C "  |  [4]  系統資訊 System Info                  |" -Color White
    Write-C "  |                                             |" -Color Green
    Write-C "  |  [Q]  離開 Exit                             |" -Color White
    Write-C "  |                                             |" -Color Green
    Write-C "  +---------------------------------------------+" -Color Green
    Write-C ""
    Write-C "  請輸入選項 / Enter choice [1/2/3/4/Q]: " -Color Yellow -NoNewLine
}

function Show-SelectiveMenu {
    $items = @(
        @{N=1;  CN="使用者暫存檔";          EN="User Temp Files"},
        @{N=2;  CN="系統暫存檔";            EN="System Temp Files"},
        @{N=3;  CN="Windows 更新快取";      EN="Windows Update Cache"},
        @{N=4;  CN="資源回收筒";            EN="Recycle Bin"},
        @{N=5;  CN="瀏覽器快取";            EN="Browser Cache"},
        @{N=6;  CN="App 快取 (LINE/Discord等)"; EN="App Cache (LINE/Discord etc.)"},
        @{N=7;  CN="Delivery Optimization"; EN="Delivery Optimization Cache"},
        @{N=8;  CN="預先擷取檔案";          EN="Prefetch Files"},
        @{N=9;  CN="縮圖快取";              EN="Thumbnail Cache"},
        @{N=10; CN="DNS 快取";              EN="DNS Cache"},
        @{N=11; CN="字型快取重建";          EN="Font Cache Rebuild"},
        @{N=12; CN="錯誤報告檔案";          EN="Windows Error Reports"},
        @{N=13; CN="舊版 Windows (Win.old)"; EN="Old Windows Install"},
        @{N=14; CN="系統日誌";              EN="System Log Files"},
        @{N=15; CN="記憶體最佳化";          EN="RAM Optimization"}
    )

    Write-C "  +-----------------------------------------------------+" -Color DarkCyan
    Write-C "  |  選擇要清理的項目 / Select items to clean           |" -Color DarkCyan
    Write-C "  +-----------------------------------------------------+" -Color DarkCyan
    foreach ($item in $items) {
        $numStr = "[$($item.N)]".PadLeft(4)
        Write-C "  $numStr  $($item.CN)" -Color White -NoNewLine
        Write-C "  /  $($item.EN)" -Color DarkGray
    }
    Write-C "  +-----------------------------------------------------+" -Color DarkCyan
    Write-C ""
    Write-C "  輸入編號 (逗號分隔，如: 1,3,5) 或 [A] 全選" -Color Yellow
    Write-C "  Enter numbers (e.g. 1,3,5) or [A] for All: " -Color DarkGray -NoNewLine

    $inputStr = Read-Host
    if ($inputStr.Trim().ToUpper() -eq "A") { return @(1..15) }
    $selected = @()
    foreach ($part in $inputStr.Split(",")) {
        $n = 0
        if ([int]::TryParse($part.Trim(), [ref]$n) -and $n -ge 1 -and $n -le 15) {
            $selected += $n
        }
    }
    return $selected
}

function Show-SystemInfo {
    Show-Header
    Write-C "  系統資訊 / System Information" -Color Cyan
    Write-C "  ─────────────────────────────────────────────────────" -Color DarkGray
    Write-C ""
    try {
        $os = Get-WmiObject -Class Win32_OperatingSystem -ErrorAction Stop
        Write-C "  作業系統: $($os.Caption) (Build $($os.BuildNumber))" -Color White
        $totalRAM = [math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
        $freeRAM  = [math]::Round($os.FreePhysicalMemory / 1MB, 1)
        $usedRAM  = [math]::Round($totalRAM - $freeRAM, 1)
        Write-C "  記憶體:   使用 ${usedRAM}GB / 總計 ${totalRAM}GB  (可用 ${freeRAM}GB)" -Color White
    } catch { Write-C "  系統資訊讀取失敗" -Color Red }
    Write-C ""
    Write-C "  磁碟空間 / Disk Space:" -Color Cyan
    try {
        $drives = [System.IO.DriveInfo]::GetDrives() | Where-Object { $_.DriveType -eq 'Fixed' }
        foreach ($d in $drives) {
            try {
                $total = [math]::Round($d.TotalSize / 1GB, 1)
                $free  = [math]::Round($d.AvailableFreeSpace / 1GB, 1)
                $used  = [math]::Round($total - $free, 1)
                $pct   = [math]::Round(($used / $total) * 100, 0)
                $bar   = "#" * [math]::Round($pct / 5)
                $empty = "-" * (20 - $bar.Length)
                Write-C "    $($d.Name)  [${bar}${empty}] $pct%  已用 ${used}GB / ${total}GB  (剩 ${free}GB)" -Color White
            } catch { }
        }
    } catch { }
    Write-C ""
    Write-C "  按任意鍵返回 / Press any key..." -Color DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ============================================================
# SECTION 5: CLEANING FUNCTIONS / 清理函式
# ============================================================

function Clear-UserTemp {
    Show-TaskStart "User Temp Files" "使用者暫存檔"
    $freed = [long]0
    foreach ($p in @($env:TEMP, $env:TMP) | Sort-Object -Unique) {
        if ($p -and (Test-Path $p)) { $freed += Remove-Contents $p }
    }
    Record-Result "User Temp" "使用者暫存檔" $freed
    $script:TotalFreed += $freed
    Show-TaskDone $freed
}

function Clear-SystemTemp {
    Show-TaskStart "System Temp Files" "系統暫存檔"
    $freed = Remove-Contents "$env:windir\Temp"
    Record-Result "System Temp" "系統暫存檔" $freed
    $script:TotalFreed += $freed
    Show-TaskDone $freed
}

function Clear-WindowsUpdateCache {
    Show-TaskStart "Windows Update Cache" "Windows 更新快取"
    $freed = [long]0
    try {
        Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
        Stop-Service -Name bits     -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        $freed = Remove-Contents "$env:windir\SoftwareDistribution\Download"
        Start-Service -Name wuauserv -ErrorAction SilentlyContinue
        Start-Service -Name bits     -ErrorAction SilentlyContinue
    } catch { }
    Record-Result "WU Cache" "Windows 更新快取" $freed
    $script:TotalFreed += $freed
    Show-TaskDone $freed
}

function Clear-RecycleBinSafe {
    Show-TaskStart "Recycle Bin" "資源回收筒"
    $freed = [long]0
    if ($script:PSMajor -ge 5) {
        try {
            $shell = New-Object -ComObject Shell.Application
            $bin   = $shell.Namespace(0xA)
            foreach ($item in $bin.Items()) { try { $freed += $item.Size } catch { } }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null
            Clear-RecycleBin -Force -ErrorAction Stop
        } catch { $freed = [long]0 }
    } else {
        $drives = [System.IO.DriveInfo]::GetDrives() | Where-Object { $_.DriveType -eq 'Fixed' }
        foreach ($drive in $drives) {
            $binPath = "$($drive.Name)`$Recycle.Bin"
            if (Test-Path $binPath) {
                $before = Get-FolderSize $binPath
                & cmd.exe /c "rd /s /q `"$binPath`"" 2>$null
                $freed += $before
            }
        }
    }
    Record-Result "Recycle Bin" "資源回收筒" $freed
    $script:TotalFreed += $freed
    Show-TaskDone $freed
}

function Clear-BrowserCache {
    Show-TaskStart "Browser Cache" "瀏覽器快取"
    $freed = [long]0

    # 安全快取目錄（只清快取，不碰書籤/密碼/設定）
    # Safe cache dirs only - never touches bookmarks, passwords, settings
    $chromeCacheDirs = @("Cache","Cache2","Code Cache","GPUCache","Media Cache","ShaderCache","Service Worker\CacheStorage")

    # Chrome
    $chromeBase = "$env:LOCALAPPDATA\Google\Chrome\User Data"
    if (Test-Path $chromeBase) {
        $profiles = @("Default") + @(Get-ChildItem $chromeBase -Directory -Filter "Profile *" -EA SilentlyContinue | Select-Object -ExpandProperty Name)
        foreach ($prof in $profiles) {
            foreach ($cd in $chromeCacheDirs) { $freed += Remove-Contents "$chromeBase\$prof\$cd" }
        }
        Write-Host ""; Write-C "    Chrome: $(Format-Bytes $freed)" -Color DarkGreen
    }

    # Edge (Chromium)
    $edgeBase = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
    $ef = [long]0
    if (Test-Path $edgeBase) {
        $profiles = @("Default") + @(Get-ChildItem $edgeBase -Directory -Filter "Profile *" -EA SilentlyContinue | Select-Object -ExpandProperty Name)
        foreach ($prof in $profiles) {
            foreach ($cd in $chromeCacheDirs) { $ef += Remove-Contents "$edgeBase\$prof\$cd" }
        }
        Write-C "    Edge: $(Format-Bytes $ef)" -Color DarkGreen
        $freed += $ef
    }

    # Legacy Edge
    $freed += Remove-Contents "$env:LOCALAPPDATA\Packages\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\AC\MicrosoftEdge\Cache"

    # Firefox
    $ffBase = "$env:APPDATA\Mozilla\Firefox\Profiles"
    $ff = [long]0
    if (Test-Path $ffBase) {
        Get-ChildItem $ffBase -Directory -EA SilentlyContinue | ForEach-Object {
            $ff += Remove-Contents "$($_.FullName)\cache2"
            $ff += Remove-Contents "$($_.FullName)\OfflineCache"
            $ff += Remove-Contents "$($_.FullName)\thumbnails"
            $ff += Remove-Contents "$($_.FullName)\startupCache"
        }
        Write-C "    Firefox: $(Format-Bytes $ff)" -Color DarkGreen
        $freed += $ff
    }

    # Brave
    $braveBase = "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data"
    if (Test-Path $braveBase) {
        $profiles = @("Default") + @(Get-ChildItem $braveBase -Directory -Filter "Profile *" -EA SilentlyContinue | Select-Object -ExpandProperty Name)
        foreach ($prof in $profiles) {
            foreach ($cd in @("Cache","Code Cache","GPUCache")) { $freed += Remove-Contents "$braveBase\$prof\$cd" }
        }
    }

    # IE
    $freed += Remove-Contents "$env:LOCALAPPDATA\Microsoft\Windows\INetCache"
    $freed += Remove-Contents "$env:USERPROFILE\AppData\Local\Microsoft\Windows\Temporary Internet Files"

    Record-Result "Browser Cache" "瀏覽器快取" $freed
    $script:TotalFreed += $freed
    Write-C "  [v] " -Color Green -NoNewLine
    Write-C "瀏覽器快取合計: " -Color White -NoNewLine
    Write-C (Format-Bytes $freed) -Color Yellow
}

function Clear-AppCache {
    <#
    ★ 安全邊界說明 / Safety boundaries:
    LINE    - 只清 Cache 子資料夾，絕不碰對話資料庫 (.db) 和帳號資料
    Discord - 只清 Cache/GPUCache/Code Cache，不碰 Local Storage (帳號/設定)
    Teams   - 同上，不碰 databases
    Slack   - 同上
    Steam   - 只清 httpcache (網路圖片)，不碰遊戲檔案
    Zoom    - 只清 logs 資料夾
    Telegram - 不清理（資料庫與快取難以區分，跳過保安全）
    WeChat   - 不清理（同上）
    #>

    Show-TaskStart "App Cache" "App 快取 (LINE/Discord/Teams 等)"
    $freed = [long]0

    # ── LINE ──────────────────────────────────────────────────
    # 只清 Cache 子資料夾；Data 根目錄含對話資料庫，絕對不碰
    # Only clears Cache subfolder; Data root contains message DB, never touched
    $lineCache = "$env:LOCALAPPDATA\LINE\Data\Cache"
    $lineLogs  = "$env:LOCALAPPDATA\LINE\Log"
    $lf = [long]0
    $lf += Remove-Contents $lineCache
    $lf += Remove-Contents $lineLogs -Filter "*.log"
    if ($lf -gt 0) { Write-Host ""; Write-C "    LINE: $(Format-Bytes $lf)" -Color DarkGreen }
    $freed += $lf

    # ── Discord ────────────────────────────────────────────────
    # 只清以下快取，不碰 Local Storage / databases（存帳號和設定）
    $discordBase = "$env:APPDATA\Discord"
    $df = [long]0
    foreach ($cd in @("Cache","Code Cache","GPUCache","blob_storage")) {
        $df += Remove-Contents "$discordBase\$cd"
    }
    if ($df -gt 0) {
        if ($lf -eq 0) { Write-Host "" }
        Write-C "    Discord: $(Format-Bytes $df)" -Color DarkGreen
    }
    $freed += $df

    # ── Microsoft Teams (Classic) ──────────────────────────────
    $teamsBase = "$env:APPDATA\Microsoft\Teams"
    $tf = [long]0
    foreach ($cd in @("Cache","blob_storage","Code Cache","GPUCache","tmp")) {
        $tf += Remove-Contents "$teamsBase\$cd"
    }
    # Teams New (Win11)
    $teamsNew = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\Default\Cache"
    $tf += Remove-Contents $teamsNew
    if ($tf -gt 0) { Write-C "    Teams: $(Format-Bytes $tf)" -Color DarkGreen }
    $freed += $tf

    # ── Slack ──────────────────────────────────────────────────
    $slackBase = "$env:APPDATA\Slack"
    $sf = [long]0
    foreach ($cd in @("Cache","Code Cache","GPUCache")) {
        $sf += Remove-Contents "$slackBase\$cd"
    }
    if ($sf -gt 0) { Write-C "    Slack: $(Format-Bytes $sf)" -Color DarkGreen }
    $freed += $sf

    # ── Zoom (只清 logs) ───────────────────────────────────────
    $zoomLogs = "$env:APPDATA\Zoom\logs"
    $zf = Remove-Contents $zoomLogs -Filter "*.log"
    if ($zf -gt 0) { Write-C "    Zoom: $(Format-Bytes $zf)" -Color DarkGreen }
    $freed += $zf

    # ── Steam (只清網路圖片快取，不碰遊戲) ────────────────────
    # Find Steam path from registry; fall back to default
    $steamPath = ""
    try {
        $steamPath = (Get-ItemProperty "HKCU:\Software\Valve\Steam" -EA Stop).SteamPath
    } catch { }
    if (-not $steamPath) { $steamPath = "${env:ProgramFiles(x86)}\Steam" }
    $stf = [long]0
    $stf += Remove-Contents "$steamPath\appcache\httpcache"
    $stf += Remove-Contents "$steamPath\appcache\librarycache" -Filter "*.jpg"
    $stf += Remove-Contents "$steamPath\appcache\librarycache" -Filter "*.png"
    if ($stf -gt 0) { Write-C "    Steam: $(Format-Bytes $stf)" -Color DarkGreen }
    $freed += $stf

    # ── Skype ──────────────────────────────────────────────────
    $skypeCache = "$env:APPDATA\Microsoft\Skype"
    if (Test-Path "$skypeCache\Media") {
        $skf = Remove-Contents "$skypeCache\Media"
        if ($skf -gt 0) { Write-C "    Skype: $(Format-Bytes $skf)" -Color DarkGreen }
        $freed += $skf
    }

    if ($freed -eq 0) {
        Write-Host ""
        Write-C "  [v] 未找到相關 App / No apps found" -Color DarkGray
    }

    Record-Result "App Cache" "App 快取" $freed
    $script:TotalFreed += $freed
    Write-C "  [v] " -Color Green -NoNewLine
    Write-C "App 快取合計: " -Color White -NoNewLine
    Write-C (Format-Bytes $freed) -Color Yellow
}

function Clear-DeliveryOptimization {
    Show-TaskStart "Delivery Optimization" "Delivery Optimization 快取"
    $freed = [long]0

    # Windows 偷偷幫鄰居分享更新所產生的快取，完全安全可刪
    # Cache used for P2P Windows Update sharing - completely safe to delete
    $doPaths = @(
        "$env:windir\SoftwareDistribution\DeliveryOptimization",
        "$env:LOCALAPPDATA\Microsoft\Windows\DeliveryOptimization\Cache"
    )
    foreach ($p in $doPaths) { $freed += Remove-Contents $p }

    # PS5+ 有原生指令
    if ($script:PSMajor -ge 5) {
        try { Clear-DeliveryOptimizationCache -Force -EA SilentlyContinue } catch { }
    }

    Record-Result "Delivery Opt." "Delivery Optimization" $freed
    $script:TotalFreed += $freed
    Show-TaskDone $freed
}

function Clear-PrefetchFiles {
    Show-TaskStart "Prefetch Files" "預先擷取檔案"
    $freed  = Remove-Contents "$env:windir\Prefetch" -Filter "*.pf"
    $freed += Remove-Contents "$env:windir\Prefetch\ReadyBoot"
    Record-Result "Prefetch" "預先擷取檔案" $freed
    $script:TotalFreed += $freed
    Show-TaskDone $freed
}

function Clear-ThumbnailCache {
    Show-TaskStart "Thumbnail Cache" "縮圖快取"
    $freed     = [long]0
    $thumbPath = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"
    try {
        if ($script:WinVersion -ge 8) { & ie4uinit.exe -show 2>$null }
        if (Test-Path $thumbPath) {
            Stop-Process -Name explorer -Force -EA SilentlyContinue
            Start-Sleep -Milliseconds 800
            Get-ChildItem -LiteralPath $thumbPath -Filter "thumbcache_*.db" -EA SilentlyContinue |
            ForEach-Object {
                try { $freed += $_.Length; Remove-Item -LiteralPath $_.FullName -Force -EA Stop } catch { }
            }
            $iconCache = "$env:LOCALAPPDATA\IconCache.db"
            if (Test-Path $iconCache) {
                try { $freed += (Get-Item $iconCache -EA Stop).Length; Remove-Item $iconCache -Force -EA Stop } catch { }
            }
            Start-Process explorer
        }
    } catch { }
    Record-Result "Thumbnail" "縮圖快取" $freed
    $script:TotalFreed += $freed
    Show-TaskDone $freed
}

function Clear-DNSCache {
    Show-TaskStart "DNS Cache" "DNS 快取"
    try {
        & ipconfig /flushdns 2>&1 | Out-Null
        if ($script:PSMajor -ge 3 -and $script:WinVersion -ge 8) {
            Clear-DnsClientCache -EA SilentlyContinue
        }
        Record-Result "DNS Cache" "DNS 快取" 0
        Show-TaskDone 0 "DNS 快取已清除 / DNS cache flushed"
    } catch {
        Show-TaskSkip "無法清除"
    }
}

function Rebuild-FontCache {
    Show-TaskStart "Font Cache Rebuild" "字型快取重建"
    $freed = [long]0
    try {
        Stop-Service -Name FontCache  -Force -EA SilentlyContinue
        Stop-Service -Name FontCache3 -Force -EA SilentlyContinue
        Start-Sleep -Seconds 1
        foreach ($p in @(
            "$env:windir\ServiceProfiles\LocalService\AppData\Local\FontCache",
            "$env:windir\ServiceProfiles\LocalService\AppData\Local\FontCache-System",
            "$env:windir\system32\FNTCACHE.DAT"
        )) {
            if (Test-Path $p) {
                try {
                    $item = Get-Item $p -EA Stop
                    $size = if ($item.PSIsContainer) { Get-FolderSize $p } else { $item.Length }
                    Remove-Item $p -Recurse -Force -EA Stop
                    $freed += $size
                } catch { }
            }
        }
        Start-Service -Name FontCache -EA SilentlyContinue
    } catch { }
    Record-Result "Font Cache" "字型快取" $freed
    $script:TotalFreed += $freed
    Show-TaskDone $freed
}

function Clear-ErrorReports {
    Show-TaskStart "Windows Error Reports" "錯誤報告檔案"
    $freed = [long]0
    foreach ($p in @(
        "$env:ProgramData\Microsoft\Windows\WER\ReportQueue",
        "$env:ProgramData\Microsoft\Windows\WER\ReportArchive",
        "$env:LOCALAPPDATA\Microsoft\Windows\WER\ReportQueue",
        "$env:LOCALAPPDATA\Microsoft\Windows\WER\ReportArchive"
    )) { $freed += Remove-Contents $p }
    Record-Result "Error Reports" "錯誤報告" $freed
    $script:TotalFreed += $freed
    Show-TaskDone $freed
}

function Clear-WindowsOld {
    Show-TaskStart "Windows.old" "舊版 Windows 安裝"
    $path = "$env:SystemDrive\Windows.old"
    if (-not (Test-Path $path)) { Show-TaskSkip "找不到 Windows.old"; return }

    $before = Get-FolderSize $path
    Write-Host ""
    Write-C "  !! 注意：此資料夾約 $(Format-Bytes $before)，刪除後無法還原！" -Color Red
    Write-C "     這是升級前的舊系統備份，刪掉後如需退回舊版 Windows 將無法進行" -Color Yellow
    Write-C "     確認刪除? [Y/N]: " -Color Yellow -NoNewLine
    $confirm = Read-Host
    if ($confirm -notmatch "^[YySs是對]") { Show-TaskSkip "使用者取消"; return }

    & takeown /f "$path" /r /d y 2>&1 | Out-Null
    & icacls "$path" /grant administrators:F /t /q 2>&1 | Out-Null
    $freed = [long]0
    try {
        Remove-Item -LiteralPath $path -Recurse -Force -EA Stop
        $freed = $before
    } catch {
        & cmd.exe /c "rd /s /q `"$path`"" 2>$null
        $freed = if (Test-Path $path) { $before - (Get-FolderSize $path) } else { $before }
    }
    Record-Result "Windows.old" "舊版 Windows" $freed
    $script:TotalFreed += $freed
    Show-TaskDone $freed
}

function Clear-LogFiles {
    Show-TaskStart "System Log Files" "系統日誌"
    $freed = [long]0
    foreach ($spec in @(
        @{P="$env:windir\Logs\CBS";  F="*.log"},
        @{P="$env:windir\Logs\DISM"; F="*.log"},
        @{P="$env:windir\inf";       F="*.log"},
        @{P="$env:windir\Temp";      F="*.log"},
        @{P="$env:windir\Temp";      F="*.tmp"}
    )) { $freed += Remove-Contents $spec.P -Filter $spec.F }

    $cbsLog = "$env:windir\Logs\CBS\CBS.log"
    if (Test-Path $cbsLog) {
        try { $freed += (Get-Item $cbsLog).Length; Remove-Item $cbsLog -Force } catch { }
    }
    if ($script:WinVersion -ge 7) {
        foreach ($log in @("Application","System","Setup","HardwareEvents")) {
            try { & wevtutil cl $log 2>$null } catch { }
        }
    }
    Record-Result "Log Files" "系統日誌" $freed
    $script:TotalFreed += $freed
    Show-TaskDone $freed
}

function Optimize-RAM {
    Show-TaskStart "RAM Optimization" "記憶體最佳化"
    try {
        $wmi      = Get-WmiObject Win32_OperatingSystem -EA Stop
        $beforeMB = [long]([math]::Round($wmi.FreePhysicalMemory / 1KB, 0))
        $code = @'
using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
public class MemCleaner {
    [DllImport("kernel32.dll")]
    public static extern bool SetProcessWorkingSetSize(IntPtr proc, IntPtr min, IntPtr max);
    public static void EmptyAll() {
        foreach (Process p in Process.GetProcesses()) {
            try { SetProcessWorkingSetSize(p.Handle, (IntPtr)(-1), (IntPtr)(-1)); } catch { }
        }
    }
}
'@
        Add-Type -TypeDefinition $code -Language CSharp -EA Stop
        [MemCleaner]::EmptyAll()
        Start-Sleep -Seconds 1
        $wmiAfter = Get-WmiObject Win32_OperatingSystem
        $afterMB  = [long]([math]::Round($wmiAfter.FreePhysicalMemory / 1KB, 0))
        $gainMB   = [math]::Max(0, $afterMB - $beforeMB)
        Record-Result "RAM" "記憶體" ([long]($gainMB * 1MB))
        $script:TotalFreed += [long]($gainMB * 1MB)
        Write-Host ""
        Write-C "  [v] RAM: " -Color Green -NoNewLine
        Write-C "之前可用 ${beforeMB}MB → 現在可用 ${afterMB}MB " -Color White -NoNewLine
        Write-C "(+${gainMB}MB)" -Color Yellow
    } catch {
        Write-Host ""
        Write-C "  [-] 記憶體最佳化失敗" -Color DarkGray
    }
}

# ============================================================
# SECTION 6: STARTUP MANAGER / 啟動項目管理
# ============================================================

function Get-StartupItems {
    $items = @()
    $regPaths = @(
        @{Key="HKCU:\Software\Microsoft\Windows\CurrentVersion\Run";      Scope="目前使用者"},
        @{Key="HKLM:\Software\Microsoft\Windows\CurrentVersion\Run";      Scope="所有使用者"},
        @{Key="HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run"; Scope="32位元"}
    )
    $approvedBase = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"

    foreach ($rp in $regPaths) {
        try {
            $key = Get-Item $rp.Key -EA Stop
            foreach ($name in $key.GetValueNames()) {
                $val  = $key.GetValue($name)
                $isDisabled = $false
                # 檢查是否已被工作管理員停用
                try {
                    $approvedVal = (Get-ItemProperty $approvedBase -EA Stop).$name
                    if ($approvedVal -and $approvedVal[0] -eq 3) { $isDisabled = $true }
                } catch { }

                $items += [PSCustomObject]@{
                    Name       = $name
                    Path       = $val
                    Scope      = $rp.Scope
                    RegKey     = $rp.Key
                    IsDisabled = $isDisabled
                }
            }
        } catch { }
    }
    return $items
}

function Disable-StartupItem {
    param([string]$Name)
    # 使用跟工作管理員完全相同的方式停用（寫入 StartupApproved）
    # Uses exact same method as Task Manager - fully reversible via Task Manager
    $approvedPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"
    try {
        if (-not (Test-Path $approvedPath)) {
            New-Item $approvedPath -Force | Out-Null
        }
        # 0x03 = 停用 / disabled (跟工作管理員寫的值相同)
        $disabledValue = [byte[]](3,0,0,0,0,0,0,0,0,0,0,0)
        Set-ItemProperty -Path $approvedPath -Name $Name -Value $disabledValue -Type Binary -EA Stop
        return $true
    } catch {
        return $false
    }
}

function Manage-StartupItems {
    Show-Header
    Write-C "  啟動項目管理 / Startup Manager" -Color Cyan
    Write-C "  停用不需要的啟動項目，開機更快、系統更順" -Color DarkGray
    Write-C "  Disable unnecessary startup items for faster boot" -Color DarkGray
    Write-C "  ─────────────────────────────────────────────────────" -Color DarkGray
    Write-C ""
    Write-C "  正在掃描... / Scanning..." -Color Yellow
    Write-C ""

    $items = Get-StartupItems

    if ($items.Count -eq 0) {
        Write-C "  找不到啟動項目 / No startup items found." -Color DarkGray
        Write-C ""
        Write-C "  按任意鍵返回 / Press any key..." -Color DarkGray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        return
    }

    # 顯示項目清單 / Display items
    Write-C "  編號  狀態      名稱                          執行路徑" -Color Cyan
    Write-C "  ────  ────────  ────────────────────────────  ──────────────────────────" -Color DarkGray

    for ($i = 0; $i -lt $items.Count; $i++) {
        $item      = $items[$i]
        $num       = "[$($i+1)]".PadLeft(5)
        $statusTxt = if ($item.IsDisabled) { "[已停用]" } else { "[啟用中]" }
        $statusCol = if ($item.IsDisabled) { [ConsoleColor]::DarkGray } else { [ConsoleColor]::Green }
        $nameStr   = $item.Name.Substring(0, [Math]::Min($item.Name.Length, 28)).PadRight(28)
        $pathStr   = $item.Path
        if ($pathStr.Length -gt 45) { $pathStr = "..." + $pathStr.Substring($pathStr.Length - 42) }

        Write-C "  $num " -Color White -NoNewLine
        Write-C "$statusTxt  " -Color $statusCol -NoNewLine
        Write-C "$nameStr  " -Color White -NoNewLine
        Write-C $pathStr -Color DarkGray
    }

    Write-C ""
    Write-C "  ─────────────────────────────────────────────────────" -Color DarkGray
    Write-C ""
    Write-C "  [注意] 停用後可隨時從工作管理員 > 啟動 重新啟用" -Color Yellow
    Write-C "  [Note] Can re-enable anytime in Task Manager > Startup" -Color DarkGray
    Write-C ""
    Write-C "  輸入要停用的編號 (逗號分隔，如: 2,4,5)" -Color White
    Write-C "  Enter numbers to DISABLE (e.g. 2,4,5), or [Enter] to go back: " -Color DarkGray -NoNewLine

    $inputStr = (Read-Host).Trim()
    if ($inputStr -eq "") {
        Write-C "  沒有變更 / No changes made." -Color DarkGray
        Start-Sleep -Seconds 1
        return
    }

    $disabledCount = 0
    foreach ($part in $inputStr.Split(",")) {
        $n = 0
        if ([int]::TryParse($part.Trim(), [ref]$n) -and $n -ge 1 -and $n -le $items.Count) {
            $item = $items[$n - 1]
            if ($item.IsDisabled) {
                Write-C "  [$n] $($item.Name) 已經是停用狀態 / Already disabled" -Color DarkGray
            } else {
                $ok = Disable-StartupItem -Name $item.Name
                if ($ok) {
                    Write-C "  [v] 已停用 / Disabled: $($item.Name)" -Color Green
                    $disabledCount++
                } else {
                    Write-C "  [x] 停用失敗 / Failed: $($item.Name)" -Color Red
                }
            }
        }
    }

    Write-C ""
    if ($disabledCount -gt 0) {
        Write-C "  ★ 完成！已停用 $disabledCount 個啟動項目。重新啟動後生效。" -Color Green
        Write-C "  ★ Done! Disabled $disabledCount item(s). Takes effect after restart." -Color Green
    }
    Write-C ""
    Write-C "  按任意鍵返回 / Press any key..." -Color DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ============================================================
# SECTION 7: ORCHESTRATION / 協調執行
# ============================================================

function Invoke-FullClean {
    Show-Header
    Write-C "  開始全面清理... / Starting Full Clean..." -Color Green
    Write-C "  ─────────────────────────────────────────────────────" -Color DarkGray
    Write-C ""

    Clear-UserTemp
    Clear-SystemTemp
    Clear-WindowsUpdateCache
    Clear-RecycleBinSafe
    Clear-BrowserCache
    Clear-AppCache
    Clear-DeliveryOptimization
    Clear-PrefetchFiles
    Clear-ThumbnailCache
    Clear-DNSCache
    Rebuild-FontCache
    Clear-ErrorReports
    Clear-LogFiles
    Optimize-RAM

    Write-C ""
    Write-C "  ─────────────────────────────────────────────────────" -Color DarkGray
    Write-C ""

    # Windows.old 選擇性詢問
    $winOldPath = "$env:SystemDrive\Windows.old"
    if (Test-Path $winOldPath) {
        $size = Get-FolderSize $winOldPath
        Write-C "  偵測到 Windows.old ($(Format-Bytes $size))，是否一併清除? [Y/N]: " -Color Yellow -NoNewLine
        $ans = Read-Host
        if ($ans -match "^[YySs是對]") { Clear-WindowsOld }
    }

    Show-Summary
}

function Invoke-SelectiveClean {
    Show-Header
    $selected = Show-SelectiveMenu

    if ($selected.Count -eq 0) {
        Write-C "  未選擇任何項目 / No items selected." -Color Red
        Start-Sleep -Seconds 2
        return
    }

    Write-C ""
    Write-C "  開始選擇性清理... / Starting Selective Clean..." -Color Green
    Write-C "  ─────────────────────────────────────────────────────" -Color DarkGray
    Write-C ""

    $map = @{
        1  = { Clear-UserTemp }
        2  = { Clear-SystemTemp }
        3  = { Clear-WindowsUpdateCache }
        4  = { Clear-RecycleBinSafe }
        5  = { Clear-BrowserCache }
        6  = { Clear-AppCache }
        7  = { Clear-DeliveryOptimization }
        8  = { Clear-PrefetchFiles }
        9  = { Clear-ThumbnailCache }
        10 = { Clear-DNSCache }
        11 = { Rebuild-FontCache }
        12 = { Clear-ErrorReports }
        13 = { Clear-WindowsOld }
        14 = { Clear-LogFiles }
        15 = { Optimize-RAM }
    }

    foreach ($num in $selected) {
        if ($map.ContainsKey([int]$num)) { & $map[[int]$num] }
    }

    Show-Summary
}

# ============================================================
# SECTION 8: SUMMARY / 清理結果
# ============================================================

function Show-Summary {
    Write-C ""
    Write-C "  +========================================================+" -Color Cyan
    Write-C "  |              清理完成！/ Cleaning Complete!            |" -Color Cyan
    Write-C "  +========================================================+" -Color Cyan
    Write-C ""
    Write-C "  總共釋放 / Total Freed: " -Color White -NoNewLine
    Write-C (Format-Bytes $script:TotalFreed) -Color Yellow
    Write-C ""
    Write-C "  項目明細 / Breakdown:" -Color Cyan
    Write-C "  ─────────────────────────────────────────────────────" -Color DarkGray

    foreach ($r in $script:TaskResults) {
        $nameStr = ($r.CN + " / " + $r.EN).PadRight(36)
        $sizeStr = (Format-Bytes $r.Freed).PadLeft(12)
        $color   = if ($r.Freed -gt 100MB) { [ConsoleColor]::Yellow }
                   elseif ($r.Freed -gt 0)  { [ConsoleColor]::White }
                   else                     { [ConsoleColor]::DarkGray }
        Write-C "    $nameStr $sizeStr" -Color $color
    }

    Write-C "  ─────────────────────────────────────────────────────" -Color DarkGray
    Write-C ""

    if ($script:TotalFreed -gt 1GB) {
        Write-C "  ★★ 釋放超過 1GB！系統應明顯更順暢！" -Color Green
    } elseif ($script:TotalFreed -gt 100MB) {
        Write-C "  ★  清理成功！釋放了可觀的空間。" -Color Green
    } else {
        Write-C "  ✓  清理完成。系統狀態良好。" -Color DarkGreen
    }

    Write-C ""
    Write-C "  建議重新啟動電腦以完成所有變更 / Restart recommended" -Color Yellow
    Write-C ""
}

# ============================================================
# SECTION 9: MAIN LOOP / 主程式迴圈
# ============================================================

$script:WinVersion = Get-WindowsVersion

while ($true) {
    $script:TotalFreed  = [long]0
    $script:TaskResults = [System.Collections.ArrayList]@()

    Show-Header
    Show-MainMenu
    $choice = (Read-Host).Trim()

    switch ($choice.ToUpper()) {
        "1" {
            Invoke-FullClean
            Write-C "  按任意鍵返回選單 / Press any key..." -Color DarkGray
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        "2" {
            Invoke-SelectiveClean
            Write-C "  按任意鍵返回選單 / Press any key..." -Color DarkGray
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        "3" {
            Manage-StartupItems
        }
        "4" {
            Show-SystemInfo
        }
        { $_ -in "Q","離開","EXIT","QUIT" } {
            Write-C ""
            Write-C "  感謝使用！再見！/ Thank you! Goodbye!" -Color Cyan
            Write-C ""
            Start-Sleep -Seconds 1
            exit
        }
        default {
            Write-C "  無效選項，請重試 / Invalid choice." -Color Red
            Start-Sleep -Seconds 1
        }
    }
}