# ================================
# Office Recovery Toolkit v5.1
# PowerShell 5.1 Compatible
# ================================

[Console]::OutputEncoding = [Text.Encoding]::UTF8
$ErrorActionPreference = 'SilentlyContinue'

# -----------------------------
# Initial State
# -----------------------------
$script:Lang = if ((Get-UICulture).Name -eq 'zh-TW') { 'zh-TW' } else { 'en-US' }
$script:ScanRoot = if (Test-Path 'C:\RecoveredFiles') { 'C:\RecoveredFiles' } else { $PSScriptRoot }
$script:OutputRoot = Join-Path $PSScriptRoot 'Output'
$script:Results = @()
$script:LastStatus = ''
$script:LastSummary = ''
$script:LastHtmlReport = ''
$script:LastOrganizerLog = ''

# v4.1 / v5 settings
$script:OrganizeMode = 'Copy'   # Copy / Move
$script:OrganizePrimaryOnly = $true
$script:SelectedMenu = 0

# -----------------------------
# Text Helper
# -----------------------------
function T {
    param(
        [string]$Zh,
        [string]$En
    )
    if ($script:Lang -eq 'zh-TW') { return $Zh }
    return $En
}

# -----------------------------
# Utility
# -----------------------------
function Ensure-Folder {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Wait-Return {
    Write-Host ''
    Write-Host (T '按任意鍵返回主選單...' 'Press any key to return to the main menu...') -ForegroundColor Yellow
    try {
        [void][Console]::ReadKey($true)
    }
    catch {
        Read-Host | Out-Null
    }
}

function Update-Status {
    param(
        [string]$Status,
        [string]$Summary
    )
    $script:LastStatus = $Status
    $script:LastSummary = $Summary
}

function Get-SafeHtml {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    Add-Type -AssemblyName System.Web
    return [System.Web.HttpUtility]::HtmlEncode($Text)
}

function Get-SizeKB {
    param([Int64]$Bytes)
    if ($Bytes -le 0) { return 0 }
    return [Math]::Round(($Bytes / 1KB), 2)
}

function Get-ResultCount {
    return @($script:Results).Count
}

function Get-SafeFileName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }

    $safe = $Name.Trim()

    foreach ($c in [IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace($c, '_')
    }

    $safe = $safe -replace '\s+', ' '
    $safe = $safe -replace '^[\.\s]+', ''
    $safe = $safe -replace '[\.\s]+$', ''
    $safe = $safe -replace '_{2,}', '_'

    return $safe.Trim()
}

function Get-FileCategory {
    param([string]$Ext)

    switch ($Ext.ToLowerInvariant()) {
        '.xlsx' { return 'Excel' }
        '.xls'  { return 'Excel' }
        '.docx' { return 'Word' }
        '.doc'  { return 'Word' }
        '.pptx' { return 'PowerPoint' }
        '.ppt'  { return 'PowerPoint' }
        default { return 'Other' }
    }
}


# -----------------------------
# Console / LightBar UI Helpers
# -----------------------------
function Get-ShortDisplayText {
    param(
        [string]$Text,
        [int]$MaxLength = 80
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    if ($Text.Length -le $MaxLength) { return $Text }
    if ($MaxLength -lt 8) { return $Text.Substring(0, $MaxLength) }
    return $Text.Substring(0, $MaxLength - 3) + '...'
}

function Get-DisplayCellWidth {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return 0 }
    $w = 0
    foreach ($ch in $Text.ToCharArray()) {
        if ([int][char]$ch -gt 255) { $w += 2 } else { $w += 1 }
    }
    return $w
}

function Fit-DisplayText {
    param(
        [string]$Text,
        [int]$Width,
        [switch]$PadRight,
        [switch]$UseEllipsis
    )
    if ($Width -le 0) { return '' }
    if ($null -eq $Text) { $Text = '' }
    $s = [string]$Text
    $result = ''
    $used = 0
    $ellipsis = if ($UseEllipsis) { '...' } else { '' }
    $ellipsisWidth = 3
    foreach ($ch in $s.ToCharArray()) {
        $cw = if ([int][char]$ch -gt 255) { 2 } else { 1 }
        $limit = $Width
        if ($UseEllipsis) { $limit = $Width - $ellipsisWidth }
        if ($limit -lt 0) { $limit = 0 }
        if ($used + $cw -gt $limit -and $UseEllipsis -and ($used + $ellipsisWidth) -le $Width) {
            break
        }
        if ($used + $cw -gt $Width) { break }
        $result += $ch
        $used += $cw
    }
    if ($UseEllipsis -and (Get-DisplayCellWidth $s) -gt $Width) {
        if ((Get-DisplayCellWidth $result) + $ellipsisWidth -le $Width) {
            $result += $ellipsis
            $used += $ellipsisWidth
        }
    }
    if ($PadRight) {
        $pad = $Width - $used
        if ($pad -gt 0) { $result += (' ' * $pad) }
    }
    return $result
}

function Safe-SetCursorPosition {
    param([int]$Left, [int]$Top)
    try {
        if ($Left -lt 0) { $Left = 0 }
        if ($Top -lt 0) { $Top = 0 }
        [Console]::SetCursorPosition($Left, $Top)
    }
    catch {
    }
}

function Write-At {
    param(
        [int]$Left,
        [int]$Top,
        [string]$Text,
        [ConsoleColor]$Foreground = [ConsoleColor]::Gray,
        [ConsoleColor]$Background = [ConsoleColor]::Black,
        [switch]$NoPad,
        [int]$FixedWidth = 0,
        [switch]$Ellipsis
    )

    try {
        $width = [Console]::WindowWidth
        if ($width -lt 20) { return }
        if ($Left -ge $width) { return }

        Safe-SetCursorPosition -Left $Left -Top $Top

        $available = $width - $Left
        if ($available -lt 1) { return }
        if ($FixedWidth -gt 0 -and $FixedWidth -lt $available) {
            $available = $FixedWidth
        }

        $out = Fit-DisplayText -Text $Text -Width $available -PadRight:(-not $NoPad) -UseEllipsis:$Ellipsis

        $oldFg = [Console]::ForegroundColor
        $oldBg = [Console]::BackgroundColor
        [Console]::ForegroundColor = $Foreground
        [Console]::BackgroundColor = $Background
        [Console]::Write($out)
        [Console]::ForegroundColor = $oldFg
        [Console]::BackgroundColor = $oldBg
    }
    catch {
    }
}

function Clear-Line {
    param([int]$Top)
    try {
        $blank = ''.PadRight([Math]::Max([Console]::WindowWidth - 1, 1))
        Write-At -Left 0 -Top $Top -Text $blank -Foreground DarkGray -Background Black -NoPad
    }
    catch {
    }
}

function Draw-Frame {
    $w = [Console]::WindowWidth
    $h = [Console]::WindowHeight

    if ($w -lt 80 -or $h -lt 26) {
        Clear-Host
        Write-Host (T '視窗太小，請放大 PowerShell 視窗後再使用。' 'Console window too small. Please enlarge the PowerShell window.') -ForegroundColor Yellow
        Write-Host ('Width=' + $w + ' Height=' + $h) -ForegroundColor DarkYellow
        return $false
    }

    for ($i = 0; $i -lt $h; $i++) {
        Clear-Line -Top $i
    }

    Write-At 0 0  ('=' * ($w - 1)) Cyan Black
    Write-At 2 1  (T 'Office 檔案救援分析工具 v5 LightBar' 'Office Recovery Analyzer v5 LightBar') White DarkBlue
    Write-At 2 2  (T '↑↓ 光棒選擇  Enter 執行  數字快速鍵  L 切換語系  Esc 離開' '↑↓ Select  Enter Run  Number hotkeys  L switch language  Esc exit') Gray Black
    Write-At 0 3  ('=' * ($w - 1)) Cyan Black
    return $true
}

function Get-MenuItems {
    @(
        @{ Key='1'; Text=(T '開始掃描' 'Start Scan'); Action='Scan' },
        @{ Key='2'; Text=(T '匯出 HTML 報表' 'Export HTML Report'); Action='ExportHtml' },
        @{ Key='3'; Text=(T '開啟輸出資料夾' 'Open Output Folder'); Action='OpenOutput' },
        @{ Key='4'; Text=(T '設定掃描資料夾' 'Set Scan Folder'); Action='SetScanRoot' },
        @{ Key='5'; Text=(T '模擬自動改檔名' 'Preview Auto Rename'); Action='PreviewRename' },
        @{ Key='6'; Text=(T '實際自動改檔名' 'Apply Auto Rename'); Action='ApplyRename' },
        @{ Key='7'; Text=(T '開啟最新 HTML 報表' 'Open Latest HTML Report'); Action='OpenHtml' },
        @{ Key='8'; Text=(T '整理主檔/重複檔到資料夾' 'Organize Primary/Duplicate Files'); Action='Organize' },
        @{ Key='9'; Text=(T '切換 Copy/Move 模式' 'Toggle Copy/Move Mode'); Action='ToggleMode' },
        @{ Key='0'; Text=(T '切換是否只整理主檔' 'Toggle Primary Only Mode'); Action='TogglePrimaryOnly' },
        @{ Key='L'; Text=(T '切換語系' 'Switch Language'); Action='ToggleLang' },
        @{ Key='Esc'; Text=(T '離開' 'Exit'); Action='Exit' }
    )
}

function Get-MenuLayout {
    return [PSCustomObject]@{
        Top        = 6
        Left       = 4
        Width      = 42
        RightLeft  = 52
        RightWidth = [Math]::Max(([Console]::WindowWidth - 54), 20)
    }
}

function Draw-OneMenuItem {
    param(
        [int]$Index,
        [switch]$Selected
    )

    $menu = Get-MenuItems
    if ($Index -lt 0 -or $Index -ge $menu.Count) { return }

    $layout = Get-MenuLayout
    $item = $menu[$Index]
    $line = ('[{0}] {1}' -f $item.Key, $item.Text)
    $text = '  ' + $line
    $top = $layout.Top + $Index

    if ($Selected) {
        Write-At $layout.Left $top $text Black Gray -FixedWidth $layout.Width -Ellipsis
    }
    else {
        Write-At $layout.Left $top $text Gray Black -FixedWidth $layout.Width -Ellipsis
    }
}

function Update-LightBarSelection {
    param(
        [int]$OldIndex,
        [int]$NewIndex
    )

    if ($OldIndex -eq $NewIndex) { return }
    Draw-OneMenuItem -Index $OldIndex
    Draw-OneMenuItem -Index $NewIndex -Selected
}

function Draw-SettingsPanel {
    $layout = Get-MenuLayout
    $rightLeft = $layout.RightLeft
    $rightWidth = $layout.RightWidth

    Write-At $rightLeft 5 (T '目前設定' 'Current Settings') Yellow Black -FixedWidth $rightWidth
    Write-At $rightLeft 7 ((T '掃描路徑' 'Scan Root') + ' :') Gray Black -FixedWidth $rightWidth
    Write-At $rightLeft 8 $script:ScanRoot Cyan Black -FixedWidth $rightWidth -Ellipsis
    Write-At $rightLeft 10 ((T '輸出資料夾' 'Output Folder') + ' :') Gray Black -FixedWidth $rightWidth
    Write-At $rightLeft 11 $script:OutputRoot Cyan Black -FixedWidth $rightWidth -Ellipsis
    Write-At $rightLeft 13 ((T '最新 HTML' 'Latest HTML') + ' :') Gray Black -FixedWidth $rightWidth
    Write-At $rightLeft 14 $script:LastHtmlReport DarkCyan Black -FixedWidth $rightWidth -Ellipsis
    Write-At $rightLeft 16 ((T '最新整理紀錄' 'Latest Organize Log') + ' :') Gray Black -FixedWidth $rightWidth
    Write-At $rightLeft 17 $script:LastOrganizerLog DarkCyan Black -FixedWidth $rightWidth -Ellipsis
}

function Draw-LightBarMenu {
    $menu = Get-MenuItems
    $layout = Get-MenuLayout

    Write-At $layout.Left 5 (T '主選單' 'Main Menu') Yellow Black -FixedWidth $layout.Width

    for ($i = 0; $i -lt $menu.Count; $i++) {
        Draw-OneMenuItem -Index $i -Selected:($i -eq $script:SelectedMenu)
    }

    Draw-SettingsPanel
}

function Draw-StatusBar {
    $w = [Console]::WindowWidth
    $h = [Console]::WindowHeight
    $top = $h - 4
    $resultCount = Get-ResultCount

    Write-At 0 $top     (' '.PadRight($w - 1)) Black DarkCyan -NoPad
    Write-At 0 ($top+1) (' '.PadRight($w - 1)) Black DarkCyan -NoPad
    Write-At 0 ($top+2) (' '.PadRight($w - 1)) Black DarkCyan -NoPad
    Write-At 0 ($top+3) (' '.PadRight($w - 1)) Black DarkCyan -NoPad

    $line1 = '{0}: {1} | {2}: {3} | {4}: {5}' -f `
        (T '模式' 'Mode'), $script:OrganizeMode, `
        (T '只整理主檔' 'Primary Only'), $script:OrganizePrimaryOnly, `
        (T '語系' 'Language'), $script:Lang

    $line2 = '{0}: {1} | {2}: {3}' -f `
        (T '結果筆數' 'Result Count'), $resultCount, `
        (T '最新狀態' 'Last Status'), (Get-ShortDisplayText $script:LastStatus 45)

    $line3 = '{0}: {1}' -f (T '摘要' 'Summary'), (Get-ShortDisplayText $script:LastSummary 100)
    $line4 = (T '熱鍵：↑↓ 選擇 / Enter 執行 / 數字快速鍵 / L 語系 / Esc 離開' 'Hotkeys: ↑↓ select / Enter run / numbers / L language / Esc exit')

    Write-At 1 $top     $line1 Black DarkCyan -FixedWidth ($w - 2) -Ellipsis
    Write-At 1 ($top+1) $line2 Black DarkCyan -FixedWidth ($w - 2) -Ellipsis
    Write-At 1 ($top+2) $line3 Black DarkCyan -FixedWidth ($w - 2) -Ellipsis
    Write-At 1 ($top+3) $line4 Black DarkCyan -FixedWidth ($w - 2) -Ellipsis
}

# -----------------------------
# UI
# -----------------------------
function Draw-UI {
    $ok = Draw-Frame
    if (-not $ok) { return }
    Draw-LightBarMenu
    Draw-StatusBar
}

function Show-ProgressLine {
    param(
        [int]$Current,
        [int]$Total,
        [string]$FileName
    )

    $barWidth = 40
    if ($Total -gt 0) {
        $pct = [Math]::Floor(($Current * 100) / $Total)
    } else {
        $pct = 0
    }

    $filled = [Math]::Floor(($pct * $barWidth) / 100)
    $empty = $barWidth - $filled
    $bar = ('#' * $filled) + ('-' * $empty)

    Write-Progress -Activity (T '掃描中' 'Scanning') -Status $FileName -PercentComplete $pct
    Write-Host ('[{0}] {1,3}%  ({2}/{3})  {4}' -f $bar, $pct, $Current, $Total, $FileName)
}

# -----------------------------
# Hash / Content
# -----------------------------
function Get-TextHash {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

    $bytes = [Text.Encoding]::UTF8.GetBytes($Text)
    $stream = New-Object IO.MemoryStream(,$bytes)
    try {
        return (Get-FileHash -InputStream $stream -Algorithm SHA256).Hash
    }
    finally {
        $stream.Dispose()
    }
}

function Get-ZipEntryText {
    param(
        [string]$ZipPath,
        [string[]]$Candidates
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)
        foreach ($name in $Candidates) {
            $entry = $zip.Entries | Where-Object { $_.FullName -ieq $name } | Select-Object -First 1
            if ($entry) {
                $sr = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
                try {
                    return $sr.ReadToEnd()
                }
                finally {
                    $sr.Close()
                }
            }
        }
    }
    catch {
        return $null
    }
    finally {
        if ($zip) { $zip.Dispose() }
    }

    return $null
}

function Normalize-XmlText {
    param([string]$XmlText)

    if ([string]::IsNullOrWhiteSpace($XmlText)) { return '' }

    $t = $XmlText
    $t = [Regex]::Replace($t, '<[^>]+>', ' ')
    $t = $t.Replace('&amp;', '&').Replace('&lt;', '<').Replace('&gt;', '>').Replace('&quot;', '"').Replace('&apos;', "'")
    $t = [Regex]::Replace($t, '\s+', ' ').Trim()

    return $t
}

function Get-ExcelSmartName {
    param([string]$FilePath)

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = $null
    $sheetNames = @()
    $cellText = ''

    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)

        $wb = $zip.Entries | Where-Object { $_.FullName -ieq 'xl/workbook.xml' } | Select-Object -First 1
        if ($wb) {
            $sr = New-Object IO.StreamReader($wb.Open(), [Text.Encoding]::UTF8)
            try {
                $xml = $sr.ReadToEnd()
                $matches = [regex]::Matches($xml, 'name="([^"]+)"')
                foreach ($m in $matches) {
                    $sheetNames += $m.Groups[1].Value
                }
            }
            finally { $sr.Close() }
        }

        $sheetEntries = $zip.Entries |
            Where-Object { $_.FullName -match '^xl/worksheets/sheet\d+\.xml$' } |
            Sort-Object FullName

        foreach ($sheet in $sheetEntries) {
            $sr2 = New-Object IO.StreamReader($sheet.Open(), [Text.Encoding]::UTF8)
            try {
                $xml2 = $sr2.ReadToEnd()
                $plain = Normalize-XmlText $xml2

                if ($plain.Length -gt 5) {
                    $cellText = $plain.Substring(0, [Math]::Min(50, $plain.Length))
                    break
                }
            }
            finally { $sr2.Close() }
        }

        $nameParts = @()

        if ($sheetNames.Count -gt 0) {
            $nameParts += $sheetNames[0]
        }

        if ($cellText) {
            $nameParts += $cellText
        }

        $final = ($nameParts -join '_')
        $final = Get-SafeFileName $final

        if ($final.Length -gt 60) {
            $final = $final.Substring(0, 60)
        }

        return $final
    }
    catch {
        return ''
    }
    finally {
        if ($zip) { $zip.Dispose() }
    }
}

function Get-OfficeContentInfo {
    param([string]$FilePath)

    $ext = [IO.Path]::GetExtension($FilePath).ToLowerInvariant()

    $result = [ordered]@{
        ContentHash = ''
        PreviewText = ''
        ParseStatus = T '解析失敗' 'Parse failed'
        ParseReason = ''
    }

    try {
        switch ($ext) {
            '.docx' {
                $xml = Get-ZipEntryText -ZipPath $FilePath -Candidates @('word/document.xml')
                if ($xml) {
                    $plain = Normalize-XmlText $xml
                    if ($plain) {
                        $result.ContentHash = Get-TextHash $plain
                        $result.PreviewText = $plain.Substring(0, [Math]::Min(200, $plain.Length))
                        $result.ParseStatus = T '解析成功' 'Parsed'
                    }
                    else {
                        $result.ParseStatus = T '無法取得內容' 'No content extracted'
                    }
                }
                else {
                    $result.ParseStatus = T '無法取得內容' 'No content extracted'
                    $result.ParseReason = 'word/document.xml'
                }
            }

            '.xlsx' {
                Add-Type -AssemblyName System.IO.Compression.FileSystem

                $zip = $null
                $parts = @()

                try {
                    $zip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)

                    $sharedEntry = $zip.Entries | Where-Object { $_.FullName -ieq 'xl/sharedStrings.xml' } | Select-Object -First 1
                    if ($sharedEntry) {
                        $sr = New-Object IO.StreamReader($sharedEntry.Open(), [Text.Encoding]::UTF8)
                        try {
                            $sharedXml = $sr.ReadToEnd()
                            if (-not [string]::IsNullOrWhiteSpace($sharedXml)) {
                                $parts += $sharedXml
                            }
                        }
                        finally {
                            $sr.Close()
                        }
                    }

                    $sheetEntries = $zip.Entries |
                        Where-Object { $_.FullName -match '^xl/worksheets/sheet\d+\.xml$' } |
                        Sort-Object FullName

                    foreach ($sheetEntry in $sheetEntries) {
                        $sr2 = New-Object IO.StreamReader($sheetEntry.Open(), [Text.Encoding]::UTF8)
                        try {
                            $sheetXml = $sr2.ReadToEnd()
                            if (-not [string]::IsNullOrWhiteSpace($sheetXml)) {
                                $parts += $sheetXml
                            }
                        }
                        finally {
                            $sr2.Close()
                        }
                    }

                    $workbookEntry = $zip.Entries | Where-Object { $_.FullName -ieq 'xl/workbook.xml' } | Select-Object -First 1
                    if ($workbookEntry) {
                        $sr3 = New-Object IO.StreamReader($workbookEntry.Open(), [Text.Encoding]::UTF8)
                        try {
                            $workbookXml = $sr3.ReadToEnd()
                            if (-not [string]::IsNullOrWhiteSpace($workbookXml)) {
                                $parts += $workbookXml
                            }
                        }
                        finally {
                            $sr3.Close()
                        }
                    }

                    $combined = ($parts -join ' ')

                    if (-not [string]::IsNullOrWhiteSpace($combined)) {
                        $plain = Normalize-XmlText $combined
                        if (-not [string]::IsNullOrWhiteSpace($plain)) {
                            $result.ContentHash = Get-TextHash $plain
                            $result.PreviewText = $plain.Substring(0, [Math]::Min(200, $plain.Length))
                            $result.ParseStatus = T '解析成功' 'Parsed'
                        }
                        else {
                            $result.ParseStatus = T '無法取得內容' 'No content extracted'
                        }
                    }
                    else {
                        $result.ParseStatus = T '無法取得內容' 'No content extracted'
                        $result.ParseReason = 'xl/sharedStrings.xml + xl/worksheets/sheet*.xml + xl/workbook.xml'
                    }
                }
                catch {
                    $result.ParseStatus = T '解析失敗' 'Parse failed'
                    $result.ParseReason = $_.Exception.Message
                }
                finally {
                    if ($zip) { $zip.Dispose() }
                }
            }

            '.pptx' {
                $xml = Get-ZipEntryText -ZipPath $FilePath -Candidates @('ppt/slides/slide1.xml')
                if ($xml) {
                    $plain = Normalize-XmlText $xml
                    if ($plain) {
                        $result.ContentHash = Get-TextHash $plain
                        $result.PreviewText = $plain.Substring(0, [Math]::Min(200, $plain.Length))
                        $result.ParseStatus = T '解析成功' 'Parsed'
                    }
                    else {
                        $result.ParseStatus = T '無法取得內容' 'No content extracted'
                    }
                }
                else {
                    $result.ParseStatus = T '無法取得內容' 'No content extracted'
                    $result.ParseReason = 'ppt/slides/slide1.xml'
                }
            }

            '.doc' {
                $bytes = [IO.File]::ReadAllBytes($FilePath)
                $text = [Text.Encoding]::ASCII.GetString($bytes)
                $text = [Regex]::Replace($text, '[^\u0020-\u007E\u4E00-\u9FFF\r\n\t]', ' ')
                $text = [Regex]::Replace($text, '\s+', ' ').Trim()
                if ($text) {
                    $result.ContentHash = Get-TextHash $text
                    $result.PreviewText = $text.Substring(0, [Math]::Min(200, $text.Length))
                    $result.ParseStatus = T '解析成功' 'Parsed'
                }
                else {
                    $result.ParseStatus = T '無法取得內容' 'No content extracted'
                }
            }

            '.xls' {
                $bytes = [IO.File]::ReadAllBytes($FilePath)
                $text = [Text.Encoding]::ASCII.GetString($bytes)
                $text = [Regex]::Replace($text, '[^\u0020-\u007E\u4E00-\u9FFF\r\n\t]', ' ')
                $text = [Regex]::Replace($text, '\s+', ' ').Trim()
                if ($text) {
                    $result.ContentHash = Get-TextHash $text
                    $result.PreviewText = $text.Substring(0, [Math]::Min(200, $text.Length))
                    $result.ParseStatus = T '解析成功' 'Parsed'
                }
                else {
                    $result.ParseStatus = T '無法取得內容' 'No content extracted'
                }
            }

            default {
                $result.ParseStatus = T '不支援' 'Unsupported'
                $result.ParseReason = 'Unsupported extension'
            }
        }
    }
    catch {
        $result.ParseStatus = T '解析失敗' 'Parse failed'
        $result.ParseReason = $_.Exception.Message
    }

    return New-Object PSObject -Property $result
}

function Get-ExtensionLabel {
    param([string]$Ext)

    switch ($Ext.ToLowerInvariant()) {
        '.docx' { return (T 'Word 文件' 'Word Document') }
        '.xlsx' { return (T 'Excel 活頁簿' 'Excel Workbook') }
        '.pptx' { return (T 'PowerPoint 簡報' 'PowerPoint Presentation') }
        '.doc'  { return (T 'Word 舊版文件' 'Legacy Word Document') }
        '.xls'  { return (T 'Excel 舊版活頁簿' 'Legacy Excel Workbook') }
        default { return (T '未知' 'Unknown') }
    }
}

# -----------------------------
# Scoring / Grouping
# -----------------------------
function Get-FileQualityScore {
    param($Row)

    $score = 0

    if ($Row.ParseStatus -eq (T '解析成功' 'Parsed')) { $score += 50 }

    if ($Row.PreviewText) {
        $len = $Row.PreviewText.Length
        if ($len -gt 20) { $score += 30 }
        if ($len -gt 100) { $score += 20 }
    }

    if ($Row.SizeKB -gt 50) { $score += 10 }
    if ($Row.SizeKB -gt 500) { $score += 10 }

    if ($Row.PreviewText -match '[\u4e00-\u9fa5A-Za-z]{3,}') {
        $score += 20
    }

    return $score
}

function Apply-Grouping {
    param([array]$Rows)

    $groupId = 1

    $contentGroups = $Rows | Group-Object ContentHash | Where-Object { $_.Name -and $_.Count -gt 1 }
    foreach ($g in $contentGroups) {
        foreach ($row in $g.Group) {
            $row.LogicalGroup = 'CG-' + $groupId.ToString('0000')
            $row.DuplicateType = T '相同內容' 'Duplicate by Content'
        }
        $groupId++
    }

    $fileGroups = $Rows | Group-Object FileHash | Where-Object { $_.Name -and $_.Count -gt 1 }
    foreach ($g in $fileGroups) {
        foreach ($row in $g.Group) {
            if (-not $row.LogicalGroup) {
                $row.LogicalGroup = 'FG-' + $groupId.ToString('0000')
                $row.DuplicateType = T '相同檔案' 'Duplicate by File'
            }
        }
        $groupId++
    }

    foreach ($row in $Rows) {
        if (-not $row.LogicalGroup) {
            $row.LogicalGroup = ''
            if ($row.ParseStatus -eq (T '解析失敗' 'Parse failed')) {
                $row.DuplicateType = T '損毀/解析失敗' 'Corrupt/Parse Failed'
            }
            else {
                $row.DuplicateType = T '唯一' 'Unique'
            }
        }
    }

    return $Rows
}

function Set-PrimaryAndDuplicateRoles {
    param([array]$Rows)

    if (-not $Rows) { return $Rows }

    foreach ($r in $Rows) {
        if ([string]::IsNullOrWhiteSpace($r.LogicalGroup)) {
            if ($r.ParseStatus -eq (T '解析失敗' 'Parse failed')) {
                $r.Role = 'Broken'
                $r.RoleRank = 9999
            }
            else {
                $r.Role = 'Unique'
                $r.RoleRank = 1
            }
        }
    }

    $groups = $Rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.LogicalGroup) } | Group-Object LogicalGroup

    foreach ($g in $groups) {
        $ordered = $g.Group | Sort-Object `
            @{ Expression = { - (Get-FileQualityScore $_) } }, `
            @{ Expression = { -($_.SizeKB) } }, `
            @{ Expression = { $_.FileName } }

        $idx = 0
        foreach ($item in $ordered) {
            if ($idx -eq 0) {
                $item.Role = 'Primary'
                $item.RoleRank = 1
            }
            else {
                $item.Role = 'Duplicate'
                $item.RoleRank = $idx + 1
            }
            $idx++
        }
    }

    return $Rows
}

# -----------------------------
# Rename helpers
# -----------------------------
function Get-SuggestedBaseName {
    param($Row)

    $ext = $Row.Extension

    if ($ext -eq '.xlsx') {
        $smart = Get-ExcelSmartName -FilePath $Row.FullPath
        if (-not [string]::IsNullOrWhiteSpace($smart)) {
            if (-not [string]::IsNullOrWhiteSpace($Row.LogicalGroup)) {
                return '{0}_{1}' -f $smart, $Row.LogicalGroup
            }
            return $smart
        }
    }

    $text = $Row.PreviewText

    if ([string]::IsNullOrWhiteSpace($text)) {
        return ''
    }

    $name = $text
    $name = $name -replace '\s+', ' '
    $name = $name -replace '^[\-\_\.\s]+', ''
    $name = $name -replace '[\-\_\.\s]+$', ''
    $name = $name -replace 'sheet\d+', ''
    $name = $name -replace '工作表\d+', ''
    $name = $name -replace '^[0-9\s]+$', ''

    if ($name.Length -gt 60) {
        $name = $name.Substring(0, 60)
    }

    $name = Get-SafeFileName $name

    if ([string]::IsNullOrWhiteSpace($name)) {
        return ''
    }

    if (-not [string]::IsNullOrWhiteSpace($Row.LogicalGroup)) {
        $name = '{0}_{1}' -f $name, $Row.LogicalGroup
    }

    return $name
}

function Get-UniqueTargetPath {
    param(
        [string]$Folder,
        [string]$BaseName,
        [string]$Extension
    )

    $candidate = Join-Path $Folder ($BaseName + $Extension)

    if (-not (Test-Path -LiteralPath $candidate)) {
        return $candidate
    }

    for ($i = 1; $i -le 9999; $i++) {
        $newName = '{0}_{1:000}' -f $BaseName, $i
        $candidate = Join-Path $Folder ($newName + $Extension)
        if (-not (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    return $null
}

function Test-IsRecoveredGenericName {
    param([string]$FileName)

    if ([string]::IsNullOrWhiteSpace($FileName)) { return $false }

    if ($FileName -match '^(file|doc|xls|ppt|recovered|found|chk)[-_]?\d+(\.[^.]+)?$') {
        return $true
    }

    return $false
}

function Get-RenamePlan {
    param(
        [array]$Rows,
        [switch]$OnlyGenericNames
    )

    $plans = @()
    $dupCounters = @{}

    foreach ($row in $Rows) {
        if (-not (Test-Path -LiteralPath $row.FullPath)) {
            continue
        }

        if ($OnlyGenericNames -and -not (Test-IsRecoveredGenericName $row.FileName)) {
            continue
        }

        $folder = Split-Path -Parent $row.FullPath
        $ext = [IO.Path]::GetExtension($row.FullPath)
        $baseName = ''

        if ($row.Role -eq 'Primary') {
            $baseName = Get-SuggestedBaseName -Row $row
        }
        elseif ($row.Role -eq 'Duplicate') {
            $group = $row.LogicalGroup
            if ([string]::IsNullOrWhiteSpace($group)) {
                $group = 'DUP'
            }

            if (-not $dupCounters.ContainsKey($group)) {
                $dupCounters[$group] = 1
            }
            else {
                $dupCounters[$group]++
            }

            $baseName = 'DUP_{0}_{1:000}' -f $group, $dupCounters[$group]
        }
        elseif ($row.Role -eq 'Unique') {
            $baseName = Get-SuggestedBaseName -Row $row
        }
        else {
            continue
        }

        if ([string]::IsNullOrWhiteSpace($baseName)) {
            continue
        }

        $targetPath = Get-UniqueTargetPath -Folder $folder -BaseName $baseName -Extension $ext
        if ([string]::IsNullOrWhiteSpace($targetPath)) {
            continue
        }

        $targetName = Split-Path -Leaf $targetPath
        if ($targetName -ieq $row.FileName) {
            continue
        }

        $plans += New-Object PSObject -Property ([ordered]@{
            OriginalName  = $row.FileName
            OriginalPath  = $row.FullPath
            SuggestedName = $targetName
            SuggestedPath = $targetPath
            Extension     = $row.Extension
            PreviewText   = $row.PreviewText
            LogicalGroup  = $row.LogicalGroup
            DuplicateType = $row.DuplicateType
            Role          = $row.Role
        })
    }

    return $plans
}

# -----------------------------
# Scan
# -----------------------------
function Start-Scan {
    Clear-Host
    Ensure-Folder $script:OutputRoot

    if ([string]::IsNullOrWhiteSpace($script:ScanRoot)) {
        Write-Host (T '掃描路徑未設定。' 'Scan path is not set.') -ForegroundColor Red
        Update-Status -Status (T '失敗' 'Failed') -Summary (T '掃描路徑未設定' 'Scan path is not set')
        Wait-Return
        return
    }

    if (-not (Test-Path -LiteralPath $script:ScanRoot)) {
        Write-Host (T '掃描資料夾不存在。' 'Scan folder does not exist.') -ForegroundColor Red
        Write-Host $script:ScanRoot -ForegroundColor Yellow
        Update-Status -Status (T '失敗' 'Failed') -Summary (T '掃描資料夾不存在' 'Scan folder does not exist')
        Wait-Return
        return
    }

    $rootItem = Get-Item -LiteralPath $script:ScanRoot -ErrorAction SilentlyContinue
    if (-not $rootItem) {
        Write-Host (T '無法存取掃描路徑。' 'Cannot access scan path.') -ForegroundColor Red
        Write-Host $script:ScanRoot -ForegroundColor Yellow
        Update-Status -Status (T '失敗' 'Failed') -Summary (T '無法存取掃描路徑' 'Cannot access scan path')
        Wait-Return
        return
    }

    if (-not $rootItem.PSIsContainer) {
        Write-Host (T '掃描路徑不是資料夾。' 'Scan path is not a folder.') -ForegroundColor Red
        Write-Host $script:ScanRoot -ForegroundColor Yellow
        Update-Status -Status (T '失敗' 'Failed') -Summary (T '掃描路徑不是資料夾' 'Scan path is not a folder')
        Wait-Return
        return
    }

    Update-Status -Status (T '掃描中' 'Scanning') -Summary (T '開始掃描' 'Starting scan')

    $files = @()
    $exts = @('*.docx','*.xlsx','*.pptx','*.doc','*.xls')

    foreach ($e in $exts) {
        try {
            $files += Get-ChildItem -LiteralPath $script:ScanRoot -Recurse -File -Filter $e -ErrorAction SilentlyContinue
        }
        catch {
        }
    }

    $files = $files | Sort-Object FullName -Unique

    if (-not $files -or $files.Count -eq 0) {
        Write-Host (T '找不到支援的 Office 檔案。' 'No supported Office files found.') -ForegroundColor Yellow
        Write-Host ((T '掃描路徑' 'Scan Root') + ': ' + $script:ScanRoot) -ForegroundColor DarkCyan
        Update-Status -Status (T '失敗' 'Failed') -Summary (T '沒有檔案' 'No files')
        Wait-Return
        return
    }

    Write-Host ((T '掃描路徑' 'Scan Root') + ': ' + $script:ScanRoot) -ForegroundColor Cyan
    Write-Host ((T '總檔案數' 'Total Files') + ': ' + $files.Count) -ForegroundColor Cyan
    Write-Host ''

    $rows = @()
    $total = $files.Count

    for ($i = 0; $i -lt $total; $i++) {
        $f = $files[$i]
        Show-ProgressLine -Current ($i + 1) -Total $total -FileName $f.Name

        $fileHash = ''
        try {
            $fileHash = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256 -ErrorAction SilentlyContinue).Hash
        }
        catch {
            $fileHash = ''
        }

        $contentInfo = Get-OfficeContentInfo -FilePath $f.FullName

        $row = New-Object PSObject -Property ([ordered]@{
            FileName      = $f.Name
            FullPath      = $f.FullName
            Extension     = $f.Extension.ToLowerInvariant()
            ExtensionName = Get-ExtensionLabel $f.Extension
            SizeKB        = Get-SizeKB $f.Length
            FileHash      = $fileHash
            ContentHash   = $contentInfo.ContentHash
            PreviewText   = $contentInfo.PreviewText
            ParseStatus   = $contentInfo.ParseStatus
            ParseReason   = $contentInfo.ParseReason
            LogicalGroup  = ''
            DuplicateType = ''
            Role          = ''
            RoleRank      = 0
            ScanTime      = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        })

        $rows += $row
    }

    Write-Progress -Activity (T '掃描中' 'Scanning') -Completed

    $rows = Apply-Grouping -Rows $rows
    $rows = Set-PrimaryAndDuplicateRoles -Rows $rows
    $script:Results = $rows

    $csvPath = Join-Path $script:OutputRoot ('OfficeRecovery_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    $dupFileGroups = (@($rows | Group-Object FileHash | Where-Object { $_.Name -and $_.Count -gt 1 })).Count
    $dupContGroups = (@($rows | Group-Object ContentHash | Where-Object { $_.Name -and $_.Count -gt 1 })).Count
    $failCount = (@($rows | Where-Object { $_.ParseStatus -eq (T '解析失敗' 'Parse failed') })).Count

    Update-Status -Status (T '完成' 'Done') -Summary ("CSV: $csvPath")

    Write-Host ''
    Write-Host (T '掃描完成。' 'Scan completed.') -ForegroundColor Green
    Write-Host ('CSV : ' + $csvPath) -ForegroundColor Green
    Write-Host ((T '相同檔案群組' 'Duplicate File Groups') + ' : ' + $dupFileGroups)
    Write-Host ((T '相同內容群組' 'Duplicate Content Groups') + ' : ' + $dupContGroups)
    Write-Host ((T '解析失敗' 'Parse Failed') + ' : ' + $failCount)

    Wait-Return
}

# -----------------------------
# HTML export
# -----------------------------
function Export-HTML {
    Clear-Host

    if (-not $script:Results -or @($script:Results).Count -eq 0) {
        Write-Host (T '尚未有掃描結果。' 'No scan results yet.') -ForegroundColor Yellow
        Update-Status -Status (T '失敗' 'Failed') -Summary (T '尚未掃描' 'No scan results')
        Wait-Return
        return
    }

    Ensure-Folder $script:OutputRoot

    $htmlPath = Join-Path $script:OutputRoot ('OfficeRecovery_{0}.html' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

    $totalFiles = Get-ResultCount
    $dupFileGroups = (@($script:Results | Group-Object FileHash | Where-Object { $_.Name -and $_.Count -gt 1 })).Count
    $dupContGroups = (@($script:Results | Group-Object ContentHash | Where-Object { $_.Name -and $_.Count -gt 1 })).Count
    $failCount = (@($script:Results | Where-Object { $_.ParseStatus -eq (T '解析失敗' 'Parse failed') })).Count

    $primary = (@($script:Results | Where-Object { $_.Role -eq 'Primary' })).Count
    $dup = (@($script:Results | Where-Object { $_.Role -eq 'Duplicate' })).Count
    $unique = (@($script:Results | Where-Object { $_.Role -eq 'Unique' })).Count
    $fail = (@($script:Results | Where-Object { $_.ParseStatus -ne (T '解析成功' 'Parsed') })).Count

    try {
        $os = Get-CimInstance Win32_OperatingSystem
        $osText = '{0} ({1})' -f $os.Caption, $os.Version
    }
    catch {
        $osText = [Environment]::OSVersion.VersionString
    }

    $detailRows = New-Object System.Text.StringBuilder
    foreach ($r in $script:Results) {
        $roleDisplay = $r.Role
        if ($r.Role -eq 'Primary') {
            $roleDisplay = '⭐ Primary'
        }
        elseif ($r.Role -eq 'Duplicate') {
            $roleDisplay = 'Duplicate'
        }

        [void]$detailRows.AppendLine('<tr>')
        [void]$detailRows.AppendLine('<td>' + (Get-SafeHtml $r.FileName) + '</td>')
        [void]$detailRows.AppendLine('<td>' + (Get-SafeHtml $r.ExtensionName) + '</td>')
        [void]$detailRows.AppendLine('<td style="text-align:right">' + (Get-SafeHtml ([string]$r.SizeKB)) + '</td>')
        [void]$detailRows.AppendLine('<td>' + (Get-SafeHtml $r.ParseStatus) + '</td>')
        [void]$detailRows.AppendLine('<td>' + (Get-SafeHtml $r.DuplicateType) + '</td>')
        [void]$detailRows.AppendLine('<td>' + (Get-SafeHtml $roleDisplay) + '</td>')
        [void]$detailRows.AppendLine('<td style="text-align:right">' + (Get-SafeHtml ([string](Get-FileQualityScore $r))) + '</td>')
        [void]$detailRows.AppendLine('<td>' + (Get-SafeHtml $r.LogicalGroup) + '</td>')
        [void]$detailRows.AppendLine('<td style="font-family:Consolas,monospace;word-break:break-all">' + (Get-SafeHtml $r.FileHash) + '</td>')
        [void]$detailRows.AppendLine('<td style="font-family:Consolas,monospace;word-break:break-all">' + (Get-SafeHtml $r.ContentHash) + '</td>')
        [void]$detailRows.AppendLine('<td>' + (Get-SafeHtml $r.PreviewText) + '</td>')
        [void]$detailRows.AppendLine('<td>' + (Get-SafeHtml $r.ParseReason) + '</td>')
        [void]$detailRows.AppendLine('</tr>')
    }

    $title = Get-SafeHtml (T 'Office 檔案救援分析報表 v5' 'Office Recovery Analysis Report v5')
    $summaryText = Get-SafeHtml (T '摘要' 'Summary')
    $detailText = Get-SafeHtml (T '明細' 'Details')
    $customerSummary = Get-SafeHtml (T '客戶報告摘要' 'Customer Summary')

    $html = @"
<!DOCTYPE html>
<html lang="$($script:Lang)">
<head>
<meta charset="utf-8" />
<title>$title</title>
<style>
body{font-family:"Segoe UI","Microsoft JhengHei",Arial,sans-serif;background:#f3f6fb;color:#1f2937;margin:0;padding:0}
.wrap{max-width:1800px;margin:24px auto;padding:0 20px}
h1{margin:0 0 10px 0;font-size:30px}
.sub{color:#64748b;margin-bottom:20px}
.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin:20px 0}
.card{background:#fff;border-radius:16px;box-shadow:0 4px 16px rgba(0,0,0,.08);padding:18px}
.card .k{font-size:13px;color:#6b7280}
.card .v{font-size:28px;font-weight:700;margin-top:8px}
.panel{background:#fff;border-radius:16px;box-shadow:0 4px 16px rgba(0,0,0,.08);padding:18px;margin-bottom:20px}
table{width:100%;border-collapse:collapse}
th,td{border:1px solid #dbe4f0;padding:8px 10px;vertical-align:top;text-align:left;font-size:13px}
th{background:#eaf2ff}
.search{width:100%;padding:10px 12px;border:1px solid #cbd5e1;border-radius:10px;margin:10px 0 16px 0;font-size:14px}
.footer{margin-top:24px;color:#64748b;font-size:12px}
.small{font-size:13px;color:#475569}
</style>
<script>
function filterTable() {
  var input = document.getElementById("searchBox");
  var filter = input.value.toLowerCase();
  var rows = document.querySelectorAll("#detailBody tr");
  for (var i = 0; i < rows.length; i++) {
    var txt = rows[i].innerText.toLowerCase();
    rows[i].style.display = txt.indexOf(filter) > -1 ? "" : "none";
  }
}
</script>
</head>
<body>
<div class="wrap">
    <h1>$title</h1>
    <div class="sub">Product-grade recovery analysis report</div>

    <div class="grid">
        <div class="card">
            <div class="k">$(Get-SafeHtml (T '總檔案數' 'Total Files'))</div>
            <div class="v">$totalFiles</div>
        </div>
        <div class="card">
            <div class="k">$(Get-SafeHtml (T '相同檔案群組' 'Duplicate File Groups'))</div>
            <div class="v">$dupFileGroups</div>
        </div>
        <div class="card">
            <div class="k">$(Get-SafeHtml (T '相同內容群組' 'Duplicate Content Groups'))</div>
            <div class="v">$dupContGroups</div>
        </div>
        <div class="card">
            <div class="k">$(Get-SafeHtml (T '解析失敗' 'Parse Failed'))</div>
            <div class="v">$failCount</div>
        </div>
    </div>

    <div class="grid">
        <div class="card">
            <div class="k">Primary</div>
            <div class="v">$primary</div>
        </div>
        <div class="card">
            <div class="k">Duplicate</div>
            <div class="v">$dup</div>
        </div>
        <div class="card">
            <div class="k">Unique</div>
            <div class="v">$unique</div>
        </div>
        <div class="card">
            <div class="k">$(Get-SafeHtml (T '非成功解析' 'Non-Parsed'))</div>
            <div class="v">$fail</div>
        </div>
    </div>

    <div class="panel">
        <h2>$customerSummary</h2>
        <div class="small">
            $(Get-SafeHtml (T '總檔案' 'Total Files')): $totalFiles<br>
            $(Get-SafeHtml (T '成功解析' 'Parsed Successfully')): $(($totalFiles - $fail))<br>
            Primary: $primary<br>
            Duplicate: $dup<br>
            Unique: $unique<br>
            $(Get-SafeHtml (T '無法完整解析' 'Not Fully Parsed')): $fail
        </div>
    </div>

    <div class="panel">
        <h2>$summaryText</h2>
        <table>
            <tr><th>$(Get-SafeHtml (T '電腦名稱' 'Computer Name'))</th><td>$(Get-SafeHtml $env:COMPUTERNAME)</td></tr>
            <tr><th>$(Get-SafeHtml (T '作業系統' 'Operating System'))</th><td>$(Get-SafeHtml $osText)</td></tr>
            <tr><th>$(Get-SafeHtml (T '使用者' 'User'))</th><td>$(Get-SafeHtml $env:USERNAME)</td></tr>
            <tr><th>$(Get-SafeHtml (T '掃描路徑' 'Scan Root'))</th><td>$(Get-SafeHtml $script:ScanRoot)</td></tr>
            <tr><th>$(Get-SafeHtml (T '報表時間' 'Report Time'))</th><td>$(Get-SafeHtml ((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')))</td></tr>
        </table>
    </div>

    <div class="panel">
        <h2>$detailText</h2>
        <input type="text" id="searchBox" class="search" onkeyup="filterTable()" placeholder="Search / 搜尋">
        <table>
            <thead>
                <tr>
                    <th>$(Get-SafeHtml (T '檔名' 'File Name'))</th>
                    <th>$(Get-SafeHtml (T '類型' 'Type'))</th>
                    <th>$(Get-SafeHtml (T '大小(KB)' 'Size(KB)'))</th>
                    <th>$(Get-SafeHtml (T '狀態' 'Status'))</th>
                    <th>$(Get-SafeHtml (T '重複判定' 'Duplicate Type'))</th>
                    <th>$(Get-SafeHtml (T '角色' 'Role'))</th>
                    <th>$(Get-SafeHtml (T '品質分數' 'Quality Score'))</th>
                    <th>$(Get-SafeHtml (T '群組' 'Group'))</th>
                    <th>$(Get-SafeHtml (T '檔案雜湊' 'File Hash'))</th>
                    <th>$(Get-SafeHtml (T '內容指紋' 'Content Hash'))</th>
                    <th>$(Get-SafeHtml (T '內容預覽' 'Preview Text'))</th>
                    <th>$(Get-SafeHtml (T '說明' 'Reason'))</th>
                </tr>
            </thead>
            <tbody id="detailBody">
                $($detailRows.ToString())
            </tbody>
        </table>
    </div>

    <div class="footer">Generated by OfficeRecoveryToolkit.ps1 v5</div>
</div>
</body>
</html>
"@

    [IO.File]::WriteAllText($htmlPath, $html, [Text.Encoding]::UTF8)
    $script:LastHtmlReport = $htmlPath

    Update-Status -Status (T '完成' 'Done') -Summary ("HTML: $htmlPath")

    Write-Host (T 'HTML 報表已輸出。' 'HTML report exported.') -ForegroundColor Green
    Write-Host $htmlPath -ForegroundColor Green
    Write-Host ''
    Write-Host (T '之後可按 [7] 用系統預設瀏覽器開啟最新 HTML 報表。' 'You can press [7] later to open the latest HTML report with the default browser.') -ForegroundColor Cyan
    Wait-Return
}

function Open-LatestHtmlReport {
    Clear-Host

    if ([string]::IsNullOrWhiteSpace($script:LastHtmlReport)) {
        Write-Host (T '尚未匯出 HTML 報表。' 'No HTML report has been exported yet.') -ForegroundColor Yellow
        Update-Status -Status (T '失敗' 'Failed') -Summary (T '尚未匯出 HTML' 'No HTML exported yet')
        Wait-Return
        return
    }

    if (-not (Test-Path -LiteralPath $script:LastHtmlReport)) {
        Write-Host (T '找不到 HTML 報表檔案。' 'HTML report file not found.') -ForegroundColor Red
        Write-Host $script:LastHtmlReport -ForegroundColor Yellow
        Update-Status -Status (T '失敗' 'Failed') -Summary (T 'HTML 檔案不存在' 'HTML file does not exist')
        Wait-Return
        return
    }

    try {
        Start-Process -FilePath $script:LastHtmlReport | Out-Null
        Update-Status -Status (T '完成' 'Done') -Summary $script:LastHtmlReport
    }
    catch {
        Write-Host (T '無法開啟 HTML 報表。' 'Failed to open HTML report.') -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Yellow
        Update-Status -Status (T '失敗' 'Failed') -Summary (T '無法開啟 HTML 報表' 'Failed to open HTML report')
        Wait-Return
    }
}

# -----------------------------
# Rename preview / apply
# -----------------------------
function Preview-RenamePlan {
    Clear-Host

    if (-not $script:Results -or @($script:Results).Count -eq 0) {
        Write-Host (T '尚未有掃描結果。' 'No scan results yet.') -ForegroundColor Yellow
        Update-Status -Status (T '失敗' 'Failed') -Summary (T '尚未掃描' 'No scan results')
        Wait-Return
        return
    }

    $plans = Get-RenamePlan -Rows $script:Results -OnlyGenericNames

    if (-not $plans -or $plans.Count -eq 0) {
        Write-Host (T '沒有可自動改名的檔案。' 'No files available for automatic renaming.') -ForegroundColor Yellow
        Update-Status -Status (T '完成' 'Done') -Summary (T '沒有可改名項目' 'No rename candidates')
        Wait-Return
        return
    }

    Ensure-Folder $script:OutputRoot
    $csvPath = Join-Path $script:OutputRoot ('RenamePreview_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $plans | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    Write-Host (T '以下為模擬改名結果（不會真的改檔名）:' 'Rename simulation results (no files will actually be renamed):') -ForegroundColor Cyan
    Write-Host ''

    $show = $plans | Select-Object -First 20
    foreach ($p in $show) {
        Write-Host ('[OLD] ' + $p.OriginalName) -ForegroundColor Gray
        Write-Host ('[NEW] ' + $p.SuggestedName + '   [' + $p.Role + ']') -ForegroundColor Green
        Write-Host ''
    }

    if ($plans.Count -gt 20) {
        Write-Host ((T '僅顯示前 20 筆，完整結果請看 CSV：' 'Showing first 20 only. Full result saved to CSV:') + ' ' + $csvPath) -ForegroundColor Yellow
    }
    else {
        Write-Host ('CSV : ' + $csvPath) -ForegroundColor Green
    }

    Update-Status -Status (T '完成' 'Done') -Summary ("Rename Preview CSV: $csvPath")
    Wait-Return
}

function Invoke-AutoRename {
    Clear-Host

    if (-not $script:Results -or @($script:Results).Count -eq 0) {
        Write-Host (T '尚未有掃描結果。' 'No scan results yet.') -ForegroundColor Yellow
        Update-Status -Status (T '失敗' 'Failed') -Summary (T '尚未掃描' 'No scan results')
        Wait-Return
        return
    }

    $plans = Get-RenamePlan -Rows $script:Results -OnlyGenericNames

    if (-not $plans -or $plans.Count -eq 0) {
        Write-Host (T '沒有可自動改名的檔案。' 'No files available for automatic renaming.') -ForegroundColor Yellow
        Update-Status -Status (T '完成' 'Done') -Summary (T '沒有可改名項目' 'No rename candidates')
        Wait-Return
        return
    }

    Write-Host (T '即將進行實際改名。' 'About to perform actual renaming.') -ForegroundColor Yellow
    Write-Host ((T '符合條件的檔案數量' 'Number of eligible files') + ' : ' + $plans.Count)
    Write-Host ''
    Write-Host (T '請輸入 YES 確認執行：' 'Type YES to confirm:') -ForegroundColor Cyan

    $confirm = Read-Host
    if ($confirm -ne 'YES') {
        Update-Status -Status (T '取消' 'Cancelled') -Summary (T '使用者取消改名' 'Rename cancelled by user')
        return
    }

    Ensure-Folder $script:OutputRoot

    $log = @()
    $success = 0
    $failed = 0

    foreach ($p in $plans) {
        try {
            Rename-Item -LiteralPath $p.OriginalPath -NewName $p.SuggestedName -ErrorAction Stop

            $log += New-Object PSObject -Property ([ordered]@{
                OriginalName  = $p.OriginalName
                SuggestedName = $p.SuggestedName
                Role          = $p.Role
                Status        = 'Renamed'
                Reason        = ''
            })

            $success++
        }
        catch {
            $log += New-Object PSObject -Property ([ordered]@{
                OriginalName  = $p.OriginalName
                SuggestedName = $p.SuggestedName
                Role          = $p.Role
                Status        = 'Failed'
                Reason        = $_.Exception.Message
            })

            $failed++
        }
    }

    $csvPath = Join-Path $script:OutputRoot ('RenameLog_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $log | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    Write-Host ''
    Write-Host (T '自動改名完成。' 'Automatic renaming completed.') -ForegroundColor Green
    Write-Host ((T '成功' 'Succeeded') + ' : ' + $success) -ForegroundColor Green
    Write-Host ((T '失敗' 'Failed') + ' : ' + $failed) -ForegroundColor Yellow
    Write-Host ('CSV : ' + $csvPath) -ForegroundColor Green

    Update-Status -Status (T '完成' 'Done') -Summary ("Rename Log CSV: $csvPath")

    $currentRoot = $script:ScanRoot
    Start-Scan
    $script:ScanRoot = $currentRoot
}

# -----------------------------
# Organize
# -----------------------------
function Invoke-OrganizeFiles {
    Clear-Host

    if (-not $script:Results -or @($script:Results).Count -eq 0) {
        Write-Host (T '尚未有掃描結果。' 'No scan results yet.') -ForegroundColor Yellow
        Update-Status -Status (T '失敗' 'Failed') -Summary (T '尚未掃描' 'No scan results')
        Wait-Return
        return
    }

    $baseFolder = Join-Path $script:OutputRoot '整理結果'

    Write-Host "====================================="
    Write-Host (T '整理模式設定' 'Organize Mode Settings') -ForegroundColor Cyan
    Write-Host "====================================="

    Write-Host ((T '模式' 'Mode') + ' : ' + $script:OrganizeMode)
    Write-Host ((T '只整理主檔' 'Primary Only') + ' : ' + $script:OrganizePrimaryOnly)
    Write-Host ('Base : ' + $baseFolder)
    Write-Host ''

    Write-Host (T '請輸入 YES 確認執行：' 'Type YES to confirm:') -ForegroundColor Yellow
    $confirm = Read-Host
    if ($confirm -ne 'YES') { return }

    $log = @()
    $success = 0
    $failed = 0

    foreach ($r in $script:Results) {
        if (-not (Test-Path -LiteralPath $r.FullPath)) { continue }

        if ($script:OrganizePrimaryOnly -and $r.Role -ne 'Primary') {
            continue
        }

        $category = Get-FileCategory $r.Extension

        $roleFolder = switch ($r.Role) {
            'Primary'   { 'Primary' }
            'Duplicate' { 'Duplicates' }
            'Unique'    { 'Unique' }
            default     { 'Other' }
        }

        $targetFolder = Join-Path $baseFolder "$roleFolder\$category"
        Ensure-Folder $targetFolder

        $baseName = [IO.Path]::GetFileNameWithoutExtension($r.FileName)
        $ext = [IO.Path]::GetExtension($r.FileName)

        if ($r.Role -eq 'Primary') {
            $smart = Get-SuggestedBaseName -Row $r
            if ($smart) { $baseName = $smart }
        }

        if ($r.Role -eq 'Duplicate') {
            $baseName = "DUP_$($r.LogicalGroup)"
        }

        $targetPath = Get-UniqueTargetPath -Folder $targetFolder -BaseName $baseName -Extension $ext

        try {
            if ($script:OrganizeMode -eq 'Copy') {
                Copy-Item -LiteralPath $r.FullPath -Destination $targetPath -Force -ErrorAction Stop
            }
            else {
                Move-Item -LiteralPath $r.FullPath -Destination $targetPath -Force -ErrorAction Stop
            }

            $log += New-Object PSObject -Property @{
                Source = $r.FullPath
                Target = $targetPath
                Role   = $r.Role
                Mode   = $script:OrganizeMode
                Status = 'OK'
            }

            $success++
        }
        catch {
            $failed++
            $log += New-Object PSObject -Property @{
                Source = $r.FullPath
                Target = $targetPath
                Role   = $r.Role
                Mode   = $script:OrganizeMode
                Status = $_.Exception.Message
            }
        }
    }

    Ensure-Folder $script:OutputRoot
    $csv = Join-Path $script:OutputRoot ("OrganizeLog_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $log | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    $script:LastOrganizerLog = $csv

    Write-Host ''
    Write-Host (T '整理完成' 'Completed') -ForegroundColor Green
    Write-Host ("OK: $success  FAIL: $failed")
    Write-Host $csv

    Wait-Return
}

function Toggle-OrganizeMode {
    if ($script:OrganizeMode -eq 'Copy') {
        $script:OrganizeMode = 'Move'
    } else {
        $script:OrganizeMode = 'Copy'
    }
}

function Toggle-PrimaryOnly {
    $script:OrganizePrimaryOnly = -not $script:OrganizePrimaryOnly
}

# -----------------------------
# Set scan root
# -----------------------------
function Set-ScanRoot {
    Clear-Host
    Write-Host (T '請輸入掃描資料夾路徑：' 'Enter scan folder path:') -ForegroundColor Cyan
    $inputPath = Read-Host

    if ([string]::IsNullOrWhiteSpace($inputPath)) {
        Update-Status -Status (T '就緒' 'Ready') -Summary $script:ScanRoot
        return
    }

    if (Test-Path $inputPath) {
        $script:ScanRoot = $inputPath
        Update-Status -Status (T '完成' 'Done') -Summary $script:ScanRoot
    }
    else {
        Write-Host (T '資料夾不存在。' 'Folder does not exist.') -ForegroundColor Red
        Update-Status -Status (T '失敗' 'Failed') -Summary (T '資料夾不存在' 'Folder does not exist')
        Wait-Return
    }
}



function Test-KeyMatch {
    param(
        $KeyInfo,
        [string]$KeyName,
        [int]$VirtualKeyCode = -1,
        [string[]]$Chars = @()
    )

    if ($null -eq $KeyInfo) { return $false }

    try {
        if ($KeyName -and $KeyInfo.Key.ToString() -eq $KeyName) { return $true }
    }
    catch {}

    try {
        if ($VirtualKeyCode -ge 0 -and $KeyInfo.VirtualKeyCode -eq $VirtualKeyCode) { return $true }
    }
    catch {}

    try {
        $ch = [string]$KeyInfo.KeyChar
        if ($Chars -and ($Chars -contains $ch)) { return $true }
    }
    catch {}

    return $false
}

# -----------------------------
# Bootstrap
# -----------------------------
Ensure-Folder $script:OutputRoot
Update-Status -Status (T '就緒' 'Ready') -Summary ''

[Console]::CursorVisible = $false
try {
    $needsFullRedraw = $true
    while ($true) {
        if ($needsFullRedraw) {
            Draw-UI
            $needsFullRedraw = $false
        }

        $key = [Console]::ReadKey($true)
        $menuItems = Get-MenuItems

        if (Test-KeyMatch -KeyInfo $key -KeyName 'UpArrow' -VirtualKeyCode 38) {
            $old = $script:SelectedMenu
            if ($script:SelectedMenu -gt 0) { $script:SelectedMenu-- } else { $script:SelectedMenu = $menuItems.Count - 1 }
            Update-LightBarSelection -OldIndex $old -NewIndex $script:SelectedMenu
            continue
        }

        if (Test-KeyMatch -KeyInfo $key -KeyName 'DownArrow' -VirtualKeyCode 40) {
            $old = $script:SelectedMenu
            if ($script:SelectedMenu -lt ($menuItems.Count - 1)) { $script:SelectedMenu++ } else { $script:SelectedMenu = 0 }
            Update-LightBarSelection -OldIndex $old -NewIndex $script:SelectedMenu
            continue
        }

        if (Test-KeyMatch -KeyInfo $key -KeyName 'L' -VirtualKeyCode 76 -Chars @('l','L')) {
            if ($script:Lang -eq 'zh-TW') {
                $script:Lang = 'en-US'
            }
            else {
                $script:Lang = 'zh-TW'
            }
            Update-Status -Status (T '就緒' 'Ready') -Summary ((T '語系已切換' 'Language switched') + ': ' + $script:Lang)
            $needsFullRedraw = $true
            continue
        }

        $action = $null

        if (Test-KeyMatch -KeyInfo $key -KeyName 'Enter' -VirtualKeyCode 13) {
            $action = $menuItems[$script:SelectedMenu].Action
        }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D1' -VirtualKeyCode 49 -Chars @('1')) { $script:SelectedMenu = 0; $action = 'Scan' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D2' -VirtualKeyCode 50 -Chars @('2')) { $script:SelectedMenu = 1; $action = 'ExportHtml' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D3' -VirtualKeyCode 51 -Chars @('3')) { $script:SelectedMenu = 2; $action = 'OpenOutput' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D4' -VirtualKeyCode 52 -Chars @('4')) { $script:SelectedMenu = 3; $action = 'SetScanRoot' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D5' -VirtualKeyCode 53 -Chars @('5')) { $script:SelectedMenu = 4; $action = 'PreviewRename' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D6' -VirtualKeyCode 54 -Chars @('6')) { $script:SelectedMenu = 5; $action = 'ApplyRename' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D7' -VirtualKeyCode 55 -Chars @('7')) { $script:SelectedMenu = 6; $action = 'OpenHtml' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D8' -VirtualKeyCode 56 -Chars @('8')) { $script:SelectedMenu = 7; $action = 'Organize' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D9' -VirtualKeyCode 57 -Chars @('9')) { $script:SelectedMenu = 8; $action = 'ToggleMode' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D0' -VirtualKeyCode 48 -Chars @('0')) { $script:SelectedMenu = 9; $action = 'TogglePrimaryOnly' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'Escape' -VirtualKeyCode 27) { $action = 'Exit' }

        switch ($action) {
            'Scan' { Start-Scan; $needsFullRedraw = $true; continue }
            'ExportHtml' { Export-HTML; $needsFullRedraw = $true; continue }
            'OpenOutput' {
                Ensure-Folder $script:OutputRoot
                Start-Process explorer.exe $script:OutputRoot | Out-Null
                Update-Status -Status (T '完成' 'Done') -Summary $script:OutputRoot
                Draw-StatusBar
                Draw-SettingsPanel
                continue
            }
            'SetScanRoot' {
                [Console]::CursorVisible = $true
                Set-ScanRoot
                [Console]::CursorVisible = $false
                $needsFullRedraw = $true
                continue
            }
            'PreviewRename' { Preview-RenamePlan; $needsFullRedraw = $true; continue }
            'ApplyRename' {
                [Console]::CursorVisible = $true
                Invoke-AutoRename
                [Console]::CursorVisible = $false
                $needsFullRedraw = $true
                continue
            }
            'OpenHtml' { Open-LatestHtmlReport; $needsFullRedraw = $true; continue }
            'Organize' {
                [Console]::CursorVisible = $true
                Invoke-OrganizeFiles
                [Console]::CursorVisible = $false
                $needsFullRedraw = $true
                continue
            }
            'ToggleMode' {
                Toggle-OrganizeMode
                Update-Status -Status (T '完成' 'Done') -Summary ((T '整理模式已切換為' 'Organize mode switched to') + ': ' + $script:OrganizeMode)
                Draw-StatusBar
                Draw-SettingsPanel
                Draw-OneMenuItem -Index $script:SelectedMenu -Selected
                continue
            }
            'TogglePrimaryOnly' {
                Toggle-PrimaryOnly
                Update-Status -Status (T '完成' 'Done') -Summary ((T '只整理主檔已切換為' 'Primary only switched to') + ': ' + $script:OrganizePrimaryOnly)
                Draw-StatusBar
                Draw-SettingsPanel
                Draw-OneMenuItem -Index $script:SelectedMenu -Selected
                continue
            }
            'Exit' { return }
        }
    }
}
finally {
    [Console]::CursorVisible = $true
    Clear-Host
}
