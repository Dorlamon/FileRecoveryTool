# ================================
# Office Recovery Toolkit v5.8.5.3.3
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
$script:LastCsvReport = ''
$script:LastRenamePreviewCsv = ''
$script:LastRenamePreviewHtml = ''
$script:LastRenameLog = ''
$script:SettingsPath = Join-Path $PSScriptRoot 'OfficeRecoveryToolkit.settings.json'
$script:NamingMode = 'Smart'
$script:LastNamingMode = 'Smart'

# v5.4 settings
$script:OrganizeMode = 'Copy'   # Copy / Move
$script:OrganizePrimaryOnly = $true
$script:LegacyQuickMode = $true
$script:LegacyMaxReadKB = 512
$script:LegacyTextPreviewLength = 200
$script:LegacyOfficeFallback = $true
$script:LegacyOfficeTimeoutSec = 45
$script:LegacyOfficeMaxFileMB = 32
$script:LegacyConversionMode = $true
$script:LegacyKeepTempConverted = $false
$script:OfficeProbeEnabled = $true
$script:OfficeProbeTimeoutSec = 15
$script:OfficeProbePath = Join-Path $PSScriptRoot 'OfficeEncryptionProbe.exe'
$script:KillOfficeProcessesBeforeScan = $true
$script:KillOfficeProcessesOnTimeout = $true
$script:SelectedMenu = 0
$script:UiCache = @{
    WindowWidth = 0
    WindowHeight = 0
    SettingsLines = @()
    StatusLines = @()
}

$script:OfficeInterop = @{ Word = $null; Excel = $null; PowerPoint = $null; Initialized = $false }

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

function Reset-UiCache {
    $script:UiCache.WindowWidth = 0
    $script:UiCache.WindowHeight = 0
    $script:UiCache.SettingsLines = @()
    $script:UiCache.StatusLines = @()
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
    Reset-UiCache
}

function Update-Status {
    param(
        [string]$Status,
        [string]$Summary
    )
    $script:LastStatus = $Status
    $script:LastSummary = $Summary
}

function Save-AppState {
    try {
        $state = [ordered]@{
            Lang = $script:Lang
            ScanRoot = $script:ScanRoot
            OutputRoot = $script:OutputRoot
            OrganizeMode = $script:OrganizeMode
            OrganizePrimaryOnly = $script:OrganizePrimaryOnly
            LegacyQuickMode = $script:LegacyQuickMode
            LegacyMaxReadKB = $script:LegacyMaxReadKB
            LegacyTextPreviewLength = $script:LegacyTextPreviewLength
            LegacyOfficeFallback = $script:LegacyOfficeFallback
            LegacyOfficeTimeoutSec = $script:LegacyOfficeTimeoutSec
            LegacyOfficeMaxFileMB = $script:LegacyOfficeMaxFileMB
            LegacyConversionMode = $script:LegacyConversionMode
            LegacyKeepTempConverted = $script:LegacyKeepTempConverted
            OfficeProbeEnabled = $script:OfficeProbeEnabled
            OfficeProbeTimeoutSec = $script:OfficeProbeTimeoutSec
            OfficeProbePath = $script:OfficeProbePath
            KillOfficeProcessesBeforeScan = $script:KillOfficeProcessesBeforeScan
            KillOfficeProcessesOnTimeout = $script:KillOfficeProcessesOnTimeout
            LastHtmlReport = $script:LastHtmlReport
            LastOrganizerLog = $script:LastOrganizerLog
            LastCsvReport = $script:LastCsvReport
            LastRenamePreviewCsv = $script:LastRenamePreviewCsv
            LastRenamePreviewHtml = $script:LastRenamePreviewHtml
            LastRenameLog = $script:LastRenameLog
            NamingMode = $script:NamingMode
            LastStatus = $script:LastStatus
            LastSummary = $script:LastSummary
            SavedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        }
        $json = $state | ConvertTo-Json -Depth 4
        [IO.File]::WriteAllText($script:SettingsPath, $json, [Text.Encoding]::UTF8)
    }
    catch {}
}

function Load-AppState {
    if (-not (Test-Path -LiteralPath $script:SettingsPath)) { return }
    try {
        $cfg = Get-Content -LiteralPath $script:SettingsPath -Raw | ConvertFrom-Json
        if ($cfg.Lang) { $script:Lang = [string]$cfg.Lang }
        if ($cfg.ScanRoot) { $script:ScanRoot = [string]$cfg.ScanRoot }
        if ($cfg.OutputRoot) { $script:OutputRoot = [string]$cfg.OutputRoot }
        if ($cfg.OrganizeMode) { $script:OrganizeMode = [string]$cfg.OrganizeMode }
        if ($null -ne $cfg.OrganizePrimaryOnly) { $script:OrganizePrimaryOnly = [bool]$cfg.OrganizePrimaryOnly }
        if ($null -ne $cfg.LegacyQuickMode) { $script:LegacyQuickMode = [bool]$cfg.LegacyQuickMode }
        if ($cfg.LegacyMaxReadKB) { $script:LegacyMaxReadKB = [int]$cfg.LegacyMaxReadKB }
        if ($cfg.LegacyTextPreviewLength) { $script:LegacyTextPreviewLength = [int]$cfg.LegacyTextPreviewLength }
        if ($null -ne $cfg.LegacyOfficeFallback) { $script:LegacyOfficeFallback = [bool]$cfg.LegacyOfficeFallback }
        if ($cfg.LegacyOfficeTimeoutSec) { $script:LegacyOfficeTimeoutSec = [int]$cfg.LegacyOfficeTimeoutSec }
        if ($cfg.LegacyOfficeMaxFileMB) { $script:LegacyOfficeMaxFileMB = [int]$cfg.LegacyOfficeMaxFileMB }
        if ($null -ne $cfg.LegacyConversionMode) { $script:LegacyConversionMode = [bool]$cfg.LegacyConversionMode }
        if ($null -ne $cfg.LegacyKeepTempConverted) { $script:LegacyKeepTempConverted = [bool]$cfg.LegacyKeepTempConverted }
        if ($null -ne $cfg.OfficeProbeEnabled) { $script:OfficeProbeEnabled = [bool]$cfg.OfficeProbeEnabled }
        if ($cfg.OfficeProbeTimeoutSec) { $script:OfficeProbeTimeoutSec = [int]$cfg.OfficeProbeTimeoutSec }
        if ($cfg.OfficeProbePath) { $script:OfficeProbePath = [string]$cfg.OfficeProbePath }
        if ($null -ne $cfg.KillOfficeProcessesBeforeScan) { $script:KillOfficeProcessesBeforeScan = [bool]$cfg.KillOfficeProcessesBeforeScan }
        if ($null -ne $cfg.KillOfficeProcessesOnTimeout) { $script:KillOfficeProcessesOnTimeout = [bool]$cfg.KillOfficeProcessesOnTimeout }
        if ($cfg.LastHtmlReport) { $script:LastHtmlReport = [string]$cfg.LastHtmlReport }
        if ($cfg.LastOrganizerLog) { $script:LastOrganizerLog = [string]$cfg.LastOrganizerLog }
        if ($cfg.LastCsvReport) { $script:LastCsvReport = [string]$cfg.LastCsvReport }
        if ($cfg.LastRenamePreviewCsv) { $script:LastRenamePreviewCsv = [string]$cfg.LastRenamePreviewCsv }
        if ($cfg.LastRenamePreviewHtml) { $script:LastRenamePreviewHtml = [string]$cfg.LastRenamePreviewHtml }
        if ($cfg.LastRenameLog) { $script:LastRenameLog = [string]$cfg.LastRenameLog }
        if ($cfg.NamingMode) { $script:NamingMode = [string]$cfg.NamingMode }
        if ($cfg.LastStatus) { $script:LastStatus = [string]$cfg.LastStatus }
        if ($cfg.LastSummary) { $script:LastSummary = [string]$cfg.LastSummary }
    }
    catch {}
}


function Stop-OrphanOfficeProcesses {
    param(
        [switch]$Force
    )

    $names = @('WINWORD','EXCEL','POWERPNT','wordconv','excelcnv','ppcnvcom')
    foreach ($name in $names) {
        try {
            Get-Process -Name $name -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    if ($Force) {
                        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
                    }
                    else {
                        Stop-Process -Id $_.Id -ErrorAction SilentlyContinue
                    }
                }
                catch {}
            }
        }
        catch {}
    }
}

function Get-DefaultOfficeProbePath {
    $exe = Join-Path $PSScriptRoot 'OfficeEncryptionProbe.exe'
    if (Test-Path -LiteralPath $exe) { return $exe }
    return ''
}

function Get-OfficeProbePath {
    if (-not [string]::IsNullOrWhiteSpace($script:OfficeProbePath) -and (Test-Path -LiteralPath $script:OfficeProbePath)) {
        return $script:OfficeProbePath
    }
    return (Get-DefaultOfficeProbePath)
}

function Invoke-OfficeEncryptionProbe {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        [int]$TimeoutSec = 15
    )

    $fallback = [pscustomobject]@{
        State = 'Unavailable'
        Detail = 'Probe tool not found'
        ExitCode = 6
        CanOpenSafely = $false
        CanConvertSafely = $false
    }

    if (-not $script:OfficeProbeEnabled) { return $fallback }
    $tool = Get-OfficeProbePath
    if ([string]::IsNullOrWhiteSpace($tool) -or -not (Test-Path -LiteralPath $tool)) {
        return $fallback
    }

    $stdoutPath = Join-Path $env:TEMP ('ORT_PROBE_OUT_' + [guid]::NewGuid().ToString() + '.json')
    $stderrPath = Join-Path $env:TEMP ('ORT_PROBE_ERR_' + [guid]::NewGuid().ToString() + '.txt')

    try {
        $proc = Start-Process -FilePath $tool -ArgumentList @('--json', '--path', $Path) -WindowStyle Hidden -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
        $completed = $proc.WaitForExit($TimeoutSec * 1000)
        if (-not $completed) {
            try { $proc.Kill() } catch {}
            if ($script:KillOfficeProcessesOnTimeout) { Stop-OrphanOfficeProcesses -Force }
            return [pscustomobject]@{
                State = 'Error'
                Detail = 'Probe timeout'
                ExitCode = 6
                CanOpenSafely = $false
                CanConvertSafely = $false
            }
        }

        $stdout = ''
        $stderr = ''
        if (Test-Path -LiteralPath $stdoutPath) { try { $stdout = Get-Content -LiteralPath $stdoutPath -Raw -Encoding UTF8 } catch {} }
        if (Test-Path -LiteralPath $stderrPath) { try { $stderr = Get-Content -LiteralPath $stderrPath -Raw -Encoding UTF8 } catch {} }

        if (-not [string]::IsNullOrWhiteSpace($stdout)) {
            try {
                $obj = $stdout | ConvertFrom-Json -ErrorAction Stop
                if ($null -eq $obj.ExitCode) {
                    try { $obj | Add-Member -NotePropertyName ExitCode -NotePropertyValue ([int]$proc.ExitCode) -Force } catch {}
                }
                if ([string]::IsNullOrWhiteSpace([string]$obj.Detail) -and -not [string]::IsNullOrWhiteSpace($stderr)) {
                    try { $obj.Detail = $stderr.Trim() } catch {}
                }
                return $obj
            }
            catch {
            }
        }

        return [pscustomobject]@{
            State = 'Error'
            Detail = $(if (-not [string]::IsNullOrWhiteSpace($stderr)) { $stderr.Trim() } elseif (-not [string]::IsNullOrWhiteSpace($stdout)) { $stdout.Trim() } else { 'Probe returned invalid output' })
            ExitCode = [int]$(if ($proc) { $proc.ExitCode } else { 6 })
            CanOpenSafely = $false
            CanConvertSafely = $false
        }
    }
    catch {
        return [pscustomobject]@{
            State = 'Error'
            Detail = $_.Exception.Message
            ExitCode = 6
            CanOpenSafely = $false
            CanConvertSafely = $false
        }
        Show-ProgressLine -Current $total -Total $total -FileName '' -Force
    }
    finally {
        if (Test-Path -LiteralPath $stdoutPath) { try { Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue } catch {} }
        if (Test-Path -LiteralPath $stderrPath) { try { Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue } catch {} }
    }
}

function Test-ShouldSkipOfficeByProbe {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )

    $probe = Invoke-OfficeEncryptionProbe -Path $FilePath -TimeoutSec $script:OfficeProbeTimeoutSec
    if (-not $probe) { return $null }

    $state = [string]$probe.State
    if ($state -in @('Encrypted','WriteProtectedOnly','PossiblyProtected','Corrupt')) {
        return [pscustomobject]@{
            Skip = $true
            State = $state
            Detail = [string]$probe.Detail
            ExitCode = [int]$probe.ExitCode
            Probe = $probe
        }
    }

    if ($state -eq 'Error') {
        return [pscustomobject]@{
            Skip = $true
            State = 'ProbeError'
            Detail = if ([string]::IsNullOrWhiteSpace([string]$probe.Detail)) { 'Probe failed' } else { [string]$probe.Detail }
            ExitCode = [int]$probe.ExitCode
            Probe = $probe
        }
    }

    return [pscustomobject]@{
        Skip = $false
        State = $state
        Detail = [string]$probe.Detail
        ExitCode = [int]$probe.ExitCode
        Probe = $probe
    }
}


function Test-OoxmlProtectionState {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )

    $res = [ordered]@{
        Skip = $false
        State = 'NotProtected'
        Detail = ''
    }

    if (-not (Test-Path -LiteralPath $FilePath)) {
        $res.State = 'Missing'
        $res.Detail = 'File not found'
        return [pscustomobject]$res
    }

    try {
        $fs = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            if ($fs.Length -ge 8) {
                $sig = New-Object byte[] 8
                [void]$fs.Read($sig, 0, 8)
                $ole = @(0xD0,0xCF,0x11,0xE0,0xA1,0xB1,0x1A,0xE1)
                $isOle = $true
                for ($i = 0; $i -lt 8; $i++) {
                    if ($sig[$i] -ne $ole[$i]) { $isOle = $false; break }
                }
                if ($isOle) {
                    $res.Skip = $true
                    $res.State = 'Encrypted'
                    $res.Detail = 'OOXML password-protected package detected'
                    return [pscustomobject]$res
                }
            }
        }
        finally {
            $fs.Dispose()
        }
    }
    catch {
    }

    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue | Out-Null
        $zip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
        try {
            foreach ($entry in $zip.Entries) {
                if ($entry.FullName -eq 'EncryptionInfo' -or $entry.FullName -eq 'EncryptedPackage') {
                    $res.Skip = $true
                    $res.State = 'Encrypted'
                    $res.Detail = 'OOXML encryption markers detected'
                    return [pscustomobject]$res
                }
            }
        }
        finally {
            $zip.Dispose()
        }
    }
    catch {
    }

    return [pscustomobject]$res
}

function Test-ShouldSkipLegacyOfficeByProbe {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,
        [ValidateSet('Word','Excel','PowerPoint')][string]$AppType
    )
    return (Test-ShouldSkipOfficeByProbe -FilePath $FilePath)
}


function Restore-ResultsFromLastCsv {
    if ($script:Results -and @($script:Results).Count -gt 0) { return }

    if ([string]::IsNullOrWhiteSpace($script:LastCsvReport)) { return }
    if (-not (Test-Path -LiteralPath $script:LastCsvReport)) { return }

    try {
        $rows = Import-Csv -LiteralPath $script:LastCsvReport -Encoding UTF8
        if ($rows) {
            foreach ($row in $rows) {
                if ($row.SizeKB -ne $null -and $row.SizeKB -ne '') {
                    try { $row.SizeKB = [double]$row.SizeKB } catch {}
                }
                if ($row.RoleRank -ne $null -and $row.RoleRank -ne '') {
                    try { $row.RoleRank = [int]$row.RoleRank } catch {}
                }
            }
            $script:Results = @($rows)
        }
    }
    catch {
    }
}

function Save-StateAndStatus {
    param(
        [string]$Status = $null,
        [string]$Summary = $null
    )
    if ($null -ne $Status) { $script:LastStatus = $Status }
    if ($null -ne $Summary) { $script:LastSummary = $Summary }
    Save-AppState
}

function Get-SafeHtml {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    Add-Type -AssemblyName System.Web
    return [System.Web.HttpUtility]::HtmlEncode($Text)
}

function Get-FriendlyParseReason {
    param(
        [string]$Reason,
        [string]$ConversionStatus
    )

    $reasonText = [string]$Reason
    $convText = [string]$ConversionStatus

    if ([string]::IsNullOrWhiteSpace($reasonText)) {
        if (-not [string]::IsNullOrWhiteSpace($convText)) {
            return $convText
        }
        return ''
    }

    switch -Regex ($reasonText) {
        '^RTF$' {
            return (T 'RTF 文字已成功擷取' 'RTF text extracted successfully')
        }
        '^PDF text$' {
            return (T 'PDF 文字已成功擷取' 'PDF text extracted successfully')
        }
        '^Unsupported extension$' {
            return (T '不支援的副檔名' 'Unsupported extension')
        }
        'timeout|逾時' {
            if (-not [string]::IsNullOrWhiteSpace($convText)) {
                return ((T '轉換逾時' 'Conversion timeout') + ' / ' + $convText)
            }
            return (T '轉換逾時' 'Conversion timeout')
        }
        'PDF text could not be extracted; fell back to file hash|PDF 無法抽出文字，改用檔案雜湊' {
            return (T 'PDF 無法抽出文字，已改用檔案雜湊比對' 'PDF text could not be extracted; fell back to file hash matching')
        }
        'Converter tool not found|ConverterNotFound' {
            return (T '找不到 Office 相容性套件轉換工具' 'Office compatibility converter tool was not found')
        }
        'Converter produced no output file|ConverterFailed' {
            return (T '轉換工具執行完成，但沒有產出新版檔案' 'Converter finished but did not produce an output file')
        }
        'not installed|ActiveX component can''t create object|無法建立 ActiveX' {
            return (T '找不到對應的 Office 應用程式，無法進行轉換' 'The required Office application is not available for conversion')
        }
        'password|密碼' {
            return (T '檔案可能受密碼保護，無法自動轉換或解析' 'The file may be password-protected and could not be converted or parsed automatically')
        }
        default {
            if (-not [string]::IsNullOrWhiteSpace($convText) -and $convText -notmatch [regex]::Escape($reasonText)) {
                return ($reasonText + ' / ' + $convText)
            }
            return $reasonText
        }
    }
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
        '.rtf'  { return 'Word' }
        '.pptx' { return 'PowerPoint' }
        '.ppt'  { return 'PowerPoint' }
        '.pdf'  { return 'PDF' }
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
    Write-At 2 1  (T 'Office 檔案救援分析工具 v5.8.5.3 LightBar' 'Office Recovery Analyzer v5.8.5.3 LightBar') White DarkBlue
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
        @{ Key='C'; Text=(T '切換 Legacy 轉新版' 'Toggle Legacy Convert'); Action='ToggleLegacyConversion' },
        @{ Key='H'; Text=(T '切換智慧命名模式' 'Toggle Naming Mode'); Action='ToggleNamingMode' },
        @{ Key='L'; Text=(T '切換語系' 'Switch Language'); Action='ToggleLang' },
        @{ Key='Esc'; Text=(T '離開' 'Exit'); Action='Exit' }
    )
}

function Get-MenuLayout {
    $windowWidth = [Console]::WindowWidth
    $left = 4
    $gap = 3
    $menuWidth = 38

    if ($windowWidth -lt 110) { $menuWidth = 34 }
    if ($windowWidth -lt 96)  { $menuWidth = 30 }

    $rightLeft = $left + $menuWidth + $gap
    $rightWidth = [Math]::Max(($windowWidth - $rightLeft - 2), 18)

    return [PSCustomObject]@{
        Top        = 6
        Left       = $left
        Width      = $menuWidth
        RightLeft  = $rightLeft
        RightWidth = $rightWidth
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
    param([switch]$Force)

    $layout = Get-MenuLayout
    $rightLeft = $layout.RightLeft
    $rightWidth = $layout.RightWidth
    $valueWidth = [Math]::Max($rightWidth - 14, 6)

    $newLines = @(
        (T '目前設定' 'Current Settings'),
        ((T '掃描路徑' 'Scan') + ' : ' + (Get-ShortDisplayText $script:ScanRoot $valueWidth)),
        ((T '輸出資料夾' 'Output') + ' : ' + (Get-ShortDisplayText $script:OutputRoot $valueWidth)),
        ((T '最新 CSV' 'CSV') + ' : ' + (Get-ShortDisplayText $script:LastCsvReport $valueWidth)),
        ((T '最新 HTML' 'HTML') + ' : ' + (Get-ShortDisplayText $script:LastHtmlReport $valueWidth)),
        ((T '模擬改名 CSV' 'Preview CSV') + ' : ' + (Get-ShortDisplayText $script:LastRenamePreviewCsv $valueWidth)),
        ((T '模擬改名 HTML' 'Preview HTML') + ' : ' + (Get-ShortDisplayText $script:LastRenamePreviewHtml $valueWidth)),
        ((T 'Legacy 轉新版' 'Legacy Convert') + ' : ' + ([string]$script:LegacyConversionMode)),
        ((T '保留暫存轉檔' 'Keep Temp') + ' : ' + ([string]$script:LegacyKeepTempConverted)),
        ((T '最新改名紀錄' 'Rename Log') + ' : ' + (Get-ShortDisplayText $script:LastRenameLog $valueWidth)),
        ((T '最新整理紀錄' 'Organize Log') + ' : ' + (Get-ShortDisplayText $script:LastOrganizerLog $valueWidth)),
        ((T '命名模式' 'Naming Mode') + ' : ' + (Get-NamingModeDisplay)),
        ((T '舊檔快速模式' 'Legacy Quick') + ' : ' + ([string]$script:LegacyQuickMode)),
        ((T 'Office 背景解析' 'Office Fallback') + ' : ' + ([string]$script:LegacyOfficeFallback)),
        ((T 'Office 逾時秒數' 'Timeout Sec') + ' : ' + ([string]$script:LegacyOfficeTimeoutSec))
    )

    $oldLines = @($script:UiCache.SettingsLines)
    $max = [Math]::Max($oldLines.Count, $newLines.Count)

    for ($i = 0; $i -lt $max; $i++) {
        $newLine = if ($i -lt $newLines.Count) { [string]$newLines[$i] } else { '' }
        $oldLine = if ($i -lt $oldLines.Count) { [string]$oldLines[$i] } else { '' }

        if ($Force -or $newLine -ne $oldLine) {
            $row = 5 + $i
            $fg = [ConsoleColor]::Gray
            if ($i -eq 0) { $fg = [ConsoleColor]::Yellow }
            Write-At $rightLeft $row $newLine $fg Black -FixedWidth $rightWidth -Ellipsis
        }
    }

    $script:UiCache.SettingsLines = $newLines
}

function Draw-LightBarMenu {
    param([switch]$Force)

    $menu = Get-MenuItems
    $layout = Get-MenuLayout

    Write-At $layout.Left 5 (T '主選單' 'Main Menu') Yellow Black -FixedWidth $layout.Width

    for ($i = 0; $i -lt $menu.Count; $i++) {
        Draw-OneMenuItem -Index $i -Selected:($i -eq $script:SelectedMenu)
    }

    Draw-SettingsPanel -Force:$Force
}

function Draw-StatusBar {
    param([switch]$Force)

    $w = [Console]::WindowWidth
    $h = [Console]::WindowHeight
    $top = $h - 4
    $resultCount = Get-ResultCount

    $newLines = @(
        ('{0}: {1} | {2}: {3} | {4}: {5}' -f (T '模式' 'Mode'), $script:OrganizeMode, (T '只整理主檔' 'Primary Only'), $script:OrganizePrimaryOnly, (T '語系' 'Language'), $script:Lang),
        ('{0}: {1} | {2}: {3}' -f (T '結果筆數' 'Result Count'), $resultCount, (T '最新狀態' 'Last Status'), (Get-ShortDisplayText $script:LastStatus 45)),
        ('{0}: {1}' -f (T '摘要' 'Summary'), (Get-ShortDisplayText $script:LastSummary 100)),
        (T '熱鍵：↑↓ 選擇 / Enter 執行 / 數字快速鍵 / C Legacy轉檔 / H 命名模式 / L 語系 / Esc 離開' 'Hotkeys: ↑↓ select / Enter run / numbers / C legacy convert / H naming mode / L language / Esc exit')
    )

    $oldLines = @($script:UiCache.StatusLines)

    for ($i = 0; $i -lt 4; $i++) {
        $line = if ($i -lt $newLines.Count) { [string]$newLines[$i] } else { '' }
        $oldLine = if ($i -lt $oldLines.Count) { [string]$oldLines[$i] } else { '' }

        if ($Force -or $line -ne $oldLine) {
            Write-At 0 ($top + $i) (' ' * ($w - 1)) Black DarkCyan -NoPad
            Write-At 1 ($top + $i) $line Black DarkCyan -FixedWidth ($w - 2) -Ellipsis
        }
    }

    $script:UiCache.StatusLines = $newLines
}

# -----------------------------
# UI
# -----------------------------
function Draw-UI {
    $currentWidth = [Console]::WindowWidth
    $currentHeight = [Console]::WindowHeight
    $force = $false

    if ($script:ForceFullRedraw) {
        $script:UiCache.WindowWidth = 0
        $script:UiCache.WindowHeight = 0
        $script:UiCache.SettingsLines = @()
        $script:UiCache.StatusLines = @()
        $force = $true
        $script:ForceFullRedraw = $false
    }

    if ($script:UiCache.WindowWidth -ne $currentWidth -or $script:UiCache.WindowHeight -ne $currentHeight) {
        $script:UiCache.WindowWidth = $currentWidth
        $script:UiCache.WindowHeight = $currentHeight
        $script:UiCache.SettingsLines = @()
        $script:UiCache.StatusLines = @()
        $force = $true
    }

    $ok = Draw-Frame
    if (-not $ok) { return }
    Draw-LightBarMenu -Force:$force
    Draw-StatusBar -Force:$force
}

$script:UseNativeProgressBar = $false
$script:ProgressRefreshMs = 180
$script:LastProgressRenderAt = [datetime]::MinValue
$script:ProgressUiActive = $false
$script:ProgressBaseRow = -1
$script:ProgressLastLine = ''
$script:ProgressCursorHidden = $false
$script:ScanCancelRequested = $false
$script:LastEscPollAt = [datetime]::MinValue
$script:EscPollIntervalMs = 80

function Start-ScanProgressUi {
    if ($script:ProgressUiActive) { return }

    $script:LastProgressRenderAt = [datetime]::MinValue
    $script:ProgressLastLine = ''
    $script:ScanCancelRequested = $false
    $script:LastEscPollAt = [datetime]::MinValue

    try {
        [Console]::CursorVisible = $false
        $script:ProgressCursorHidden = $true
    }
    catch {
        $script:ProgressCursorHidden = $false
    }

    try {
        $script:ProgressBaseRow = [Console]::CursorTop
    }
    catch {
        $script:ProgressBaseRow = -1
    }

    Write-Host ''
    $script:ProgressUiActive = $true
}

function Stop-ScanProgressUi {
    if (-not $script:ProgressUiActive) { return }

    try {
        if ($script:ProgressBaseRow -ge 0) {
            [Console]::SetCursorPosition(0, $script:ProgressBaseRow)
            $blankWidth = [Math]::Max(1, [Console]::BufferWidth - 1)
            [Console]::Write((' ' * $blankWidth))
            [Console]::SetCursorPosition(0, $script:ProgressBaseRow)
        }
    }
    catch {}

    if ($script:UseNativeProgressBar) {
        Write-Progress -Activity (T '掃描中' 'Scanning') -Completed
    }

    try {
        if ($script:ProgressCursorHidden) {
            [Console]::CursorVisible = $true
        }
    }
    catch {}

    $script:ProgressUiActive = $false
    $script:ProgressBaseRow = -1
    $script:ProgressLastLine = ''
    $script:ProgressCursorHidden = $false
    $script:LastEscPollAt = [datetime]::MinValue
}


function Test-ScanCancelRequested {
    $now = Get-Date
    if ((($now - $script:LastEscPollAt).TotalMilliseconds) -lt $script:EscPollIntervalMs) {
        return $script:ScanCancelRequested
    }
    $script:LastEscPollAt = $now

    try {
        while ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq [ConsoleKey]::Escape) {
                $script:ScanCancelRequested = $true
                return $true
            }
        }
    }
    catch {
    }

    return $script:ScanCancelRequested
}


function Show-ScanCancelConfirmUi {
    param(
        [string]$PromptText
    )

    try {
        $raw = $Host.UI.RawUI
        $origVisible = $raw.CursorSize
    } catch {
        $origVisible = $null
    }

    $yesSelected = $true

    while ($true) {
        try {
            $width = [Math]::Max(40, [Console]::WindowWidth)
            $left = 0
            $top = [Math]::Max(0, [Console]::CursorTop)
            if ($script:ProgressUiActive -and $script:ProgressBaseRow -ge 0) {
                $top = [Math]::Min([Console]::BufferHeight - 2, $script:ProgressBaseRow + 1)
            }

            [Console]::SetCursorPosition($left, $top)
            $line1 = $PromptText
            if ($line1.Length -gt ($width - 1)) { $line1 = $line1.Substring(0, $width - 1) }
            $line1 = $line1.PadRight($width - 1)
            Write-Host $line1 -NoNewline -ForegroundColor Yellow

            [Console]::SetCursorPosition($left, [Math]::Min([Console]::BufferHeight - 1, $top + 1))
            if ($yesSelected) {
                $line2 = ('[ {0} ]    {1}' -f (T '是 / Y' 'Yes / Y'), (T '否 / N' 'No / N'))
            } else {
                $line2 = ('{0}    [ {1} ]' -f (T '是 / Y' 'Yes / Y'), (T '否 / N' 'No / N'))
            }
            if ($line2.Length -gt ($width - 1)) { $line2 = $line2.Substring(0, $width - 1) }
            $line2 = $line2.PadRight($width - 1)
            Write-Host $line2 -NoNewline -ForegroundColor Cyan
        } catch {
            Write-Host ''
            Write-Host $PromptText -ForegroundColor Yellow
            Write-Host (T '按 Y 確認中止，按 N 繼續掃描。' 'Press Y to cancel, N to continue.') -ForegroundColor Cyan
        }

        $key = [Console]::ReadKey($true)

        switch ($key.Key) {
            'LeftArrow' { $yesSelected = $true; continue }
            'RightArrow' { $yesSelected = $false; continue }
            'Y' { return $true }
            'N' { return $false }
            'Escape' { return $false }
            'Enter' { return $yesSelected }
            default { continue }
        }
    }
}

function Clear-ScanCancelConfirmUi {
    try {
        if ($script:ProgressUiActive -and $script:ProgressBaseRow -ge 0) {
            $width = [Math]::Max(40, [Console]::WindowWidth)
            $left = 0
            $top = [Math]::Min([Console]::BufferHeight - 2, $script:ProgressBaseRow + 1)
            [Console]::SetCursorPosition($left, $top)
            Write-Host (' ' * ($width - 1)) -NoNewline
            [Console]::SetCursorPosition($left, [Math]::Min([Console]::BufferHeight - 1, $top + 1))
            Write-Host (' ' * ($width - 1)) -NoNewline
            [Console]::SetCursorPosition($left, $script:ProgressBaseRow)
        }
    } catch {
    }
}

function Show-ProgressLine {
    param(
        [int]$Current,
        [int]$Total,
        [string]$FileName,
        [switch]$Force
    )

    if (-not $script:ProgressUiActive) {
        Start-ScanProgressUi
    }

    $now = Get-Date
    if (-not $Force) {
        $elapsed = ($now - $script:LastProgressRenderAt).TotalMilliseconds
        if ($elapsed -lt $script:ProgressRefreshMs -and $Current -lt $Total) {
            return
        }
    }
    $script:LastProgressRenderAt = $now

    $barWidth = 32
    if ($Total -gt 0) {
        $pct = [Math]::Floor(($Current * 100) / $Total)
    } else {
        $pct = 0
    }

    $filled = [Math]::Floor(($pct * $barWidth) / 100)
    $empty = $barWidth - $filled
    $bar = ('█' * $filled) + ('░' * $empty)

    $safeName = if ([string]::IsNullOrWhiteSpace($FileName)) { '' } else { $FileName }
    if ($safeName.Length -gt 58) {
        $safeName = $safeName.Substring(0, 55) + '...'
    }

    $line = ('{0} [{1}] {2,3}%  ({3}/{4})  {5}   {6}' -f (T '掃描中' 'Scanning'), $bar, $pct, $Current, $Total, $safeName, (T '按 ESC 可中止掃描' 'Press ESC to cancel scan'))

    if ($script:UseNativeProgressBar) {
        Write-Progress -Activity (T '掃描中' 'Scanning') -Status $safeName -PercentComplete $pct
    }

    try {
        $width = [Console]::BufferWidth
        if ($width -lt 20) { $width = 120 }
    }
    catch {
        $width = 120
    }

    if ($line.Length -gt ($width - 1)) {
        $line = $line.Substring(0, [Math]::Max(1, $width - 1))
    }

    $padLen = [Math]::Max(0, $width - $line.Length - 1)
    $render = $line + (' ' * $padLen)

    if (-not $Force -and $render -eq $script:ProgressLastLine -and $Current -lt $Total) {
        return
    }

    try {
        if ($script:ProgressBaseRow -ge 0) {
            [Console]::SetCursorPosition(0, $script:ProgressBaseRow)
            [Console]::Write($render)
        }
        else {
            [Console]::Write("`r$render")
        }
    }
    catch {
        [Console]::Write("`r$render")
    }

    $script:ProgressLastLine = $render
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


function Normalize-PlainText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

    $t = $Text
    $t = $t -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ' '
    $t = $t -replace '\s+', ' '
    return $t.Trim()
}

function Get-RtfPlainText {
    param([string]$FilePath)

    if (-not (Test-Path -LiteralPath $FilePath)) { return '' }

    $rtf = ''
    try {
        $sr = New-Object System.IO.StreamReader($FilePath, $true)
        try {
            $rtf = $sr.ReadToEnd()
        }
        finally {
            $sr.Close()
        }
    }
    catch {
        try {
            $rtf = [System.Text.Encoding]::Default.GetString([System.IO.File]::ReadAllBytes($FilePath))
        }
        catch {
            $rtf = ''
        }
    }

    if ([string]::IsNullOrWhiteSpace($rtf)) { return '' }

    try {
        Add-Type -AssemblyName System.Windows.Forms
        $box = New-Object System.Windows.Forms.RichTextBox
        try {
            $box.Rtf = $rtf
            return (Normalize-PlainText $box.Text)
        }
        finally {
            $box.Dispose()
        }
    }
    catch {
        $t = $rtf
        $t = $t -replace '\\par[d]?', ' '
        $t = $t -replace '\\tab', ' '
        $t = $t -replace '\\u-?\d+\??', ' '
        $t = $t -replace "\\'[0-9a-fA-F]{2}", ' '
        $t = $t -replace '\\[a-zA-Z]+\d* ?', ' '
        $t = $t -replace '[{}]', ' '
        return (Normalize-PlainText $t)
    }
}


function Get-RtfTextBestEffort {
    param([string]$FilePath)
    return (Get-RtfPlainText -FilePath $FilePath)
}

function Expand-PdfFlateBytes {
    param([byte[]]$Bytes)

    if (-not $Bytes -or $Bytes.Length -lt 6) { return '' }

    $candidates = @(
        @{ Start = 0; EndTrim = 0 },
        @{ Start = 2; EndTrim = 0 },
        @{ Start = 0; EndTrim = 1 },
        @{ Start = 2; EndTrim = 1 }
    )

    foreach ($c in $candidates) {
        try {
            $start = [int]$c.Start
            $len = $Bytes.Length - $start - [int]$c.EndTrim
            if ($len -le 4) { continue }
            $slice = New-Object byte[] $len
            [Array]::Copy($Bytes, $start, $slice, 0, $len)

            $ms = New-Object IO.MemoryStream(,$slice)
            try {
                $ds = New-Object IO.Compression.DeflateStream($ms, [IO.Compression.CompressionMode]::Decompress)
                try {
                    $out = New-Object IO.MemoryStream
                    try {
                        $buffer = New-Object byte[] 4096
                        while (($read = $ds.Read($buffer, 0, $buffer.Length)) -gt 0) {
                            $out.Write($buffer, 0, $read)
                        }
                        $raw = $out.ToArray()
                        if ($raw -and $raw.Length -gt 0) {
                            return [Text.Encoding]::GetEncoding('ISO-8859-1').GetString($raw)
                        }
                    }
                    finally {
                        $out.Dispose()
                    }
                }
                finally {
                    $ds.Dispose()
                }
            }
            finally {
                $ms.Dispose()
            }
        }
        catch {
        }
    }

    return ''
}

function Get-PdfTextBestEffort {
    param(
        [string]$FilePath,
        [int]$MaxChars = 6000
    )

    if (-not (Test-Path -LiteralPath $FilePath)) { return '' }

    try {
        $bytes = [IO.File]::ReadAllBytes($FilePath)
    }
    catch {
        return ''
    }

    if (-not $bytes -or $bytes.Length -eq 0) { return '' }

    $latin1 = [Text.Encoding]::GetEncoding('ISO-8859-1').GetString($bytes)
    $parts = New-Object System.Collections.ArrayList

    try {
        $metaMatches = [regex]::Matches($latin1, '/(Title|Author|Subject|Keywords)\s*\((.*?)\)', 'Singleline')
        foreach ($m in $metaMatches) {
            $v = Normalize-PlainText $m.Groups[2].Value
            if ($v.Length -ge 2) { [void]$parts.Add($v) }
        }
    }
    catch {
    }

    try {
        $streamMatches = [regex]::Matches($latin1, '(?s)(<<.*?>>)\s*stream\r?\n(.*?)\r?\nendstream')
        foreach ($m in $streamMatches) {
            $dict = [string]$m.Groups[1].Value
            $streamData = [string]$m.Groups[2].Value
            $streamText = ''

            if ($dict -match '/FlateDecode') {
                try {
                    $rawBytes = [Text.Encoding]::GetEncoding('ISO-8859-1').GetBytes($streamData)
                    $streamText = Expand-PdfFlateBytes -Bytes $rawBytes
                }
                catch {
                    $streamText = ''
                }
            }
            else {
                $streamText = $streamData
            }

            if ($streamText) {
                $textMatches = [regex]::Matches($streamText, '\((?:\\.|[^\\)]){2,}\)')
                foreach ($tm in $textMatches) {
                    $s = $tm.Value
                    if ($s.Length -ge 2) {
                        $s = $s.Substring(1, $s.Length - 2)
                        $s = $s -replace '\\\(', '('
                        $s = $s -replace '\\\)', ')'
                        $s = $s -replace '\\n', ' '
                        $s = $s -replace '\\r', ' '
                        $s = $s -replace '\\t', ' '
                        $s = $s -replace '\\\\', '\'
                        $s = Normalize-PlainText $s
                        if ($s.Length -ge 2) { [void]$parts.Add($s) }
                    }
                }
            }

            if ((($parts -join ' ').Length) -ge $MaxChars) { break }
        }
    }
    catch {
    }

    if ($parts.Count -eq 0) {
        try {
            $fallback = Get-LegacyReadableText -FilePath $FilePath -MaxReadKB 1024
            if ($fallback) { [void]$parts.Add($fallback) }
        }
        catch {
        }
    }

    $text = Normalize-PlainText ($parts -join ' ')
    if ($text.Length -gt $MaxChars) {
        $text = $text.Substring(0, $MaxChars)
    }

    return $text
}

function Get-ExcelSmartName {
    param([string]$FilePath)

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = $null
    $sheetNames = @()
    $cellText = ''
    $rosterProbeText = ''

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

                if (-not [string]::IsNullOrWhiteSpace($plain)) {
                    if ($rosterProbeText.Length -lt 1200) {
                        $need = 1200 - $rosterProbeText.Length
                        $rosterProbeText += ' ' + $plain.Substring(0, [Math]::Min($need, $plain.Length))
                    }

                    if ($plain.Length -gt 5 -and [string]::IsNullOrWhiteSpace($cellText)) {
                        $cellText = $plain.Substring(0, [Math]::Min(50, $plain.Length))
                    }
                }
            }
            finally { $sr2.Close() }
        }

        $combinedProbe = (($sheetNames -join ' ') + ' ' + $rosterProbeText).Trim()
        $rosterName = Get-WorkbookRosterName -Text $combinedProbe -Extension '.xlsx'
        if (-not [string]::IsNullOrWhiteSpace($rosterName)) {
            return (Get-SafeFileName $rosterName)
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


function Test-LegacyTextLooksUseful {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }

    $t = ($Text -replace '\s+', ' ').Trim()
    if ($t.Length -lt 12) { return $false }

    $questionCount = ([regex]::Matches($t, '\?')).Count
    if ($t.Length -gt 0) {
        $ratio = $questionCount / [double]$t.Length
        if ($ratio -gt 0.30) { return $false }
    }

    if ($t -notmatch '[\u4E00-\u9FFFA-Za-z0-9]{3,}') { return $false }
    return $true
}

function Get-LegacyOfficeTextViaHelper {
    param(
        [string]$FilePath,
        [ValidateSet('Word','Excel','PowerPoint')][string]$AppType,
        [int]$TimeoutSec = 8,
        [int]$MaxChars = 4000
    )

    if (-not (Test-Path -LiteralPath $FilePath)) { return '' }

    $helperOut = Join-Path $env:TEMP ('ORT_Helper_' + [guid]::NewGuid().ToString() + '.txt')
    $escapedPath = $FilePath.Replace("'", "''")
    $escapedOut  = $helperOut.Replace("'", "''")
    $escapedApp  = $AppType.Replace("'", "''")

    $helper = @"
`$ErrorActionPreference = 'SilentlyContinue'
`$filePath = '$escapedPath'
`$outPath  = '$escapedOut'
`$appType  = '$escapedApp'
`$maxChars = $MaxChars
`$result = ''

function Normalize-OfficeTextLocal([string]`$text) {
    if ([string]::IsNullOrWhiteSpace(`$text)) { return '' }
    `$t = `$text -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ' '
    `$t = `$t -replace '\s+', ' '
    `$t = `$t.Trim()
    if (`$t.Length -gt `$maxChars) { `$t = `$t.Substring(0, `$maxChars) }
    return `$t
}

if (`$appType -eq 'Excel') {
    `$excel = `$null
    `$wb = `$null
    try {
        `$excel = New-Object -ComObject Excel.Application
        `$excel.Visible = `$false
        `$excel.DisplayAlerts = `$false
        `$excel.ScreenUpdating = `$false
        `$excel.EnableEvents = `$false
        `$wb = `$excel.Workbooks.Open(`$filePath, 0, `$false)
        `$parts = New-Object System.Collections.ArrayList
        foreach (`$ws in `$wb.Worksheets) {
            try {
                if (`$ws.Name) { [void]`$parts.Add([string]`$ws.Name) }
                `$used = `$ws.UsedRange
                if (`$used) {
                    `$vals = `$used.Value2
                    if (`$vals -is [System.Array]) {
                        foreach (`$item in `$vals) {
                            if (`$null -ne `$item) { [void]`$parts.Add([string]`$item) }
                            if (((`$parts -join ' ').Length) -gt `$maxChars) { break }
                        }
                    } elseif (`$used.Text) {
                        [void]`$parts.Add([string]`$used.Text)
                    }
                }
            } catch {}
            if (((`$parts -join ' ').Length) -gt `$maxChars) { break }
        }
        `$result = Normalize-OfficeTextLocal (`$parts -join ' ')
    } catch {
        `$result = ''
    } finally {
        if (`$wb) { try { `$wb.Close(`$false) } catch {} ; try { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject(`$wb) } catch {} }
        if (`$excel) { try { `$excel.Quit() } catch {} ; try { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject(`$excel) } catch {} }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}
elseif (`$appType -eq 'Word') {
    `$word = `$null
    `$doc = `$null
    try {
        `$word = New-Object -ComObject Word.Application
        `$word.Visible = `$false
        `$word.DisplayAlerts = 0
        `$doc = `$word.Documents.Open(`$filePath, `$false, `$false)
        `$parts = @()
        try {
            if (`$doc.BuiltInDocumentProperties('Title').Value) { `$parts += [string]`$doc.BuiltInDocumentProperties('Title').Value }
        } catch {}
        try { if (`$doc.Paragraphs.Count -gt 0) { `$parts += [string]`$doc.Paragraphs.Item(1).Range.Text } } catch {}
        try { `$parts += [string]`$doc.Content.Text } catch {}
        `$result = Normalize-OfficeTextLocal (`$parts -join ' ')
    } catch {
        `$result = ''
    } finally {
        if (`$doc) { try { `$doc.Close() } catch {} ; try { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject(`$doc) } catch {} }
        if (`$word) { try { `$word.Quit() } catch {} ; try { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject(`$word) } catch {} }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}
elseif (`$appType -eq 'PowerPoint') {
    `$ppt = `$null
    `$pres = `$null
    try {
        `$ppt = New-Object -ComObject PowerPoint.Application
        `$ppt.Visible = 1
        `$pres = `$ppt.Presentations.Open(`$filePath, `$false, `$true, `$false)
        `$parts = New-Object System.Collections.ArrayList
        foreach (`$slide in `$pres.Slides) {
            try {
                foreach (`$shape in `$slide.Shapes) {
                    try {
                        if (`$shape.HasTextFrame -and `$shape.TextFrame.HasText) {
                            `$txt = [string]`$shape.TextFrame.TextRange.Text
                            if (-not [string]::IsNullOrWhiteSpace(`$txt)) { [void]`$parts.Add(`$txt) }
                        }
                    } catch {}
                    if (((`$parts -join ' ').Length) -gt `$maxChars) { break }
                }
            } catch {}
            if (((`$parts -join ' ').Length) -gt `$maxChars) { break }
        }
        `$result = Normalize-OfficeTextLocal (`$parts -join ' ')
    } catch {
        `$result = ''
    } finally {
        if (`$pres) { try { `$pres.Close() } catch {} ; try { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject(`$pres) } catch {} }
        if (`$ppt) { try { `$ppt.Quit() } catch {} ; try { [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject(`$ppt) } catch {} }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

if (`$result) {
    [System.IO.File]::WriteAllText(`$outPath, `$result, [System.Text.Encoding]::UTF8)
}
"@

    $helperPs1 = Join-Path $env:TEMP ('ORT_Helper_' + [guid]::NewGuid().ToString() + '.ps1')
    try {
        [System.IO.File]::WriteAllText($helperPs1, $helper, [System.Text.Encoding]::UTF8)
        $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File', $helperPs1) -WindowStyle Minimized -PassThru
        $completed = $proc.WaitForExit($TimeoutSec * 1000)
        if (-not $completed) {
            try { $proc.Kill() } catch {}
            if ($script:KillOfficeProcessesOnTimeout) { Stop-OrphanOfficeProcesses -Force }
            return ''
        }

        if (Test-Path -LiteralPath $helperOut) {
            try {
                $txt = Get-Content -LiteralPath $helperOut -Raw -Encoding UTF8
                return (($txt -replace '\s+', ' ').Trim())
            }
            catch { return '' }
            finally { try { Remove-Item -LiteralPath $helperOut -Force -ErrorAction SilentlyContinue } catch {} }
        }
    }
    catch {
        return ''
    }
    finally {
        if (Test-Path -LiteralPath $helperOut) { try { Remove-Item -LiteralPath $helperOut -Force -ErrorAction SilentlyContinue } catch {} }
        if (Test-Path -LiteralPath $helperPs1) { try { Remove-Item -LiteralPath $helperPs1 -Force -ErrorAction SilentlyContinue } catch {} }
    }

    return ''
}




function Release-ComObjectSafe {
    param($Object)
    if ($null -ne $Object) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object) } catch {}
    }
}


function Get-OfficeConverterPaths {
    $candidates = @(
        'C:\Program Files (x86)\Microsoft Office\Office12',
        'C:\Program Files\Microsoft Office\Office12'
    )

    foreach ($base in $candidates) {
        if (Test-Path -LiteralPath $base) {
            return @{
                Word  = Join-Path $base 'wordconv.exe'
                Excel = Join-Path $base 'excelcnv.exe'
                PowerPoint = Join-Path $base 'ppcnvcom.exe'
            }
        }
    }

    return @{
        Word  = ''
        Excel = ''
        PowerPoint = ''
    }
}

function Convert-LegacyOfficeViaConverterToTemp {
    param(
        [string]$FilePath,
        [ValidateSet('Word','Excel','PowerPoint')][string]$AppType,
        [int]$TimeoutSec = 45
    )

    $result = [ordered]@{
        Success = $false
        TempPath = ''
        ConvertedType = ''
        Status = 'Failed'
        Error = ''
        Method = 'Converter'
    }

    if (-not (Test-Path -LiteralPath $FilePath)) {
        $result.Error = 'File not found'
        return New-Object PSObject -Property $result
    }

    $conv = Get-OfficeConverterPaths
    $exe = $conv[$AppType]
    if ([string]::IsNullOrWhiteSpace($exe) -or -not (Test-Path -LiteralPath $exe)) {
        $result.Status = 'ConverterNotFound'
        $result.Error = 'Converter tool not found'
        return New-Object PSObject -Property $result
    }

    $targetExt = ''
    switch ($AppType) {
        'Word' {
            $targetExt = '.docx'
            $result.ConvertedType = 'docx'
        }
        'Excel' {
            $targetExt = '.xlsx'
            $result.ConvertedType = 'xlsx'
        }
        'PowerPoint' {
            $targetExt = '.pptx'
            $result.ConvertedType = 'pptx'
        }
    }

    $tempPath = Join-Path $env:TEMP ('ORT_CONVERT_' + [guid]::NewGuid().ToString() + $targetExt)
    $stdoutPath = Join-Path $env:TEMP ('ORT_CONVERT_STDOUT_' + [guid]::NewGuid().ToString() + '.log')
    $stderrPath = Join-Path $env:TEMP ('ORT_CONVERT_STDERR_' + [guid]::NewGuid().ToString() + '.log')

    try {
        $proc = Start-Process -FilePath $exe -ArgumentList @('-oice', '-nme', $FilePath, $tempPath) -WindowStyle Minimized -PassThru -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
        $completed = $proc.WaitForExit($TimeoutSec * 1000)
        if (-not $completed) {
            try { $proc.Kill() } catch {}
            if ($script:KillOfficeProcessesOnTimeout) { Stop-OrphanOfficeProcesses -Force }
            $result.Status = 'Timeout'
            $result.Error = 'Conversion timeout'
            return New-Object PSObject -Property $result
        }

        if (Test-Path -LiteralPath $tempPath) {
            $result.Success = $true
            $result.TempPath = $tempPath
            $result.Status = 'ConvertedByConverter'
            return New-Object PSObject -Property $result
        }

        $errText = ''
        try { if (Test-Path -LiteralPath $stderrPath) { $errText = (Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue) } } catch {}
        if ([string]::IsNullOrWhiteSpace($errText)) {
            try { if (Test-Path -LiteralPath $stdoutPath) { $errText = (Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue) } } catch {}
        }
        if ([string]::IsNullOrWhiteSpace($errText)) {
            $errText = 'Converter produced no output file'
        }

        $result.Status = 'ConverterFailed'
        $result.Error = ($errText -replace '\s+', ' ').Trim()
    }
    catch {
        $result.Status = 'ConverterFailed'
        $result.Error = $_.Exception.Message
    }
    finally {
        if ((-not $script:LegacyKeepTempConverted) -and (-not $result.Success) -and (Test-Path -LiteralPath $tempPath)) {
            try { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue } catch {}
        }
        if (Test-Path -LiteralPath $stdoutPath) { try { Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue } catch {} }
        if (Test-Path -LiteralPath $stderrPath) { try { Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue } catch {} }
    }

    return New-Object PSObject -Property $result
}

function New-OfficeInteropApp {
    param(
        [ValidateSet('Word','Excel','PowerPoint')][string]$AppType
    )

    try {
        switch ($AppType) {
            'Word' {
                $word = New-Object -ComObject Word.Application
                $word.Visible = $false
                $word.DisplayAlerts = 0
                return $word
            }
            'Excel' {
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
                try { $excel.ScreenUpdating = $false } catch {}
                try { $excel.EnableEvents = $false } catch {}
                return $excel
            }
            'PowerPoint' {
                $ppt = New-Object -ComObject PowerPoint.Application
                try { $ppt.Visible = 0 } catch {}
                return $ppt
            }
        }
    } catch {
        return $null
    }

    return $null
}

function Ensure-OfficeInteropApp {
    param(
        [ValidateSet('Word','Excel','PowerPoint')][string]$AppType
    )

    if (-not $script:OfficeInterop) {
        $script:OfficeInterop = @{ Word = $null; Excel = $null; PowerPoint = $null; Initialized = $false }
    }

    $existing = $script:OfficeInterop[$AppType]
    if ($existing) { return $existing }

    $app = New-OfficeInteropApp -AppType $AppType
    if (-not $app) {
        Start-Sleep -Milliseconds 300
        $app = New-OfficeInteropApp -AppType $AppType
    }

    $script:OfficeInterop[$AppType] = $app
    $script:OfficeInterop.Initialized = $true
    return $app
}

function Initialize-OfficeInterop {
    # 預先暖機，但不把單一失敗視為整體失敗
    [void](Ensure-OfficeInteropApp -AppType Word)
    [void](Ensure-OfficeInteropApp -AppType Excel)
    [void](Ensure-OfficeInteropApp -AppType PowerPoint)
}

function Close-OfficeInterop {
    try {
        if ($script:OfficeInterop.Word) {
            try { $script:OfficeInterop.Word.Quit() } catch {}
            Release-ComObjectSafe $script:OfficeInterop.Word
        }
    } catch {}
    try {
        if ($script:OfficeInterop.Excel) {
            try { $script:OfficeInterop.Excel.Quit() } catch {}
            Release-ComObjectSafe $script:OfficeInterop.Excel
        }
    } catch {}
    try {
        if ($script:OfficeInterop.PowerPoint) {
            try { $script:OfficeInterop.PowerPoint.Quit() } catch {}
            Release-ComObjectSafe $script:OfficeInterop.PowerPoint
        }
    } catch {}

    $script:OfficeInterop.Word = $null
    $script:OfficeInterop.Excel = $null
    $script:OfficeInterop.PowerPoint = $null
    $script:OfficeInterop.Initialized = $false

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Convert-LegacyOfficeViaHelperToTemp {
    param(
        [string]$FilePath,
        [ValidateSet('Word','Excel','PowerPoint')][string]$AppType,
        [int]$TimeoutSec = 45
    )

    $result = [ordered]@{
        Success = $false
        TempPath = ''
        ConvertedType = ''
        Status = 'Failed'
        Error = ''
    }

    if (-not (Test-Path -LiteralPath $FilePath)) {
        $result.Error = 'File not found'
        return New-Object PSObject -Property $result
    }

    $targetExt = ''
    $saveCode = ''
    switch ($AppType) {
        'Word' {
            $targetExt = '.docx'
            $result.ConvertedType = 'docx'
        }
        'Excel' {
            $targetExt = '.xlsx'
            $result.ConvertedType = 'xlsx'
        }
        'PowerPoint' {
            $targetExt = '.pptx'
            $result.ConvertedType = 'pptx'
        }
    }

    $tempPath = Join-Path $env:TEMP ('ORT_CONVERT_' + [guid]::NewGuid().ToString() + $targetExt)
    $helperOut = Join-Path $env:TEMP ('ORT_CONVERT_RESULT_' + [guid]::NewGuid().ToString() + '.json')

    $helper = @"
param(
    [string]`$InputPath,
    [string]`$OutputPath,
    [string]`$ResultPath,
    [string]`$AppType
)

`$res = [ordered]@{
    Success = `$false
    TempPath = `$OutputPath
    ConvertedType = ''
    Status = 'Failed'
    Error = ''
}

try {
    switch (`$AppType) {
        'Word' {
            `$res.ConvertedType = 'docx'
            `$app = New-Object -ComObject Word.Application
            `$app.Visible = `$false
            `$app.DisplayAlerts = 0
            `$doc = `$null
            try {
                `$doc = `$app.Documents.Open(`$InputPath, `$false, `$true)
                try {
                    `$doc.SaveAs2(`$OutputPath, 16)
                } catch {
                    `$doc.SaveAs([ref]`$OutputPath, [ref]16)
                }
            } finally {
                if (`$doc) { try { `$doc.Close(0) } catch { try { `$doc.Close() } catch {} } ; try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(`$doc) } catch {} }
                if (`$app) { try { `$app.Quit() } catch {} ; try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(`$app) } catch {} }
            }
        }
        'Excel' {
            `$res.ConvertedType = 'xlsx'
            `$app = New-Object -ComObject Excel.Application
            `$app.Visible = `$false
            `$app.DisplayAlerts = `$false
            try { `$app.ScreenUpdating = `$false } catch {}
            try { `$app.EnableEvents = `$false } catch {}
            `$wb = `$null
            try {
                `$wb = `$app.Workbooks.Open(`$InputPath, 0, `$true)
                `$wb.SaveAs(`$OutputPath, 51)
            } finally {
                if (`$wb) { try { `$wb.Close(`$false) } catch { try { `$wb.Close() } catch {} } ; try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(`$wb) } catch {} }
                if (`$app) { try { `$app.Quit() } catch {} ; try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(`$app) } catch {} }
            }
        }
        'PowerPoint' {
            `$res.ConvertedType = 'pptx'
            `$app = New-Object -ComObject PowerPoint.Application
            try { `$app.Visible = 0 } catch {}
            `$pres = `$null
            try {
                `$pres = `$app.Presentations.Open(`$InputPath, `$false, `$true, `$false)
                `$pres.SaveAs(`$OutputPath, 24)
            } finally {
                if (`$pres) { try { `$pres.Close() } catch {} ; try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(`$pres) } catch {} }
                if (`$app) { try { `$app.Quit() } catch {} ; try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(`$app) } catch {} }
            }
        }
    }

    if (Test-Path -LiteralPath `$OutputPath) {
        `$res.Success = `$true
        `$res.Status = 'Converted'
    } else {
        `$res.Error = 'No conversion result'
    }
}
catch {
    `$res.Status = 'Failed'
    `$res.Error = `$_.Exception.Message
}

try {
    [System.IO.File]::WriteAllText(`$ResultPath, ((New-Object PSObject -Property `$res) | ConvertTo-Json -Depth 4), [System.Text.Encoding]::UTF8)
} catch {}
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
"@

    $helperPs1 = Join-Path $env:TEMP ('ORT_CONVERT_HELPER_' + [guid]::NewGuid().ToString() + '.ps1')
    try {
        [System.IO.File]::WriteAllText($helperPs1, $helper, [System.Text.Encoding]::UTF8)
        $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList @(
            '-NoLogo','-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File', $helperPs1,
            '-InputPath', $FilePath,
            '-OutputPath', $tempPath,
            '-ResultPath', $helperOut,
            '-AppType', $AppType
        ) -WindowStyle Minimized -PassThru

        $completed = $proc.WaitForExit($TimeoutSec * 1000)
        if (-not $completed) {
            try { $proc.Kill() } catch {}
            $result.Status = 'Timeout'
            $result.Error = 'Conversion timeout'
            return New-Object PSObject -Property $result
        }

        if (Test-Path -LiteralPath $helperOut) {
            try {
                $raw = Get-Content -LiteralPath $helperOut -Raw -Encoding UTF8
                if ($raw) {
                    $obj = $raw | ConvertFrom-Json
                    if ($obj) {
                        return $obj
                    }
                }
            } catch {
                $result.Error = $_.Exception.Message
            }
        } else {
            $result.Error = 'No helper result'
        }
    }
    catch {
        $result.Error = $_.Exception.Message
    }
    finally {
        if ((-not $script:LegacyKeepTempConverted) -and (-not $result.Success) -and (Test-Path -LiteralPath $tempPath)) {
            try { Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue } catch {}
        }
        if (Test-Path -LiteralPath $helperOut) { try { Remove-Item -LiteralPath $helperOut -Force -ErrorAction SilentlyContinue } catch {} }
        if (Test-Path -LiteralPath $helperPs1) { try { Remove-Item -LiteralPath $helperPs1 -Force -ErrorAction SilentlyContinue } catch {} }
    }

    return New-Object PSObject -Property $result
}

function Convert-LegacyOfficeToOpenXmlTemp {
    param(
        [string]$FilePath,
        [ValidateSet('Word','Excel','PowerPoint')][string]$AppType,
        [int]$TimeoutSec = 45
    )

    $result = [ordered]@{
        Success = $false
        TempPath = ''
        ConvertedType = ''
        Status = 'Failed'
        Error = ''
        Method = ''
    }

    if (-not (Test-Path -LiteralPath $FilePath)) {
        $result.Error = 'File not found'
        return New-Object PSObject -Property $result
    }

    switch ($AppType) {
        'Word' { $result.ConvertedType = 'docx' }
        'Excel' { $result.ConvertedType = 'xlsx' }
        'PowerPoint' { $result.ConvertedType = 'pptx' }
    }

    # 優先使用 Office 2007 相容性套件 Converter
    $converterResult = Convert-LegacyOfficeViaConverterToTemp -FilePath $FilePath -AppType $AppType -TimeoutSec $TimeoutSec
    if ($converterResult -and $converterResult.Success) {
        return $converterResult
    }

    # Converter 找不到或失敗時，再退回共用 COM
    $app = Ensure-OfficeInteropApp -AppType $AppType
    if (-not $app) {
        $helperResult = Convert-LegacyOfficeViaHelperToTemp -FilePath $FilePath -AppType $AppType -TimeoutSec $TimeoutSec
        if ($helperResult) {
            if (-not $helperResult.Success -and $converterResult -and -not [string]::IsNullOrWhiteSpace($converterResult.Error)) {
                $helperResult.Error = (($converterResult.Error + ' / ' + [string]$helperResult.Error).Trim(' ','/'))
            }
            if (-not $helperResult.Status -or $helperResult.Status -eq 'Failed') {
                if ($converterResult -and $converterResult.Status) {
                    $helperResult.Status = [string]$converterResult.Status
                }
            }
            return $helperResult
        }
        return $converterResult
    }

    $targetExt = '.' + $result.ConvertedType
    $tempPath = Join-Path $env:TEMP ('ORT_CONVERT_' + [guid]::NewGuid().ToString() + $targetExt)
    $doc = $null
    $wb = $null
    $pres = $null

    try {
        switch ($AppType) {
            'Word' {
                $doc = $app.Documents.Open($FilePath, $false, $true)
                try { $doc.SaveAs2($tempPath, 16) } catch { $doc.SaveAs([ref]$tempPath, [ref]16) }
                $result.Success = (Test-Path -LiteralPath $tempPath)
            }
            'Excel' {
                $wb = $app.Workbooks.Open($FilePath, 0, $true)
                $wb.SaveAs($tempPath, 51)
                $result.Success = (Test-Path -LiteralPath $tempPath)
            }
            'PowerPoint' {
                $pres = $app.Presentations.Open($FilePath, $false, $true, $false)
                $pres.SaveAs($tempPath, 24)
                $result.Success = (Test-Path -LiteralPath $tempPath)
            }
        }

        if ($result.Success) {
            $result.TempPath = $tempPath
            $result.Status = 'ConvertedByCom'
            $result.Method = 'COM'
        } else {
            $result.Status = if ($converterResult -and $converterResult.Status) { [string]$converterResult.Status } else { 'Failed' }
            $result.Error = if ($converterResult -and $converterResult.Error) { [string]$converterResult.Error } else { 'No conversion result' }
        }
    } catch {
        try {
            if ($doc) { try { $doc.Close(0) } catch { try { $doc.Close() } catch {} } ; Release-ComObjectSafe $doc ; $doc = $null }
            if ($wb)  { try { $wb.Close($false) } catch { try { $wb.Close() } catch {} } ; Release-ComObjectSafe $wb ; $wb = $null }
            if ($pres){ try { $pres.Close() } catch {} ; Release-ComObjectSafe $pres ; $pres = $null }
        } catch {}

        try {
            if ($script:OfficeInterop[$AppType]) {
                try { $script:OfficeInterop[$AppType].Quit() } catch {}
                Release-ComObjectSafe $script:OfficeInterop[$AppType]
                $script:OfficeInterop[$AppType] = $null
            }
        } catch {}

        $helperResult = Convert-LegacyOfficeViaHelperToTemp -FilePath $FilePath -AppType $AppType -TimeoutSec $TimeoutSec
        if ($helperResult) {
            if ($helperResult.Success) {
                try { $helperResult | Add-Member -NotePropertyName Method -NotePropertyValue 'COM Helper' -Force } catch {}
            } elseif ($converterResult -and -not [string]::IsNullOrWhiteSpace($converterResult.Error)) {
                $helperResult.Error = (($converterResult.Error + ' / ' + [string]$helperResult.Error).Trim(' ','/'))
            }
            return $helperResult
        }

        if ($converterResult) { return $converterResult }
        $result.Error = $_.Exception.Message
    } finally {
        if ($doc) {
            try { $doc.Close(0) } catch { try { $doc.Close() } catch {} }
            Release-ComObjectSafe $doc
        }
        if ($wb) {
            try { $wb.Close($false) } catch { try { $wb.Close() } catch {} }
            Release-ComObjectSafe $wb
        }
        if ($pres) {
            try { $pres.Close() } catch {}
            Release-ComObjectSafe $pres
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    return New-Object PSObject -Property $result
}
function Toggle-LegacyConversionMode {
    $script:LegacyConversionMode = -not [bool]$script:LegacyConversionMode
    Save-StateAndStatus -Status (T '完成' 'Done') -Summary ((T 'Legacy 轉新版已切換為' 'Legacy conversion switched to') + ': ' + $script:LegacyConversionMode)
    Reset-UiCachesSafe
}

function Toggle-NamingMode {
    switch ($script:NamingMode) {
        'Conservative' { $script:NamingMode = 'Smart' }
        'Smart' { $script:NamingMode = 'OriginalFirst' }
        default { $script:NamingMode = 'Conservative' }
    }
    $script:LastNamingMode = $script:NamingMode
    Save-StateAndStatus -Status (T '完成' 'Done') -Summary ((T '智慧命名模式已切換為' 'Naming mode switched to') + ': ' + $script:NamingMode)
    Reset-UiCachesSafe
}

function Get-NamingModeDisplay {
    switch ($script:NamingMode) {
        'Conservative' { return (T '保守命名' 'Conservative') }
        'OriginalFirst' { return (T '原檔名優先' 'Original First') }
        default { return (T '智慧命名' 'Smart Naming') }
    }
}


function Get-NamingConfidence {
    param(
        [string]$ContentSource,
        [string]$PreviewText
    )

    $len = 0
    if ($PreviewText) { $len = $PreviewText.Length }

    if ($ContentSource -eq 'Converted OpenXML' -or $ContentSource -eq 'OpenXML') {
        if ($len -ge 20) { return 'High' }
        return 'Medium'
    }
    if ($ContentSource -eq 'PDF text' -or $ContentSource -eq 'RTF text' -or $ContentSource -eq 'Legacy text') {
        if ($len -ge 20) { return 'Medium' }
        return 'Low'
    }
    if ($ContentSource -eq 'FileHash only' -or $ContentSource -eq 'Metadata' -or [string]::IsNullOrWhiteSpace($ContentSource)) {
        return 'Low'
    }
    return 'Medium'
}

function Get-LegacyReadableText {
    param(
        [string]$FilePath,
        [int]$MaxReadKB = 512
    )

    if (-not (Test-Path -LiteralPath $FilePath)) { return '' }

    try {
        $fs = [IO.File]::Open($FilePath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        try {
            $maxBytes = $MaxReadKB * 1KB
            $readBytes = [Math]::Min([int64]$fs.Length, [int64]$maxBytes)
            if ($readBytes -le 0) { return '' }
            $buffer = New-Object byte[] $readBytes
            [void]$fs.Read($buffer, 0, $readBytes)
            $texts = @()
            $encodings = @([Text.Encoding]::Unicode,[Text.Encoding]::BigEndianUnicode,[Text.Encoding]::ASCII,[Text.Encoding]::GetEncoding(950))
            foreach ($enc in $encodings) {
                try {
                    $txt = $enc.GetString($buffer)
                    if ($txt) {
                        $txt = [Regex]::Replace($txt, '[--]', ' ')
                        $txt = [Regex]::Replace($txt, '\?{3,}', ' ')
                        $txt = [Regex]::Replace($txt, '\s+', ' ').Trim()
                        $matches = [regex]::Matches($txt, '[一-鿿A-Za-z0-9][一-鿿A-Za-z0-9\-\_\(\)\/\.\,\:\s]{5,}')
                        foreach ($m in $matches) { $texts += $m.Value }
                    }
                }
                catch {}
            }
            $texts = $texts | Where-Object { $_ -and $_.Trim().Length -ge 6 } | Select-Object -Unique
            return (($texts -join ' ') -replace '\s+', ' ').Trim()
        }
        finally { $fs.Close() }
    }
    catch { return '' }
}

function Get-OfficeContentInfo {
    param([string]$FilePath)

    $ext = [IO.Path]::GetExtension($FilePath).ToLowerInvariant()

    $result = [ordered]@{
        ContentHash = ''
        PreviewText = ''
        ParseStatus = T '解析失敗' 'Parse failed'
        ParseReason = ''
        ContentSource = ''
        NamingConfidence = 'Low'
        LegacyConverted = $false
        ConvertedType = ''
        ConversionStatus = ''
        ProtectionState = ''
        ProtectionDetail = ''
    }

    try {
        switch ($ext) {
            '.docx' {
                $ooxmlProtection = Test-OoxmlProtectionState -FilePath $FilePath
                if ($ooxmlProtection) {
                    $result.ProtectionState = [string]$ooxmlProtection.State
                    $result.ProtectionDetail = [string]$ooxmlProtection.Detail
                }
                if ($ooxmlProtection -and $ooxmlProtection.Skip) {
                    $result.ParseStatus = T '已略過' 'Skipped'
                    $result.ParseReason = if ($ooxmlProtection.Detail) { [string]$ooxmlProtection.Detail } else { [string]$ooxmlProtection.State }
                    $result.ConversionStatus = 'SkippedByProtectionCheck'
                    $result.ContentSource = 'OOXML protection check'
                    return New-Object PSObject -Property $result
                }
                $xml = Get-ZipEntryText -ZipPath $FilePath -Candidates @('word/document.xml')
                if ($xml) {
                    $plain = Normalize-XmlText $xml
                    if ($plain) {
                        $result.ContentHash = Get-TextHash $plain
                        $result.PreviewText = $plain.Substring(0, [Math]::Min(200, $plain.Length))
                        $result.ParseStatus = T '解析成功' 'Parsed'
                        $result.ContentSource = 'OpenXML'
                    }
                    else {
                        $result.ParseStatus = T '無法取得內容' 'No content extracted'
                    }
                }
                else {
                    $result.ParseStatus = T '無法取得內容' 'No content extracted'
                    $result.ParseReason = ''
                }
            }

            '.xlsx' {
                $ooxmlProtection = Test-OoxmlProtectionState -FilePath $FilePath
                if ($ooxmlProtection) {
                    $result.ProtectionState = [string]$ooxmlProtection.State
                    $result.ProtectionDetail = [string]$ooxmlProtection.Detail
                }
                if ($ooxmlProtection -and $ooxmlProtection.Skip) {
                    $result.ParseStatus = T '已略過' 'Skipped'
                    $result.ParseReason = if ($ooxmlProtection.Detail) { [string]$ooxmlProtection.Detail } else { [string]$ooxmlProtection.State }
                    $result.ConversionStatus = 'SkippedByProtectionCheck'
                    $result.ContentSource = 'OOXML protection check'
                    return New-Object PSObject -Property $result
                }
                Add-Type -AssemblyName System.IO.Compression.FileSystem
                $zip = $null
                $parts = @()
                try {
                    $zip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
                    $sharedEntry = $zip.Entries | Where-Object { $_.FullName -ieq 'xl/sharedStrings.xml' } | Select-Object -First 1
                    if ($sharedEntry) {
                        $sr = New-Object IO.StreamReader($sharedEntry.Open(), [Text.Encoding]::UTF8)
                        try { $parts += $sr.ReadToEnd() } finally { $sr.Close() }
                    }
                    $sheetEntries = $zip.Entries | Where-Object { $_.FullName -match '^xl/worksheets/sheet\d+\.xml$' } | Sort-Object FullName
                    foreach ($sheetEntry in $sheetEntries) {
                        $sr2 = New-Object IO.StreamReader($sheetEntry.Open(), [Text.Encoding]::UTF8)
                        try {
                            $sheetXml = $sr2.ReadToEnd()
                            if (-not [string]::IsNullOrWhiteSpace($sheetXml)) { $parts += $sheetXml }
                        } finally { $sr2.Close() }
                    }
                    $workbookEntry = $zip.Entries | Where-Object { $_.FullName -ieq 'xl/workbook.xml' } | Select-Object -First 1
                    if ($workbookEntry) {
                        $sr3 = New-Object IO.StreamReader($workbookEntry.Open(), [Text.Encoding]::UTF8)
                        try {
                            $workbookXml = $sr3.ReadToEnd()
                            if (-not [string]::IsNullOrWhiteSpace($workbookXml)) { $parts += $workbookXml }
                        } finally { $sr3.Close() }
                    }
                    $combined = ($parts -join ' ')
                    if (-not [string]::IsNullOrWhiteSpace($combined)) {
                        $plain = Normalize-XmlText $combined
                        if (-not [string]::IsNullOrWhiteSpace($plain)) {
                            $result.ContentHash = Get-TextHash $plain
                            $result.PreviewText = $plain.Substring(0, [Math]::Min(200, $plain.Length))
                            $result.ParseStatus = T '解析成功' 'Parsed'
                            $result.ContentSource = 'OpenXML'
                        }
                        else {
                            $result.ParseStatus = T '無法取得內容' 'No content extracted'
                        }
                    }
                    else {
                        $result.ParseStatus = T '無法取得內容' 'No content extracted'
                        $result.ParseReason = ''
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
                $ooxmlProtection = Test-OoxmlProtectionState -FilePath $FilePath
                if ($ooxmlProtection) {
                    $result.ProtectionState = [string]$ooxmlProtection.State
                    $result.ProtectionDetail = [string]$ooxmlProtection.Detail
                }
                if ($ooxmlProtection -and $ooxmlProtection.Skip) {
                    $result.ParseStatus = T '已略過' 'Skipped'
                    $result.ParseReason = if ($ooxmlProtection.Detail) { [string]$ooxmlProtection.Detail } else { [string]$ooxmlProtection.State }
                    $result.ConversionStatus = 'SkippedByProtectionCheck'
                    $result.ContentSource = 'OOXML protection check'
                    return New-Object PSObject -Property $result
                }
                Add-Type -AssemblyName System.IO.Compression.FileSystem
                $zip = $null
                $parts = @()
                try {
                    $zip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
                    $slideEntries = $zip.Entries | Where-Object { $_.FullName -match '^ppt/slides/slide\d+\.xml$' } | Sort-Object FullName
                    foreach ($entry in $slideEntries) {
                        $sr = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
                        try {
                            $slideXml = $sr.ReadToEnd()
                            if ($slideXml) { $parts += $slideXml }
                        } finally { $sr.Close() }
                    }
                    $combined = ($parts -join ' ')
                    if ($combined) {
                        $plain = Normalize-XmlText $combined
                        if ($plain) {
                            $result.ContentHash = Get-TextHash $plain
                            $result.PreviewText = $plain.Substring(0, [Math]::Min(200, $plain.Length))
                            $result.ParseStatus = T '解析成功' 'Parsed'
                            $result.ContentSource = 'OpenXML'
                        }
                        else {
                            $result.ParseStatus = T '無法取得內容' 'No content extracted'
                        }
                    }
                    else {
                        $result.ParseStatus = T '無法取得內容' 'No content extracted'
                        $result.ParseReason = ''
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

            '.doc' {
                $fileInfo = Get-Item -LiteralPath $FilePath -ErrorAction SilentlyContinue
                $canTryOffice = $fileInfo -and ($fileInfo.Length -le ($script:LegacyOfficeMaxFileMB * 1MB))

                $probeDecision = Test-ShouldSkipLegacyOfficeByProbe -FilePath $FilePath -AppType Word
                if ($probeDecision) {
                    $result.ProtectionState = [string]$probeDecision.State
                    $result.ProtectionDetail = [string]$probeDecision.Detail
                }
                if ($probeDecision -and $probeDecision.Skip) {
                    $result.ParseStatus = T '已略過' 'Skipped'
                    $result.ParseReason = if ($probeDecision.Detail) { [string]$probeDecision.Detail } else { [string]$probeDecision.State }
                    $result.ConversionStatus = 'SkippedByProbe'
                    $result.ContentSource = 'Protection probe'
                    return New-Object PSObject -Property $result
                }
                if ($script:LegacyConversionMode -and $canTryOffice) {
                    $conv = Convert-LegacyOfficeToOpenXmlTemp -FilePath $FilePath -AppType Word -TimeoutSec $script:LegacyOfficeTimeoutSec
                    if ($conv -and $conv.Success -and (Test-Path -LiteralPath $conv.TempPath)) {
                        try {
                            $tmpInfo = Get-OfficeContentInfo -FilePath $conv.TempPath
                            $result.ContentHash = $tmpInfo.ContentHash
                            $result.PreviewText = $tmpInfo.PreviewText
                            $result.ParseStatus = $tmpInfo.ParseStatus
                            $result.ParseReason = $tmpInfo.ParseReason
                            $result.ContentSource = 'Converted OpenXML'
                            $result.LegacyConverted = $true
                            $result.ConvertedType = 'docx'
                            $result.ConversionStatus = 'Success'
                        } finally {
                            if (-not $script:LegacyKeepTempConverted) { try { Remove-Item -LiteralPath $conv.TempPath -Force -ErrorAction SilentlyContinue } catch {} }
                        }
                    } else {
                        $result.ConversionStatus = if ($conv) { [string]$conv.Status } else { 'Failed' }
                        if ($conv -and $conv.Error) { $result.ParseReason = [string]$conv.Error }
                    }
                }
                if (-not $result.ContentHash) {
                    $maxRead = if ($script:LegacyQuickMode) { [Math]::Min($script:LegacyMaxReadKB, 128) } else { $script:LegacyMaxReadKB }
                    $text = Get-LegacyReadableText -FilePath $FilePath -MaxReadKB $maxRead
                    if ((-not (Test-LegacyTextLooksUseful $text)) -and $script:LegacyOfficeFallback -and $canTryOffice) {
                        $officeText = Get-LegacyOfficeTextViaHelper -FilePath $FilePath -AppType Word -TimeoutSec $script:LegacyOfficeTimeoutSec -MaxChars ($script:LegacyTextPreviewLength * 10)
                        if (Test-LegacyTextLooksUseful $officeText) { $text = $officeText }
                    }
                    if ($text) {
                        $result.ContentHash = Get-TextHash $text
                        $result.PreviewText = $text.Substring(0, [Math]::Min($script:LegacyTextPreviewLength, $text.Length))
                        $result.ParseStatus = T '解析成功' 'Parsed'
                        $result.ParseReason = ''
                        if (-not $result.ContentSource) { $result.ContentSource = 'Legacy text' }
                    }
                    elseif (-not $result.ParseStatus -or $result.ParseStatus -eq (T '解析失敗' 'Parse failed')) {
                        $result.ParseStatus = T '無法取得內容' 'No content extracted'
                    }
                }
            }

            '.xls' {
                $fileInfo = Get-Item -LiteralPath $FilePath -ErrorAction SilentlyContinue
                $canTryOffice = $fileInfo -and ($fileInfo.Length -le ($script:LegacyOfficeMaxFileMB * 1MB))

                $probeDecision = Test-ShouldSkipLegacyOfficeByProbe -FilePath $FilePath -AppType Excel
                if ($probeDecision) {
                    $result.ProtectionState = [string]$probeDecision.State
                    $result.ProtectionDetail = [string]$probeDecision.Detail
                }
                if ($probeDecision -and $probeDecision.Skip) {
                    $result.ParseStatus = T '已略過' 'Skipped'
                    $result.ParseReason = if ($probeDecision.Detail) { [string]$probeDecision.Detail } else { [string]$probeDecision.State }
                    $result.ConversionStatus = 'SkippedByProbe'
                    $result.ContentSource = 'Protection probe'
                    return New-Object PSObject -Property $result
                }
                if ($script:LegacyConversionMode -and $canTryOffice) {
                    $conv = Convert-LegacyOfficeToOpenXmlTemp -FilePath $FilePath -AppType Excel -TimeoutSec $script:LegacyOfficeTimeoutSec
                    if ($conv -and $conv.Success -and (Test-Path -LiteralPath $conv.TempPath)) {
                        try {
                            $tmpInfo = Get-OfficeContentInfo -FilePath $conv.TempPath
                            $result.ContentHash = $tmpInfo.ContentHash
                            $result.PreviewText = $tmpInfo.PreviewText
                            $result.ParseStatus = $tmpInfo.ParseStatus
                            $result.ParseReason = $tmpInfo.ParseReason
                            $result.ContentSource = 'Converted OpenXML'
                            $result.LegacyConverted = $true
                            $result.ConvertedType = 'xlsx'
                            $result.ConversionStatus = 'Success'
                        } finally {
                            if (-not $script:LegacyKeepTempConverted) { try { Remove-Item -LiteralPath $conv.TempPath -Force -ErrorAction SilentlyContinue } catch {} }
                        }
                    } else {
                        $result.ConversionStatus = if ($conv) { [string]$conv.Status } else { 'Failed' }
                        if ($conv -and $conv.Error) { $result.ParseReason = [string]$conv.Error }
                    }
                }
                if (-not $result.ContentHash) {
                    $maxRead = if ($script:LegacyQuickMode) { [Math]::Min($script:LegacyMaxReadKB, 128) } else { $script:LegacyMaxReadKB }
                    $excelText = Get-LegacyReadableText -FilePath $FilePath -MaxReadKB $maxRead
                    if ((-not (Test-LegacyTextLooksUseful $excelText)) -and $script:LegacyOfficeFallback -and $canTryOffice) {
                        $officeText = Get-LegacyOfficeTextViaHelper -FilePath $FilePath -AppType Excel -TimeoutSec $script:LegacyOfficeTimeoutSec -MaxChars ($script:LegacyTextPreviewLength * 12)
                        if (Test-LegacyTextLooksUseful $officeText) { $excelText = $officeText }
                    }
                    if ($excelText) {
                        $result.ContentHash = Get-TextHash $excelText
                        $result.PreviewText = $excelText.Substring(0, [Math]::Min($script:LegacyTextPreviewLength, $excelText.Length))
                        $result.ParseStatus = T '解析成功' 'Parsed'
                        $result.ParseReason = ''
                        if (-not $result.ContentSource) { $result.ContentSource = 'Legacy text' }
                    }
                    elseif (-not $result.ParseStatus -or $result.ParseStatus -eq (T '解析失敗' 'Parse failed')) {
                        $result.ParseStatus = T '無法取得內容' 'No content extracted'
                    }
                }
            }

            '.ppt' {
                $fileInfo = Get-Item -LiteralPath $FilePath -ErrorAction SilentlyContinue
                $canTryOffice = $fileInfo -and ($fileInfo.Length -le ($script:LegacyOfficeMaxFileMB * 1MB))

                $probeDecision = Test-ShouldSkipLegacyOfficeByProbe -FilePath $FilePath -AppType PowerPoint
                if ($probeDecision) {
                    $result.ProtectionState = [string]$probeDecision.State
                    $result.ProtectionDetail = [string]$probeDecision.Detail
                }
                if ($probeDecision -and $probeDecision.Skip) {
                    $result.ParseStatus = T '已略過' 'Skipped'
                    $result.ParseReason = if ($probeDecision.Detail) { [string]$probeDecision.Detail } else { [string]$probeDecision.State }
                    $result.ConversionStatus = 'SkippedByProbe'
                    $result.ContentSource = 'Protection probe'
                    return New-Object PSObject -Property $result
                }
                if ($script:LegacyConversionMode -and $canTryOffice) {
                    $conv = Convert-LegacyOfficeToOpenXmlTemp -FilePath $FilePath -AppType PowerPoint -TimeoutSec ([Math]::Max($script:LegacyOfficeTimeoutSec, 15))
                    if ($conv -and $conv.Success -and (Test-Path -LiteralPath $conv.TempPath)) {
                        try {
                            $tmpInfo = Get-OfficeContentInfo -FilePath $conv.TempPath
                            $result.ContentHash = $tmpInfo.ContentHash
                            $result.PreviewText = $tmpInfo.PreviewText
                            $result.ParseStatus = $tmpInfo.ParseStatus
                            $result.ParseReason = $tmpInfo.ParseReason
                            $result.ContentSource = 'Converted OpenXML'
                            $result.LegacyConverted = $true
                            $result.ConvertedType = 'pptx'
                            $result.ConversionStatus = 'Success'
                        } finally {
                            if (-not $script:LegacyKeepTempConverted) { try { Remove-Item -LiteralPath $conv.TempPath -Force -ErrorAction SilentlyContinue } catch {} }
                        }
                    } else {
                        $result.ConversionStatus = if ($conv) { [string]$conv.Status } else { 'Failed' }
                        if ($conv -and $conv.Error) { $result.ParseReason = [string]$conv.Error }
                    }
                }
                if (-not $result.ContentHash) {
                    $legacyText = Get-LegacyReadableText -FilePath $FilePath -MaxReadKB ([Math]::Min($script:LegacyMaxReadKB, 256))
                    if ((-not (Test-LegacyTextLooksUseful $legacyText)) -and $script:LegacyOfficeFallback -and $canTryOffice) {
                        $officeText = Get-LegacyOfficeTextViaHelper -FilePath $FilePath -AppType PowerPoint -TimeoutSec ([Math]::Max($script:LegacyOfficeTimeoutSec, 15)) -MaxChars ($script:LegacyTextPreviewLength * 12)
                        if (Test-LegacyTextLooksUseful $officeText) { $legacyText = $officeText }
                    }
                    if ($legacyText) {
                        $result.ContentHash = Get-TextHash $legacyText
                        $result.PreviewText = $legacyText.Substring(0, [Math]::Min($script:LegacyTextPreviewLength, $legacyText.Length))
                        $result.ParseStatus = T '解析成功' 'Parsed'
                        $result.ParseReason = ''
                        if (-not $result.ContentSource) { $result.ContentSource = 'Legacy text' }
                    }
                    elseif (-not $result.ParseStatus -or $result.ParseStatus -eq (T '解析失敗' 'Parse failed')) {
                        $result.ParseStatus = T '無法取得內容' 'No content extracted'
                    }
                }
            }

            '.rtf' {
                $rtfText = Get-RtfTextBestEffort -FilePath $FilePath
                if ($rtfText) {
                    $result.ContentHash = Get-TextHash $rtfText
                    $result.PreviewText = $rtfText.Substring(0, [Math]::Min(200, $rtfText.Length))
                    $result.ParseStatus = T '解析成功' 'Parsed'
                    $result.ParseReason = ''
                    $result.ContentSource = 'RTF text'
                }
                else {
                    $result.ParseStatus = T '無法取得內容' 'No content extracted'
                    $result.ParseReason = 'RTF'
                }
            }

            '.pdf' {
                $pdfText = Get-PdfTextBestEffort -FilePath $FilePath -MaxChars ([Math]::Max(($script:LegacyTextPreviewLength * 20), 4000))
                $fileInfo = Get-Item -LiteralPath $FilePath -ErrorAction SilentlyContinue
                $canTryOffice = $script:LegacyOfficeFallback -and $fileInfo -and ($fileInfo.Length -le ($script:LegacyOfficeMaxFileMB * 1MB))
                if ((-not (Test-LegacyTextLooksUseful $pdfText)) -and $canTryOffice) {
                    $officeText = Get-LegacyOfficeTextViaHelper -FilePath $FilePath -AppType Word -TimeoutSec $script:LegacyOfficeTimeoutSec -MaxChars ($script:LegacyTextPreviewLength * 20)
                    if (Test-LegacyTextLooksUseful $officeText) { $pdfText = $officeText }
                }
                if ($pdfText) {
                    $result.ContentHash = Get-TextHash $pdfText
                    $result.PreviewText = $pdfText.Substring(0, [Math]::Min(200, $pdfText.Length))
                    $result.ParseStatus = T '解析成功' 'Parsed'
                    $result.ParseReason = 'PDF text'
                    $result.ContentSource = 'PDF text'
                }
                else {
                    try { $result.ContentHash = (Get-FileHash -LiteralPath $FilePath -Algorithm SHA256 -ErrorAction SilentlyContinue).Hash } catch { $result.ContentHash = '' }
                    $result.PreviewText = ''
                    $result.ParseStatus = T '僅檔案層級' 'File-level only'
                    $result.ParseReason = T 'PDF 無法抽出文字，改用檔案雜湊' 'PDF text could not be extracted; fell back to file hash'
                    $result.ContentSource = 'FileHash only'
                }
            }
            '.txt' {
                try {
                    $rawText = Get-Content -LiteralPath $FilePath -Raw -Encoding UTF8 -ErrorAction Stop
                }
                catch {
                    try {
                        $rawText = Get-Content -LiteralPath $FilePath -Raw -Encoding Default -ErrorAction Stop
                    }
                    catch {
                        $rawText = ''
                    }
                }

                if (-not [string]::IsNullOrWhiteSpace($rawText)) {
                    $plain = Normalize-PlainText $rawText
                    if (-not [string]::IsNullOrWhiteSpace($plain)) {
                        $result.ContentHash = Get-TextHash $plain
                        $result.PreviewText = $plain.Substring(0, [Math]::Min(200, $plain.Length))
                        $result.ParseStatus = T '解析成功' 'Parsed'
                        $result.ParseReason = ''
                        $result.ContentSource = 'Text file'
                    }
                    else {
                        $result.ParseStatus = T '無法取得內容' 'No content extracted'
                        $result.ParseReason = 'TXT'
                    }
                }
                else {
                    $result.ParseStatus = T '無法取得內容' 'No content extracted'
                    $result.ParseReason = 'TXT'
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

    $result.NamingConfidence = Get-NamingConfidence -ContentSource $result.ContentSource -PreviewText $result.PreviewText
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
        '.ppt'  { return (T 'PowerPoint 舊版簡報' 'Legacy PowerPoint Presentation') }
        '.rtf'  { return (T 'RTF 文件' 'RTF Document') }
        '.pdf'  { return (T 'PDF 文件' 'PDF Document') }
        '.txt'  { return (T '文字檔' 'Text File') }
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

function Normalize-PreviewTextForNaming {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

    $t = [string]$Text
    $t = $t -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ' '
    $t = $t -replace '�', ' '
    $t = $t -replace '[“”„‟"]', ''
    $t = $t -replace "[‘’‚‛']", ''
    $t = $t -replace '[‐‑–—]+', '-'
    $t = $t -replace '[|¦]+', ' '
    $t = $t -replace '[\t\r\n]+', "`n"
    $t = $t -replace '[ ]{2,}', ' '
    return $t.Trim()
}

function Test-LooksLikeGarbledText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return $true }

    $t = Normalize-PreviewTextForNaming $Text
    if ([string]::IsNullOrWhiteSpace($t)) { return $true }

    $len = $t.Length
    if ($len -lt 3) { return $true }

    $replacementCount = ([regex]::Matches($Text, '�')).Count
    $goodCount = ([regex]::Matches($t, '[\p{L}\p{N}\u4e00-\u9fff]')).Count
    $symbolCount = ([regex]::Matches($t, '[^ \p{L}\p{N}\u4e00-\u9fff\-\._:/]')).Count

    if ($replacementCount -ge 2) { return $true }
    if ($goodCount -lt [Math]::Max(3, [int]($len * 0.25))) { return $true }
    if ($symbolCount -gt [int]($len * 0.35)) { return $true }

    $lines = $t -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    if (@($lines).Count -gt 0) {
        $best = $lines | Sort-Object { - (($_ -replace '\s','').Length) } | Select-Object -First 1
        if ($best -match '^[0-9\W_]+$') { return $true }
    }

    return $false
}

function Get-BestNamingLine {
    param([string]$Text)

    $t = Normalize-PreviewTextForNaming $Text
    if ([string]::IsNullOrWhiteSpace($t)) { return '' }

    $lines = $t -split "`r?`n" |
        ForEach-Object { $_.Trim() } |
        Where-Object { $_ -and $_ -notmatch '^(page|sheet|工作表|第\s*\d+\s*頁)\b' }

    if (-not $lines) { return '' }

    $scored = foreach ($line in $lines) {
        $score = 0
        $cleanLen = ($line -replace '\s','').Length
        if ($cleanLen -ge 4) { $score += 3 }
        if ($line -match '[\u4e00-\u9fffA-Za-z]{3,}') { $score += 4 }
        if ($line -match '(invoice|report|contract|quotation|quote|statement|resume|meeting|proposal|spec|manual|presentation|agenda|minutes|發票|報告|合約|契約|估價|報價|對帳|履歷|會議|提案|規格|手冊|簡報)') { $score += 5 }
        if ($line -match '^[0-9\W_]+$') { $score -= 6 }
        if ($line.Length -gt 80) { $score -= 2 }
        [pscustomobject]@{
            Line  = $line
            Score = $score
            Len   = $cleanLen
        }
    }

    $pick = $scored | Sort-Object @{Expression='Score';Descending=$true}, @{Expression='Len';Descending=$true} | Select-Object -First 1
    if ($pick.Score -lt 1) { return '' }
    return [string]$pick.Line
}

function Get-NamingDateToken {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

    $patterns = @(
        '\b(20\d{2})[\/\.-](0?[1-9]|1[0-2])[\/\.-](0?[1-9]|[12]\d|3[01])\b',
        '\b(20\d{2})(0[1-9]|1[0-2])([0-2]\d|3[01])\b',
        '\b(20\d{2})年\s*(1[0-2]|0?[1-9])月\s*(3[01]|[12]?\d)日?\b',
        '\b(1[01]\d|0?\d{2})[\/\.-](0?[1-9]|1[0-2])[\/\.-](0?[1-9]|[12]\d|3[01])\b'
    )

    foreach ($p in $patterns) {
        $m = [regex]::Match($Text, $p)
        if ($m.Success) {
            try {
                if ($m.Groups[1].Value.Length -le 3) {
                    $y = [int]$m.Groups[1].Value + 1911
                } else {
                    $y = [int]$m.Groups[1].Value
                }
                $mo = [int]$m.Groups[2].Value
                $d = [int]$m.Groups[3].Value
                $dt = Get-Date -Year $y -Month $mo -Day $d -Hour 0 -Minute 0 -Second 0
                return $dt.ToString('yyyy-MM-dd')
            } catch {}
        }
    }

    return ''
}

function Get-NamingAmountToken {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

    $matches = [regex]::Matches($Text, '(?:NT\$|TWD|\$|USD|金額|總計|合計|total|amount)[^\d]{0,8}(\d{1,3}(?:,\d{3})+|\d{4,})', 'IgnoreCase')
    if ($matches.Count -gt 0) {
        $num = $matches[0].Groups[1].Value -replace ',', ''
        if ($num.Length -gt 8) { $num = $num.Substring(0,8) }
        return $num
    }

    return ''
}


function Test-WorkbookLooksLikeRoster {
    param(
        [string]$Text,
        [string]$Extension
    )

    $ext = ($Extension + '').ToLowerInvariant()
    if ($ext -notin @('.xls', '.xlsx')) { return $false }
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }

    $t = Normalize-PreviewTextForNaming ([string]$Text)

    $hits = 0

    $groups = @(
        '(?i)(^|[\s,_|:/\-])姓名([\s,_|:/\-]|$)|name',
        '(?i)(^|[\s,_|:/\-])日期([\s,_|:/\-]|$)|date|出生日期|填表日期',
        '(?i)身份證|身分證|身份證號|身分證號|身份證字號|身分證字號|id\s*no|id\s*number|身分證統一編號'
    )

    foreach ($g in $groups) {
        if ($t -match $g) { $hits++ }
    }

    return ($hits -ge 2)
}

function Get-WorkbookRosterName {
    param(
        [string]$Text,
        [string]$Extension
    )

    if (-not (Test-WorkbookLooksLikeRoster -Text $Text -Extension $Extension)) {
        return ''
    }

    $dateToken = Get-NamingDateToken -Text $Text
    if (-not [string]::IsNullOrWhiteSpace($dateToken)) {
        return ('名冊_{0}' -f $dateToken)
    }

    return '名冊'
}

function Get-NamingKeywordToken {
    param([string]$Text, [string]$Extension)

    if (Test-WorkbookLooksLikeRoster -Text $Text -Extension $Extension) {
        return '名冊'
    }

    $pairs = @(
        @{ Pattern='invoice|發票'; Value='Invoice' },
        @{ Pattern='quotation|quote|報價|估價'; Value='Quotation' },
        @{ Pattern='contract|agreement|合約|契約'; Value='Contract' },
        @{ Pattern='statement|對帳'; Value='Statement' },
        @{ Pattern='resume|cv|履歷'; Value='Resume' },
        @{ Pattern='meeting|minutes|會議紀錄|會議'; Value='Meeting' },
        @{ Pattern='report|報告'; Value='Report' },
        @{ Pattern='proposal|提案'; Value='Proposal' },
        @{ Pattern='spec|specification|規格'; Value='Spec' },
        @{ Pattern='manual|手冊|說明書'; Value='Manual' },
        @{ Pattern='presentation|slide|簡報'; Value='Presentation' },
        @{ Pattern='budget|預算'; Value='Budget' }
    )

    foreach ($p in $pairs) {
        if ($Text -match $p.Pattern) { return $p.Value }
    }

    switch ($Extension.ToLowerInvariant()) {
        '.docx' { return 'Document' }
        '.doc'  { return 'Document' }
        '.xlsx' { return 'Workbook' }
        '.xls'  { return 'Workbook' }
        '.pptx' { return 'Presentation' }
        '.ppt'  { return 'Presentation' }
        '.pdf'  { return 'PDF' }
        '.txt'  { return 'Text' }
        default { return 'File' }
    }
}

function Get-OriginalNameFallback {
    param($Row)

    $base = [IO.Path]::GetFileNameWithoutExtension([string]$Row.FileName)
    if ([string]::IsNullOrWhiteSpace($base)) { $base = 'RecoveredFile' }
    $base = $base -replace '^[\s\._-]+',''
    $base = $base -replace '[\s\._-]+$',''
    $base = $base -replace '(?i)^(copy of |copy_|recovered_|document_|scan_)',''
    $base = Get-SafeFileName $base
    if ([string]::IsNullOrWhiteSpace($base)) { $base = 'RecoveredFile' }
    return $base
}

function Build-SmartNameFromTokens {
    param(
        [string]$Keyword,
        [string]$DateToken,
        [string]$AmountToken,
        [string]$TitleToken
    )

    $parts = @()

    if (-not [string]::IsNullOrWhiteSpace($Keyword)) { $parts += $Keyword }
    if (-not [string]::IsNullOrWhiteSpace($DateToken)) { $parts += $DateToken }
    if (-not [string]::IsNullOrWhiteSpace($AmountToken)) { $parts += ('Amt' + $AmountToken) }

    if (-not [string]::IsNullOrWhiteSpace($TitleToken)) {
        $title = $TitleToken
        $title = $title -replace '(?i)\b(invoice|report|contract|quotation|statement|resume|meeting|proposal|spec|manual|presentation|document|workbook|text|pdf)\b', ''
        $title = $title -replace '\s+', ' '
        $title = $title.Trim()
        if ($title.Length -gt 28) { $title = $title.Substring(0, 28).Trim() }
        if ($title) { $parts += $title }
    }

    $parts = @($parts | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $joined = ($parts -join '_')
    $joined = Get-SafeFileName $joined
    if ($joined.Length -gt 80) { $joined = $joined.Substring(0,80).Trim(' ','_','-','.') }
    return $joined
}


function Get-NamingConfidenceScore {
    param(
        [string]$Keyword,
        [string]$DateToken,
        [string]$AmountToken,
        [string]$TitleToken,
        [bool]$LooksBad,
        [string]$Mode,
        [string]$FallbackBase
    )

    $score = 20
    if (-not [string]::IsNullOrWhiteSpace($Keyword)) { $score += 20 }
    if (-not [string]::IsNullOrWhiteSpace($DateToken)) { $score += 20 }
    if (-not [string]::IsNullOrWhiteSpace($AmountToken)) { $score += 15 }
    if (-not [string]::IsNullOrWhiteSpace($TitleToken)) { $score += 25 }
    if ($LooksBad) { $score -= 25 }

    switch ($Mode) {
        'Conservative' { $score -= 5 }
        'OriginalFirst' { $score -= 10 }
        default { }
    }

    if (-not [string]::IsNullOrWhiteSpace($FallbackBase) -and -not [string]::IsNullOrWhiteSpace($TitleToken) -and ($TitleToken -eq $FallbackBase)) {
        $score -= 10
    }

    if ($score -lt 0) { $score = 0 }
    if ($score -gt 100) { $score = 100 }
    return [int]$score
}

function Get-NamingConfidenceLabel {
    param([int]$Score)
    if ($Score -ge 80) { return 'High' }
    if ($Score -ge 55) { return 'Medium' }
    return 'Low'
}

function Get-SuggestedBaseNameInfo {
    param($Row)

    $ext = [string]$Row.Extension
    $fallback = Get-OriginalNameFallback -Row $Row
    $mode = [string]$script:NamingMode
    $baseName = ''
    $reason = ''
    $source = ''
    $keyword = ''
    $dateToken = ''
    $amountToken = ''
    $bestLine = ''
    $looksBad = $false

    if ($ext -eq '.xlsx') {
        $excelSmart = Get-ExcelSmartName -FilePath $Row.FullPath
        if (-not [string]::IsNullOrWhiteSpace($excelSmart)) {
            $baseName = $excelSmart
            $reason = 'ExcelSmart'
            $source = 'Excel'
        }
    }

    if ([string]::IsNullOrWhiteSpace($baseName) -and $mode -eq 'OriginalFirst') {
        $baseName = $fallback
        $reason = 'OriginalFirst'
        $source = 'Original'
    }

    if ([string]::IsNullOrWhiteSpace($baseName)) {
        $text = Normalize-PreviewTextForNaming ([string]$Row.PreviewText)
        $looksBad = Test-LooksLikeGarbledText $text

        if (-not $looksBad) {
            $bestLine = Get-BestNamingLine $text
            $keyword = Get-NamingKeywordToken -Text $text -Extension $ext
            $dateToken = Get-NamingDateToken -Text $text
            $amountToken = Get-NamingAmountToken -Text $text

            if ($ext -in @('.xls', '.xlsx')) {
                $rosterSmart = Get-WorkbookRosterName -Text $text -Extension $ext
                if (-not [string]::IsNullOrWhiteSpace($rosterSmart)) {
                    $baseName = $rosterSmart
                    $reason = 'RosterWorkbook'
                    $source = 'Workbook'
                    $keyword = '名冊'
                }
            }
        }

        $smart = ''
        if ([string]::IsNullOrWhiteSpace($baseName)) {
            $smart = Build-SmartNameFromTokens -Keyword $keyword -DateToken $dateToken -AmountToken $amountToken -TitleToken $bestLine
        }

        if ($mode -eq 'Conservative') {
            $tokenCount = 0
            foreach ($v in @($keyword,$dateToken,$amountToken,$bestLine)) {
                if (-not [string]::IsNullOrWhiteSpace([string]$v)) { $tokenCount++ }
            }
            if ($tokenCount -lt 2) {
                $smart = ''
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($smart)) {
            $baseName = $smart
            $reason = 'SmartTokens'
            $source = 'Content'
        }
    }

    if ([string]::IsNullOrWhiteSpace($baseName)) {
        $baseName = $fallback
        if ([string]::IsNullOrWhiteSpace($reason)) { $reason = 'FallbackOriginal' }
        if ([string]::IsNullOrWhiteSpace($source)) { $source = 'Original' }
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$Row.LogicalGroup)) {
        $baseName = '{0}_{1}' -f $baseName, $Row.LogicalGroup
    }

    $baseName = Get-SafeFileName $baseName
    if ([string]::IsNullOrWhiteSpace($baseName)) {
        $baseName = 'RecoveredFile'
    }

    $score = Get-NamingConfidenceScore -Keyword $keyword -DateToken $dateToken -AmountToken $amountToken -TitleToken $bestLine -LooksBad:$looksBad -Mode $mode -FallbackBase $fallback
    if ($source -eq 'Original' -and $score -gt 45) { $score = 45 }
    if ($source -eq 'Excel' -and $score -lt 85) { $score = 85 }
    if ($reason -eq 'RosterWorkbook' -and $score -lt 88) { $score = 88 }

    [pscustomobject]@{
        BaseName = $baseName
        Mode = $mode
        ConfidenceScore = $score
        ConfidenceLabel = (Get-NamingConfidenceLabel -Score $score)
        Reason = $reason
        Source = $source
        Keyword = $keyword
        DateToken = $dateToken
        AmountToken = $amountToken
        TitleToken = $bestLine
        LooksGarbled = $looksBad
    }
}

function Get-SuggestedBaseName {
    param($Row)
    $info = Get-SuggestedBaseNameInfo -Row $Row
    return [string]$info.BaseName
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

    $base = [IO.Path]::GetFileNameWithoutExtension($FileName)

    if ($FileName -match '^(file|doc|xls|ppt|recovered|found|chk|image|data|scan|dump|tmp)[-_]?\d+(\.[^.]+)?$') {
        return $true
    }

    if ($base -match '^[A-Za-z]{1,4}\d{3,}$') {
        return $true
    }

    if ($base -match '^\d{4,}$') {
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

        $nameInfo = $null
        if ($row.Role -eq 'Primary') {
            $nameInfo = Get-SuggestedBaseNameInfo -Row $row
            $baseName = $nameInfo.BaseName
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
            $nameInfo = Get-SuggestedBaseNameInfo -Row $row
            $baseName = $nameInfo.BaseName
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
            NamingMode    = $(if ($nameInfo) { $nameInfo.Mode } else { $script:NamingMode })
            NamingConfidence = $(if ($nameInfo) { $nameInfo.ConfidenceLabel } else { 'Low' })
            NamingConfidenceScore = $(if ($nameInfo) { $nameInfo.ConfidenceScore } else { 20 })
            NamingReason  = $(if ($nameInfo) { $nameInfo.Reason } else { 'DuplicateCounter' })
        })
    }

    return $plans
}

# -----------------------------
# Scan
# -----------------------------
function Start-Scan {
    Clear-Host
    Write-Host (T '掃描處理中，請稍後...' 'Scan is in progress, please wait...') -ForegroundColor Cyan
	Write-Host ""
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

    if ($script:KillOfficeProcessesBeforeScan) {
        Stop-OrphanOfficeProcesses -Force
    }

    $files = @()
    $exts = @('*.docx','*.xlsx','*.pptx','*.doc','*.xls','*.ppt','*.rtf','*.pdf','*.txt')

    foreach ($e in $exts) {
        try {
            $files += Get-ChildItem -LiteralPath $script:ScanRoot -Recurse -File -Filter $e -ErrorAction SilentlyContinue
        }
        catch {
        }
    }

    $files = $files | Sort-Object FullName -Unique

    if (-not $files -or $files.Count -eq 0) {
        Write-Host (T '找不到支援的檔案。' 'No supported files found.') -ForegroundColor Yellow
        Write-Host ((T '掃描路徑' 'Scan Root') + ': ' + $script:ScanRoot) -ForegroundColor DarkCyan
        Update-Status -Status (T '失敗' 'Failed') -Summary (T '沒有檔案' 'No files')
        Wait-Return
        return
    }

    Write-Host ((T '掃描路徑' 'Scan Root') + ': ' + $script:ScanRoot) -ForegroundColor Cyan
    Write-Host ((T '總檔案數' 'Total Files') + ': ' + $files.Count) -ForegroundColor Cyan
    Write-Host ''
    Start-ScanProgressUi

    $rows = @()
    $total = $files.Count

    if ($script:LegacyConversionMode) {
        Initialize-OfficeInterop
    }

    try {
        for ($i = 0; $i -lt $total; $i++) {
        $f = $files[$i]
        Show-ProgressLine -Current ($i + 1) -Total $total -FileName $f.Name

        if (Test-ScanCancelRequested) {
            $confirm = Show-ScanCancelConfirmUi -PromptText (T '偵測到 ESC，是否中止掃描？' 'ESC detected. Cancel scan?')
            Clear-ScanCancelConfirmUi

            if ($confirm) {
                Update-Status -Status (T '已中止' 'Cancelled') -Summary (T '使用者確認中止掃描。' 'User confirmed cancel.')
                break
            }
            else {
                $script:ScanCancelRequested = $false
                Show-ProgressLine -Current ($i + 1) -Total $total -FileName $f.Name
            }
        }

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
            ProtectionState = $contentInfo.ProtectionState
            ProtectionDetail = $contentInfo.ProtectionDetail
            ContentSource = $contentInfo.ContentSource
            NamingConfidence = $contentInfo.NamingConfidence
            LegacyConverted = $contentInfo.LegacyConverted
            ConvertedType = $contentInfo.ConvertedType
            ConversionStatus = $contentInfo.ConversionStatus
            LogicalGroup  = ''
            DuplicateType = ''
            Role          = ''
            RoleRank      = 0
            ScanTime      = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        })

        $rows += $row
    }
        Show-ProgressLine -Current $total -Total $total -FileName '' -Force
    }
    finally {
        Stop-ScanProgressUi
        if ($script:LegacyConversionMode) {
            Close-OfficeInterop
        }
    }

    if ($script:ScanCancelRequested) {
        Write-Host ''
        Write-Host (T '已中止掃描。已保留目前已完成的掃描結果。' 'Scan cancelled. Completed results have been kept.') -ForegroundColor Yellow
    }

    $rows = Apply-Grouping -Rows $rows
    $rows = Set-PrimaryAndDuplicateRoles -Rows $rows
    $script:Results = $rows

    $csvPath = Join-Path $script:OutputRoot ('OfficeRecovery_{0}.csv' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    $script:LastCsvReport = $csvPath

    $dupFileGroups = (@($rows | Group-Object FileHash | Where-Object { $_.Name -and $_.Count -gt 1 })).Count
    $dupContGroups = (@($rows | Group-Object ContentHash | Where-Object { $_.Name -and $_.Count -gt 1 })).Count
    $failCount = (@($rows | Where-Object { $_.ParseStatus -eq (T '解析失敗' 'Parse failed') })).Count

    Save-StateAndStatus -Status (T '完成' 'Done') -Summary ("CSV: $csvPath")

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
    Write-Host (T 'HTML 報表處理中，請稍後...' 'Generating HTML report, please wait...') -ForegroundColor Cyan
	Write-Host ""
    Save-StateAndStatus -Status (T '處理中' 'Processing') -Summary (T '正在輸出 HTML' 'Generating HTML')

    try {
        if (-not $script:Results -or @($script:Results).Count -eq 0) {
            Restore-ResultsFromLastCsv
        }

        if (-not $script:Results -or @($script:Results).Count -eq 0) {
            Write-Host (T '尚未有掃描結果。' 'No scan results yet.') -ForegroundColor Yellow
            Update-Status -Status (T '失敗' 'Failed') -Summary (T '尚未掃描' 'No scan results')
            Wait-Return
            return
        }

    Ensure-Folder $script:OutputRoot

    Update-Status -Status (T '處理中' 'Processing') -Summary (T 'HTML 報表處理中，請稍後...' 'Generating HTML report, please wait...')

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
        $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue

        $displayVersion = ''
        if ($cv) {
            if ($cv.DisplayVersion) {
                $displayVersion = [string]$cv.DisplayVersion
            }
            elseif ($cv.ReleaseId) {
                $displayVersion = [string]$cv.ReleaseId
            }
        }

        $buildText = ''
        if ($os.BuildNumber) {
            $buildText = [string]$os.BuildNumber
        }

        if (-not [string]::IsNullOrWhiteSpace($displayVersion) -and -not [string]::IsNullOrWhiteSpace($buildText)) {
            $osText = '{0} ({1}, Build {2}, {3})' -f $os.Caption, $os.Version, $buildText, $displayVersion
        }
        elseif (-not [string]::IsNullOrWhiteSpace($buildText)) {
            $osText = '{0} ({1}, Build {2})' -f $os.Caption, $os.Version, $buildText
        }
        else {
            $osText = '{0} ({1})' -f $os.Caption, $os.Version
        }
    }
    catch {
        $osText = [Environment]::OSVersion.VersionString
    }

    $detailRows = New-Object System.Text.StringBuilder
    foreach ($r in $script:Results) {
        $roleDisplay = $r.Role
        if ($r.Role -eq 'Primary') {
            $roleDisplay = '⭐ ' + (T '主檔' 'Primary')
        }
        elseif ($r.Role -eq 'Duplicate') {
            $roleDisplay = (T '重複檔' 'Duplicate')
        }
        elseif ($r.Role -eq 'Unique') {
            $roleDisplay = (T '唯一檔' 'Unique')
        }
        elseif ($r.Role -eq 'Broken') {
            $roleDisplay = (T '損毀檔' 'Broken')
        }

        [void]$detailRows.AppendLine('<tr>')
        [void]$detailRows.AppendLine('<td>' + (Get-SafeHtml $r.FileName) + '</td>')
        [void]$detailRows.AppendLine('<td>' + (Get-SafeHtml $r.ExtensionName) + '</td>')
        [void]$detailRows.AppendLine('<td style="text-align:right">' + (Get-SafeHtml ([string]$r.SizeKB)) + '</td>')
        [void]$detailRows.AppendLine('<td>' + (Get-SafeHtml $r.ParseStatus) + '</td>')
        [void]$detailRows.AppendLine('<td>' + (Get-SafeHtml $r.ProtectionState) + '</td>')
        [void]$detailRows.AppendLine('<td>' + (Get-SafeHtml $r.DuplicateType) + '</td>')
        [void]$detailRows.AppendLine('<td>' + (Get-SafeHtml $roleDisplay) + '</td>')
        [void]$detailRows.AppendLine('<td>' + (Get-SafeHtml $r.ContentSource) + '</td>')
        [void]$detailRows.AppendLine('<td>' + (Get-SafeHtml $r.NamingConfidence) + '</td>')
        [void]$detailRows.AppendLine('<td style="text-align:right">' + (Get-SafeHtml ([string](Get-FileQualityScore $r))) + '</td>')
        [void]$detailRows.AppendLine('<td>' + (Get-SafeHtml $r.LogicalGroup) + '</td>')
        [void]$detailRows.AppendLine('<td style="font-family:Consolas,monospace;word-break:break-all">' + (Get-SafeHtml $r.FileHash) + '</td>')
        [void]$detailRows.AppendLine('<td style="font-family:Consolas,monospace;word-break:break-all">' + (Get-SafeHtml $r.ContentHash) + '</td>')
        [void]$detailRows.AppendLine('<td class="preview-cell">' + (Get-SafeHtml $r.PreviewText) + '</td>')
        $reasonDisplay = Get-FriendlyParseReason -Reason $r.ParseReason -ConversionStatus $r.ConversionStatus
        [void]$detailRows.AppendLine('<td class="reason-cell">' + (Get-SafeHtml $reasonDisplay) + '</td>')
        [void]$detailRows.AppendLine('<td class="reason-cell">' + (Get-SafeHtml $r.ProtectionDetail) + '</td>')
        [void]$detailRows.AppendLine('</tr>')
    }

    $title = Get-SafeHtml (T 'Office 檔案救援分析報表 v5' 'Office Recovery Analysis Report v5')
    $summaryText = Get-SafeHtml (T '摘要' 'Summary')
    $detailText = Get-SafeHtml (T '明細' 'Details')
    $customerSummary = Get-SafeHtml (T '客戶報告摘要' 'Customer Summary')
    $subTitle = Get-SafeHtml (T '產品級救援分析報表' 'Product-grade recovery analysis report')
    $searchPlaceholder = Get-SafeHtml (T '搜尋' 'Search')
    $generatedBy = Get-SafeHtml (T '由 OfficeRecoveryToolkit.ps1 v5 產生' 'Generated by OfficeRecoveryToolkit.ps1 v5')

	function New-CardLabelHtml {
		param(
			[string]$Zh,
			[string]$En
		)
		return ('<div class="k-zh">{0}</div><div class="k-en">{1}</div>' -f (Get-SafeHtml $Zh), (Get-SafeHtml $En))
	}

	function New-ThLabelHtml {
		param(
			[string]$Zh,
			[string]$En
		)
		return ('<div class="th-zh">{0}</div><div class="th-en">{1}</div>' -f (Get-SafeHtml $Zh), (Get-SafeHtml $En))
	}

    $cardTotalFiles   = New-CardLabelHtml -Zh '總檔案數'   -En 'Total Files'
    $cardDupFileGrp   = New-CardLabelHtml -Zh '相同檔案群組' -En 'Duplicate File Groups'
    $cardDupContGrp   = New-CardLabelHtml -Zh '相同內容群組' -En 'Duplicate Content Groups'
    $cardParseFail    = New-CardLabelHtml -Zh '解析失敗'   -En 'Parse Failed'
    $cardPrimary      = New-CardLabelHtml -Zh '主檔'       -En 'Primary'
    $cardDuplicate    = New-CardLabelHtml -Zh '重複檔'     -En 'Duplicate'
    $cardUnique       = New-CardLabelHtml -Zh '唯一檔'     -En 'Unique'
    $cardNonParsed    = New-CardLabelHtml -Zh '非成功解析' -En 'Non-Parsed'
	$thFileName      = New-ThLabelHtml -Zh '檔名'       -En 'File Name'
    $thType          = New-ThLabelHtml -Zh '類型'       -En 'Type'
    $thSizeKB        = New-ThLabelHtml -Zh '大小(KB)'   -En 'Size (KB)'
    $thStatus        = New-ThLabelHtml -Zh '狀態'       -En 'Status'
    $thProtection    = New-ThLabelHtml -Zh '保護狀態'   -En 'Protection State'
    $thDupType       = New-ThLabelHtml -Zh '重複判定'   -En 'Duplicate Type'
    $thRole          = New-ThLabelHtml -Zh '角色'       -En 'Role'
    $thSource        = New-ThLabelHtml -Zh '內容來源'   -En 'Content Source'
    $thConfidence    = New-ThLabelHtml -Zh '命名信心'   -En 'Naming Confidence'
    $thQuality       = New-ThLabelHtml -Zh '品質分數'   -En 'Quality Score'
    $thGroup         = New-ThLabelHtml -Zh '群組'       -En 'Group'
    $thFileHash      = New-ThLabelHtml -Zh '檔案雜湊'   -En 'File Hash'
    $thContentHash   = New-ThLabelHtml -Zh '內容指紋'   -En 'Content Hash'
    $thPreview       = New-ThLabelHtml -Zh '內容預覽'   -En 'Preview Text'
    $thReason        = New-ThLabelHtml -Zh '說明'       -En 'Reason'
    $thProtectionDetail = New-ThLabelHtml -Zh '保護細節' -En 'Protection Detail'
	$thComputerName = New-ThLabelHtml -Zh '電腦名稱' -En 'Computer Name'
    $thOS           = New-ThLabelHtml -Zh '作業系統' -En 'Operating System'
    $thUser         = New-ThLabelHtml -Zh '使用者'   -En 'User'
    $thScanRoot     = New-ThLabelHtml -Zh '掃描路徑' -En 'Scan Root'
    $thReportTime   = New-ThLabelHtml -Zh '報表時間' -En 'Report Time'
	$thComputerName = New-ThLabelHtml -Zh '電腦名稱' -En 'Computer Name'
    $thOS           = New-ThLabelHtml -Zh '作業系統' -En 'Operating System'
    $thUser         = New-ThLabelHtml -Zh '使用者'   -En 'User'
    $thScanRoot     = New-ThLabelHtml -Zh '掃描路徑' -En 'Scan Root'
    $thReportTime   = New-ThLabelHtml -Zh '報表時間' -En 'Report Time'
	$searchPlaceholder = Get-SafeHtml (T '輸入關鍵字搜尋（檔名 / 群組 / 內容）' 'Search by file / group / content')
	$searchLabel = Get-SafeHtml (T '搜尋' 'Search')
	$clearLabel = Get-SafeHtml (T '清除' 'Clear')
	$searchStatAll = Get-SafeHtml (T '全部筆數' 'Total Rows')
	$searchStatMatched = Get-SafeHtml (T '符合筆數' 'Matched Rows')

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
.card .k-zh{font-size:14px;color:#334155;font-weight:600;line-height:1.25}
.card .k-en{font-size:11px;color:#64748b;line-height:1.2;margin-top:2px}
.card .v{
    font-size:30px;
    font-weight:700;
    margin-top:10px;
    color:#0f172a;
}
.panel{background:#fff;border-radius:16px;box-shadow:0 4px 16px rgba(0,0,0,.08);padding:18px;margin-bottom:20px;overflow:visible}
.table-wrap{width:100%;max-width:100%;overflow:auto;-webkit-overflow-scrolling:touch;border:1px solid #dbe4f0;border-radius:12px;background:#fff}
.table-wrap table{width:max-content;min-width:2200px;border-collapse:separate;border-spacing:0;table-layout:auto}
th,td{border-right:1px solid #dbe4f0;border-bottom:1px solid #dbe4f0;padding:8px 10px;vertical-align:top;text-align:left;font-size:13px;overflow-wrap:anywhere;word-break:break-word;background:#fff}
th:last-child,td:last-child{border-right:none}
thead th{position:sticky;top:0;z-index:2;background:#eaf2ff}
.th-zh{font-size:13px;font-weight:700;line-height:1.2;color:#1e293b}
.th-en{font-size:10px;font-weight:400;line-height:1.15;color:#64748b;margin-top:2px}
th{white-space:normal;min-width:120px}
.preview-cell{min-width:320px;max-width:520px}
.reason-cell{min-width:220px;max-width:360px;color:#7c2d12;background:#fff7ed}
.search-wrap{
    display:flex;
    align-items:center;
    gap:10px;
    width:100%;
    max-width:720px;
    margin:10px auto 16px auto;
    flex-wrap:wrap;
}

.search-box{
    display:flex;
    align-items:center;
    flex:1 1 520px;
    min-width:320px;
    border:1px solid #cbd5e1;
    border-radius:10px;
    background:#fff;
    overflow:hidden;
}

.search-icon{
    padding:0 12px;
    color:#64748b;
    font-size:14px;
    user-select:none;
}

.search{
    flex:1 1 auto;
    min-width:120px;
    padding:10px 12px 10px 0;
    border:none;
    font-size:14px;
    background:#fff;
}

.search:focus{
    outline:none;
}

.search-box:focus-within{
    border-color:#3b82f6;
    box-shadow:0 0 0 2px rgba(59,130,246,0.2);
}

.search-clear{
    border:1px solid #cbd5e1;
    background:#fff;
    color:#334155;
    border-radius:10px;
    padding:10px 14px;
    font-size:13px;
    cursor:pointer;
}

.search-clear:hover{
    background:#f8fafc;
}

.search-stat{
    font-size:12px;
    color:#64748b;
    margin:0 0 14px 0;
}
.reason-cell{max-width:220px;overflow-wrap:anywhere;word-break:break-word}
.preview-cell{max-width:320px;overflow-wrap:anywhere;word-break:break-word}
.footer{margin-top:24px;color:#64748b;font-size:12px}
.small{font-size:13px;color:#475569}
</style>
<script>
function updateSearchStats(total, matched) {
  var el = document.getElementById("searchStat");
  if (!el) return;
  var totalLabel = el.getAttribute("data-total-label") || "Total Rows";
  var matchedLabel = el.getAttribute("data-matched-label") || "Matched Rows";
  el.innerText = matchedLabel + ": " + matched + " / " + totalLabel + ": " + total;
}

function filterTable() {
  var input = document.getElementById("searchBox");
  var filter = (input.value || "").toLowerCase();
  var rows = document.querySelectorAll("#detailBody tr");
  var matched = 0;
  var total = rows.length;

  for (var i = 0; i < rows.length; i++) {
    var txt = (rows[i].innerText || "").toLowerCase();
    var hit = txt.indexOf(filter) > -1;
    rows[i].style.display = hit ? "" : "none";
    if (hit) matched++;
  }

  updateSearchStats(total, matched);
}

function clearSearch() {
  var input = document.getElementById("searchBox");
  if (!input) return;
  input.value = "";
  filterTable();
  input.focus();
}

window.addEventListener("load", function() {
  filterTable();
});
</script>
</head>
<body>
<div class="wrap">
    <h1>$title</h1>
    <div class="sub">$subTitle</div>

	<div class="grid">
        <div class="card">
            $cardTotalFiles
            <div class="v">$totalFiles</div>
        </div>
        <div class="card">
            $cardDupFileGrp
            <div class="v">$dupFileGroups</div>
        </div>
        <div class="card">
            $cardDupContGrp
            <div class="v">$dupContGroups</div>
        </div>
        <div class="card">
            $cardParseFail
            <div class="v">$failCount</div>
        </div>
    </div>

    <div class="grid">
        <div class="card">
            $cardPrimary
            <div class="v">$primary</div>
        </div>
        <div class="card">
            $cardDuplicate
            <div class="v">$dup</div>
        </div>
        <div class="card">
            $cardUnique
            <div class="v">$unique</div>
        </div>
        <div class="card">
            $cardNonParsed
            <div class="v">$fail</div>
        </div>
    </div>

    <div class="panel">
        <h2>$customerSummary</h2>
        <div class="small">
            $(Get-SafeHtml (T '總檔案' 'Total Files')): $totalFiles<br>
            $(Get-SafeHtml (T '成功解析' 'Parsed Successfully')): $(($totalFiles - $fail))<br>
            $(Get-SafeHtml (T '主檔' 'Primary')): $primary<br>
            $(Get-SafeHtml (T '重複檔' 'Duplicate')): $dup<br>
            $(Get-SafeHtml (T '唯一檔' 'Unique')): $unique<br>
            $(Get-SafeHtml (T '無法完整解析' 'Not Fully Parsed')): $fail
        </div>
    </div>

   <div class="panel">
		<h2>$summaryText</h2>
        <div class="table-wrap">
		<table>
			<tr>
				<th>$thComputerName</th>
				<td>$(Get-SafeHtml $env:COMPUTERNAME)</td>
			</tr>
			<tr>
				<th>$thOS</th>
				<td>$(Get-SafeHtml $osText)</td>
			</tr>
			<tr>
				<th>$thUser</th>
				<td>$(Get-SafeHtml $env:USERNAME)</td>
			</tr>
			<tr>
				<th>$thScanRoot</th>
				<td>$(Get-SafeHtml $script:ScanRoot)</td>
			</tr>
			<tr>
				<th>$thReportTime</th>
				<td>$(Get-SafeHtml ((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')))</td>
			</tr>
		</table>
        </div>
	</div>

    <div class="panel">
		<h2>$detailText</h2>
		<div class="search-wrap">
			<div class="search-box">
				<div class="search-icon">🔍</div>
				<input type="text" id="searchBox" class="search" onkeyup="filterTable()" placeholder="$searchPlaceholder">
			</div>
			<button type="button" class="search-clear" onclick="clearSearch()">$clearLabel</button>
		</div>
		<div id="searchStat" class="search-stat" data-total-label="$searchStatAll" data-matched-label="$searchStatMatched"></div>
        <div class="table-wrap">
		<table>
			<thead>
                <tr>
                    <th>$thFileName</th>
                    <th>$thType</th>
                    <th>$thSizeKB</th>
                    <th>$thStatus</th>
                    <th>$thProtection</th>
                    <th>$thDupType</th>
                    <th>$thRole</th>
                    <th>$thSource</th>
                    <th>$thConfidence</th>
                    <th>$thQuality</th>
                    <th>$thGroup</th>
                    <th>$thFileHash</th>
                    <th>$thContentHash</th>
                    <th>$thPreview</th>
                    <th>$thReason</th>
                    <th>$thProtectionDetail</th>
                </tr>
            </thead>
            <tbody id="detailBody">
                $($detailRows.ToString())
            </tbody>
        </table>
        </div>
        </div>
    </div>

    <div class="footer">$generatedBy</div>
</div>
</body>
</html>
"@

    [IO.File]::WriteAllText($htmlPath, $html, [Text.Encoding]::UTF8)
    $script:LastHtmlReport = $htmlPath

    Save-StateAndStatus -Status (T '完成' 'Done') -Summary ("HTML: $htmlPath")

    Write-Host (T 'HTML 報表已輸出。' 'HTML report exported.') -ForegroundColor Green
    Write-Host $htmlPath -ForegroundColor Green
    Write-Host ''
    Write-Host (T '之後可按 [7] 用系統預設瀏覽器開啟最新 HTML 報表。' 'You can press [7] later to open the latest HTML report with the default browser.') -ForegroundColor Cyan
    }
    catch {
        $errMsg = $_.Exception.Message
        if ([string]::IsNullOrWhiteSpace($errMsg)) {
            $errMsg = (T '未知錯誤' 'Unknown error')
        }

        Write-Host (T 'HTML 報表輸出失敗。' 'Failed to export HTML report.') -ForegroundColor Red
        Write-Host $errMsg -ForegroundColor Yellow
        Update-Status -Status (T '失敗' 'Failed') -Summary ((T 'HTML 輸出失敗' 'HTML export failed') + ': ' + $errMsg)
    }

    Wait-Return
}

function Open-LatestHtmlReport {
    Clear-Host
    Reset-UiCachesSafe

    $choices = @()

    if (-not [string]::IsNullOrWhiteSpace($script:LastHtmlReport) -and (Test-Path -LiteralPath $script:LastHtmlReport)) {
        $choices += [pscustomobject]@{
            Key  = '1'
            Name = (T '最新正式 HTML 報表' 'Latest Main HTML Report')
            Path = $script:LastHtmlReport
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($script:LastRenamePreviewHtml) -and (Test-Path -LiteralPath $script:LastRenamePreviewHtml)) {
        $choices += [pscustomobject]@{
            Key  = '2'
            Name = (T '最新模擬改名 HTML 報表' 'Latest Rename Preview HTML Report')
            Path = $script:LastRenamePreviewHtml
        }
    }

    if (-not $choices -or $choices.Count -eq 0) {
        Write-Host (T '尚未找到可開啟的 HTML 報表。' 'No HTML reports are available to open yet.') -ForegroundColor Yellow
        Update-Status -Status (T '失敗' 'Failed') -Summary (T '尚未匯出任何 HTML' 'No HTML reports exported yet')
        Wait-Return
        return
    }

    Write-Host (T '請選擇要開啟的 HTML 報表：' 'Choose which HTML report to open:') -ForegroundColor Cyan
    Write-Host ''
    foreach ($item in $choices) {
        Write-Host ('[' + $item.Key + '] ' + $item.Name) -ForegroundColor White
        Write-Host ('    ' + $item.Path) -ForegroundColor DarkGray
    }
    Write-Host ('[Esc/Enter] ' + (T '取消' 'Cancel')) -ForegroundColor DarkGray
    Write-Host ''

    $selected = $null
    while (-not $selected) {
        $key = [Console]::ReadKey($true)
        if ($key.Key -eq 'Escape' -or $key.Key -eq 'Enter') {
            Update-Status -Status (T '完成' 'Done') -Summary (T '取消開啟 HTML' 'Open HTML cancelled')
            return
        }

        $char = $key.KeyChar.ToString()
        $selected = @($choices | Where-Object { $_.Key -eq $char } | Select-Object -First 1)
    }

    try {
        Start-Process -FilePath $selected.Path | Out-Null
        Update-Status -Status (T '完成' 'Done') -Summary $selected.Path
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

function Get-RenamePlanWithFallback {
    param([array]$Rows)

    $plans = Get-RenamePlan -Rows $Rows -OnlyGenericNames
    $mode = 'generic'

    if (-not $plans -or @($plans).Count -eq 0) {
        $plans = Get-RenamePlan -Rows $Rows
        $mode = 'all'
    }

    return New-Object PSObject -Property @{
        Plans = @($plans)
        Mode  = $mode
    }
}

function New-RenamePreviewDataset {
    param(
        [array]$Rows,
        [array]$Plans
    )

    $planMap = @{}
    foreach ($plan in @($Plans)) {
        if (-not [string]::IsNullOrWhiteSpace($plan.OriginalPath)) {
            $planMap[$plan.OriginalPath.ToLowerInvariant()] = $plan
        }
    }

    $dataset = @()
    foreach ($row in @($Rows)) {
        $action = ''
        $suggestedName = ''
        $reason = ''

        $key = ''
        if (-not [string]::IsNullOrWhiteSpace($row.FullPath)) {
            $key = $row.FullPath.ToLowerInvariant()
        }

        if ($key -and $planMap.ContainsKey($key)) {
            $plan = $planMap[$key]
            $action = T '建議改名' 'Rename Suggested'
            $suggestedName = [string]$plan.SuggestedName
            $reason = T '符合自動改名條件' 'Matched auto-rename rules'
        }
        else {
            $suggestedName = [string]$row.FileName

            if ($row.Role -eq 'Broken') {
                $action = T '略過' 'Skipped'
                $reason = T '解析失敗或損毀檔案' 'Parse failed or broken file'
            }
            else {
                $action = T '不變更' 'No Change'
                $reason = T '目前無需改名' 'No rename needed'
            }
        }

        $previewLines = @()
        if (-not [string]::IsNullOrWhiteSpace([string]$row.PreviewText)) {
            $previewLines = ([string]$row.PreviewText -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        }

        $dataset += New-Object PSObject -Property ([ordered]@{
            Action         = $action
            Role           = $row.Role
            OriginalName   = $row.FileName
            SuggestedName  = $suggestedName
            Extension      = $row.Extension
            LogicalGroup   = $row.LogicalGroup
            NamingMode     = $(if ($plan) { [string]$plan.NamingMode } else { [string]$script:NamingMode })
            NamingConfidence = $(if ($plan) { [string]$plan.NamingConfidence } else { 'Low' })
            NamingConfidenceScore = $(if ($plan) { [string]$plan.NamingConfidenceScore } else { '20' })
            Reason         = $reason
            PreviewLine1   = $(if ($previewLines.Count -ge 1) { $previewLines[0] } else { '' })
            PreviewLine2   = $(if ($previewLines.Count -ge 2) { $previewLines[1] } else { '' })
            PreviewLine3   = $(if ($previewLines.Count -ge 3) { $previewLines[2] } else { '' })
            PreviewText    = ([string]$row.PreviewText)
            OriginalPath   = $row.FullPath
        })
    }

    return @($dataset)
}

function Preview-RenamePlan {
    Clear-Host

    Write-Host ''
    Write-Host (T '模擬改名處理中，請稍後...' 'Rename preview in progress, please wait...') -ForegroundColor Cyan

    try {
        if (-not $script:Results -or @($script:Results).Count -eq 0) {
            Restore-ResultsFromLastCsv
        }

        if (-not $script:Results -or @($script:Results).Count -eq 0) {
            Write-Host (T '尚未有掃描結果。' 'No scan results yet.') -ForegroundColor Yellow
            Update-Status -Status (T '失敗' 'Failed') -Summary (T '尚未掃描' 'No scan results')
            Wait-Return
            return
        }

        $renameInfo = Get-RenamePlanWithFallback -Rows $script:Results
        $plans = @($renameInfo.Plans)
        $previewRows = @(New-RenamePreviewDataset -Rows $script:Results -Plans $plans)

        if (-not $previewRows -or $previewRows.Count -eq 0) {
            Write-Host (T '沒有可輸出的模擬改名資料。' 'No rename preview data to export.') -ForegroundColor Yellow
            Save-StateAndStatus -Status (T '完成' 'Done') -Summary (T '沒有模擬改名資料' 'No rename preview data')
            Wait-Return
            return
        }

        Ensure-Folder $script:OutputRoot
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $csvPath = Join-Path $script:OutputRoot ('RenamePreview_{0}.csv' -f $stamp)
        $htmlPath = Join-Path $script:OutputRoot ('RenamePreview_{0}.html' -f $stamp)

        $script:LastRenamePreviewCsv = $csvPath
        $script:LastRenamePreviewHtml = $htmlPath

        $previewRows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

        if ($renameInfo.Mode -eq 'all') {
            Write-Host (T '未找到典型救援檔名，已改為依所有掃描結果產生預覽。' 'No typical recovered generic names were found, preview generated from all scanned files.') -ForegroundColor Yellow
            Write-Host ''
        }

        $rows = New-Object System.Text.StringBuilder
        foreach ($p in $previewRows) {
            [void]$rows.AppendLine('<tr>')
            [void]$rows.AppendLine('<td>' + (Get-SafeHtml $p.Action) + '</td>')
            [void]$rows.AppendLine('<td>' + (Get-SafeHtml $p.Role) + '</td>')
            [void]$rows.AppendLine('<td>' + (Get-SafeHtml $p.OriginalName) + '</td>')
            [void]$rows.AppendLine('<td>' + (Get-SafeHtml $p.SuggestedName) + '</td>')
            [void]$rows.AppendLine('<td>' + (Get-SafeHtml $p.Extension) + '</td>')
            [void]$rows.AppendLine('<td>' + (Get-SafeHtml $p.LogicalGroup) + '</td>')
            [void]$rows.AppendLine('<td>' + (Get-SafeHtml $p.NamingMode) + '</td>')
            [void]$rows.AppendLine('<td>' + (Get-SafeHtml ([string]$p.NamingConfidence + ' (' + [string]$p.NamingConfidenceScore + ')')) + '</td>')
            [void]$rows.AppendLine('<td>' + (Get-SafeHtml $p.Reason) + '</td>')
            [void]$rows.AppendLine('<td>' + (Get-SafeHtml $p.PreviewLine1) + '</td>')
            [void]$rows.AppendLine('<td>' + (Get-SafeHtml $p.PreviewLine2) + '</td>')
            [void]$rows.AppendLine('<td>' + (Get-SafeHtml $p.PreviewLine3) + '</td>')
            [void]$rows.AppendLine('<td style="overflow-wrap:anywhere;word-break:break-word;">' + (Get-SafeHtml $p.OriginalPath) + '</td>')
            [void]$rows.AppendLine('</tr>')
        }

        $html = @"
<!DOCTYPE html>
<html lang="$($script:Lang)">
<head>
<meta charset="utf-8" />
<title>$(Get-SafeHtml (T '模擬改名預覽報表' 'Rename Preview Report'))</title>
<style>
body{font-family:"Segoe UI","Microsoft JhengHei",Arial,sans-serif;background:#f3f6fb;color:#1f2937;margin:0;padding:24px}
.wrap{max-width:1400px;margin:0 auto}
.panel{background:#fff;border-radius:16px;box-shadow:0 4px 16px rgba(0,0,0,.08);padding:18px}
.table-wrap{width:100%;max-width:100%;overflow:auto;border:1px solid #dbe4f0;border-radius:12px;background:#fff}
.table-wrap table{width:100%;min-width:1400px;border-collapse:separate;border-spacing:0;table-layout:auto}
th,td{border-right:1px solid #dbe4f0;border-bottom:1px solid #dbe4f0;padding:8px 10px;vertical-align:top;text-align:left;font-size:13px;overflow-wrap:anywhere;word-break:break-word;background:#fff}
th:last-child,td:last-child{border-right:none}
thead th{position:sticky;top:0;z-index:2;background:#eaf2ff}
th{min-width:120px}
</style>
</head>
<body>
<div class="wrap">
    <div class="panel">
        <h1>$(Get-SafeHtml (T '模擬改名預覽報表' 'Rename Preview Report'))</h1>
        <p>$(Get-SafeHtml (T '此報表已包含全部掃描結果；建議改名、不變更與略過項目都會列出。內容預覽已拆成多欄顯示，避免全部資訊擠在同一格。' 'This preview includes all scanned results. Preview text has been split into multiple columns so the information is easier to read.'))</p>
        <div class="table-wrap">
        <table>
            <thead>
                <tr>
                    <th>$(Get-SafeHtml (T '動作' 'Action'))</th>
                    <th>$(Get-SafeHtml (T '角色' 'Role'))</th>
                    <th>$(Get-SafeHtml (T '原始檔名' 'Original Name'))</th>
                    <th>$(Get-SafeHtml (T '建議檔名' 'Suggested Name'))</th>
                    <th>$(Get-SafeHtml (T '副檔名' 'Extension'))</th>
                    <th>$(Get-SafeHtml (T '群組' 'Group'))</th>
                    <th>$(Get-SafeHtml (T '命名模式' 'Naming Mode'))</th>
                    <th>$(Get-SafeHtml (T '命名信心' 'Naming Confidence'))</th>
                    <th>$(Get-SafeHtml (T '原因' 'Reason'))</th>
                    <th>$(Get-SafeHtml (T '預覽 1' 'Preview 1'))</th>
                    <th>$(Get-SafeHtml (T '預覽 2' 'Preview 2'))</th>
                    <th>$(Get-SafeHtml (T '預覽 3' 'Preview 3'))</th>
                    <th>$(Get-SafeHtml (T '原始路徑' 'Original Path'))</th>
                </tr>
            </thead>
            <tbody>
                $($rows.ToString())
            </tbody>
        </table>
        </div>
    </div>
</div>
</body>
</html>
"@

        [IO.File]::WriteAllText($htmlPath, $html, [Text.Encoding]::UTF8)

        Save-StateAndStatus -Status (T '完成' 'Done') -Summary ("Rename Preview HTML: $htmlPath")

        Write-Host (T '模擬改名預覽已輸出。' 'Rename preview exported.') -ForegroundColor Green
        Write-Host ((T '總筆數' 'Total Rows') + ' : ' + $previewRows.Count) -ForegroundColor Green
        Write-Host ((T '建議改名' 'Rename Suggested') + ' : ' + @($previewRows | Where-Object { $_.Action -eq (T '建議改名' 'Rename Suggested') }).Count) -ForegroundColor Cyan
        Write-Host ('CSV  : ' + $csvPath) -ForegroundColor Green
        Write-Host ('HTML : ' + $htmlPath) -ForegroundColor Green
        Wait-Return
    }
    catch {
        Write-Host (T '模擬改名預覽失敗。' 'Rename preview failed.') -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Yellow
        Save-StateAndStatus -Status (T '失敗' 'Failed') -Summary (T '模擬改名失敗' 'Rename preview failed')
        Wait-Return
    }
}

function Invoke-AutoRename {
    Clear-Host

    try {
        if (-not $script:Results -or @($script:Results).Count -eq 0) {
            Restore-ResultsFromLastCsv
        }

        if (-not $script:Results -or @($script:Results).Count -eq 0) {
            Write-Host (T '尚未有掃描結果。' 'No scan results yet.') -ForegroundColor Yellow
            Update-Status -Status (T '失敗' 'Failed') -Summary (T '尚未掃描' 'No scan results')
            Wait-Return
            return
        }

        $renameInfo = Get-RenamePlanWithFallback -Rows $script:Results
        $plans = @($renameInfo.Plans)

        if (-not $plans -or $plans.Count -eq 0) {
            Write-Host (T '沒有可自動改名的檔案。' 'No files available for automatic renaming.') -ForegroundColor Yellow
            Update-Status -Status (T '完成' 'Done') -Summary (T '沒有可改名項目' 'No rename candidates')
            Wait-Return
            return
        }

        if ($renameInfo.Mode -eq 'all') {
            Write-Host (T '未找到典型救援檔名，將對所有可建議改名的檔案執行改名。' 'No typical recovered generic names were found; renaming will apply to all renameable files.') -ForegroundColor Yellow
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
                    NamingMode    = $p.NamingMode
                    NamingConfidence = $p.NamingConfidence
                    NamingConfidenceScore = $p.NamingConfidenceScore
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
                    NamingMode    = $p.NamingMode
                    NamingConfidence = $p.NamingConfidence
                    NamingConfidenceScore = $p.NamingConfidenceScore
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

        $script:LastRenameLog = $csvPath
        Save-StateAndStatus -Status (T '完成' 'Done') -Summary ("Rename Log CSV: $csvPath")

        if ($success -gt 0) {
            $currentRoot = $script:ScanRoot
            Start-Scan
            $script:ScanRoot = $currentRoot
        }
        else {
            Wait-Return
        }
    }
    catch {
        Write-Host (T '自動改名流程發生錯誤。' 'Automatic rename failed.') -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Yellow
        Save-StateAndStatus -Status (T '失敗' 'Failed') -Summary (T '自動改名失敗' 'Automatic rename failed')
        Wait-Return
    }
    finally {
        $script:IsRenameInProgress = $false
    }
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
    Save-StateAndStatus -Status (T '完成' 'Done') -Summary ("Organize Log CSV: $csv")

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
    Save-AppState
}

function Toggle-PrimaryOnly {
    $script:OrganizePrimaryOnly = -not $script:OrganizePrimaryOnly
    Save-AppState
}


# -----------------------------
# Reset UiCaches
# -----------------------------
function Reset-UiCachesSafe {
    try {
        $script:LastRightPanelLines = @()
    } catch {}

    try {
        $script:LastStatusBarLines = @()
    } catch {}

    try {
        $script:LastMenuRenderWidth = 0
    } catch {}

    try {
        $script:UiCache.SettingsLines = @()
        $script:UiCache.StatusLines = @()
        $script:UiCache.WindowWidth = 0
        $script:UiCache.WindowHeight = 0
    } catch {}

    try {
        $script:ForceFullRedraw = $true
    } catch {}
}

# -----------------------------
# Set scan root
# -----------------------------
function Set-ScanRoot {
    Clear-Host
    Write-Host (T '請輸入掃描資料夾路徑：' 'Enter scan folder path:') -ForegroundColor Cyan
    $inputPath = Read-Host

    if ([string]::IsNullOrWhiteSpace($inputPath)) {
        Save-StateAndStatus -Status (T '就緒' 'Ready') -Summary $script:ScanRoot

        # 清除 UI 快取，避免右側資訊區消失
        Reset-UiCachesSafe
        return
    }

    if (Test-Path -LiteralPath $inputPath) {
        $script:ScanRoot = $inputPath
        Save-StateAndStatus -Status (T '完成' 'Done') -Summary $script:ScanRoot

        # 很重要：清掉快取，讓主選單回來時完整重畫
        Reset-UiCachesSafe
    }
    else {
        Write-Host (T '資料夾不存在。' 'Folder does not exist.') -ForegroundColor Red
        Save-StateAndStatus -Status (T '失敗' 'Failed') -Summary (T '資料夾不存在' 'Folder does not exist')

        Reset-UiCachesSafe
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
Load-AppState
Restore-ResultsFromLastCsv
Ensure-Folder $script:OutputRoot
Update-Status -Status (T '就緒' 'Ready') -Summary $script:LastSummary
Save-AppState

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

        if (Test-KeyMatch -KeyInfo $key -KeyName 'C' -VirtualKeyCode 67 -Chars @('c','C')) {
            Toggle-LegacyConversionMode
            $needsFullRedraw = $true
            continue
        }

        if (Test-KeyMatch -KeyInfo $key -KeyName 'H' -VirtualKeyCode 72 -Chars @('h','H')) {
            Toggle-NamingMode
            $needsFullRedraw = $true
            continue
        }

        if (Test-KeyMatch -KeyInfo $key -KeyName 'L' -VirtualKeyCode 76 -Chars @('l','L')) {
            if ($script:Lang -eq 'zh-TW') {
                $script:Lang = 'en-US'
            }
            else {
                $script:Lang = 'zh-TW'
            }
            Save-StateAndStatus -Status (T '就緒' 'Ready') -Summary ((T '語系已切換' 'Language switched') + ': ' + $script:Lang)
            $needsFullRedraw = $true
            continue
        }

        $action = $null

        if (Test-KeyMatch -KeyInfo $key -KeyName 'Enter' -VirtualKeyCode 13) {
            $action = $menuItems[$script:SelectedMenu].Action
        }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D1' -VirtualKeyCode 49 -Chars @('1')) { $old = $script:SelectedMenu; $script:SelectedMenu = 0; if ($old -ne $script:SelectedMenu) { Update-LightBarSelection -OldIndex $old -NewIndex $script:SelectedMenu }; $action = 'Scan' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D2' -VirtualKeyCode 50 -Chars @('2')) { $old = $script:SelectedMenu; $script:SelectedMenu = 1; if ($old -ne $script:SelectedMenu) { Update-LightBarSelection -OldIndex $old -NewIndex $script:SelectedMenu }; $action = 'ExportHtml' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D3' -VirtualKeyCode 51 -Chars @('3')) { $old = $script:SelectedMenu; $script:SelectedMenu = 2; if ($old -ne $script:SelectedMenu) { Update-LightBarSelection -OldIndex $old -NewIndex $script:SelectedMenu }; $action = 'OpenOutput' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D4' -VirtualKeyCode 52 -Chars @('4')) { $old = $script:SelectedMenu; $script:SelectedMenu = 3; if ($old -ne $script:SelectedMenu) { Update-LightBarSelection -OldIndex $old -NewIndex $script:SelectedMenu }; $action = 'SetScanRoot' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D5' -VirtualKeyCode 53 -Chars @('5')) { $old = $script:SelectedMenu; $script:SelectedMenu = 4; if ($old -ne $script:SelectedMenu) { Update-LightBarSelection -OldIndex $old -NewIndex $script:SelectedMenu }; $action = 'PreviewRename' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D6' -VirtualKeyCode 54 -Chars @('6')) { $old = $script:SelectedMenu; $script:SelectedMenu = 5; if ($old -ne $script:SelectedMenu) { Update-LightBarSelection -OldIndex $old -NewIndex $script:SelectedMenu }; $action = 'ApplyRename' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D7' -VirtualKeyCode 55 -Chars @('7')) { $old = $script:SelectedMenu; $script:SelectedMenu = 6; if ($old -ne $script:SelectedMenu) { Update-LightBarSelection -OldIndex $old -NewIndex $script:SelectedMenu }; $action = 'OpenHtml' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D8' -VirtualKeyCode 56 -Chars @('8')) { $old = $script:SelectedMenu; $script:SelectedMenu = 7; if ($old -ne $script:SelectedMenu) { Update-LightBarSelection -OldIndex $old -NewIndex $script:SelectedMenu }; $action = 'Organize' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D9' -VirtualKeyCode 57 -Chars @('9')) { $old = $script:SelectedMenu; $script:SelectedMenu = 8; if ($old -ne $script:SelectedMenu) { Update-LightBarSelection -OldIndex $old -NewIndex $script:SelectedMenu }; $action = 'ToggleMode' }
        elseif (Test-KeyMatch -KeyInfo $key -KeyName 'D0' -VirtualKeyCode 48 -Chars @('0')) { $old = $script:SelectedMenu; $script:SelectedMenu = 9; if ($old -ne $script:SelectedMenu) { Update-LightBarSelection -OldIndex $old -NewIndex $script:SelectedMenu }; $action = 'TogglePrimaryOnly' }
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
                Save-StateAndStatus -Status (T '完成' 'Done') -Summary ((T '整理模式已切換為' 'Organize mode switched to') + ': ' + $script:OrganizeMode)
                Draw-StatusBar
                Draw-SettingsPanel
                Draw-OneMenuItem -Index $script:SelectedMenu -Selected
                continue
            }
            'TogglePrimaryOnly' {
                Toggle-PrimaryOnly
                Save-StateAndStatus -Status (T '完成' 'Done') -Summary ((T '只整理主檔已切換為' 'Primary only switched to') + ': ' + $script:OrganizePrimaryOnly)
                Draw-StatusBar
                Draw-SettingsPanel
                Draw-OneMenuItem -Index $script:SelectedMenu -Selected
                continue
            }
            'ToggleLegacyConversion' {
                Toggle-LegacyConversionMode
                $needsFullRedraw = $true
                continue
            }
            'ToggleNamingMode' {
                Toggle-NamingMode
                $needsFullRedraw = $true
                continue
            }
            'Exit' { return }
        }
    }
}
finally {
    Save-AppState
    [Console]::CursorVisible = $true
    Clear-Host
}
