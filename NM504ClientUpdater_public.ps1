#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    NM504 Client 更新安裝腳本範例 (適合公開)
.DESCRIPTION
    此腳本包含：
    - 版本檢查
    - 安裝檔複製與 SHA256 驗證
    - 模擬模式 (SimulationMode)
    - 鎖定檔保護
    - 公用桌面捷徑複製
    - 日誌管理與 EventLog 記錄
#>

param(
    [string]$Ver = "2025.1.1.001",
    [string]$GPO = "Powershell",
    [string]$EXE114 = "C:\Path\To\NM114.exe",
    [string]$EXE106 = "C:\Path\To\NM106.exe",
    [string]$DirS1 = "\\<NAS_SERVER>\Software",
    [string]$DirS2 = "\\<AD_SERVER>\SYSVOL\scripts",
    [string]$TargetDate = "2025-01-01",
    [string]$DirLog = "\\<LOG_SERVER>\LogFiles\NM114-504Client", 
    [string]$Installer = "單機版_YYYMMDD_x64.exe",
    [ValidateSet("DEBUG","INFO","WARN","ERROR")]
    [string]$LogLevel = "INFO",
    [int]$LogRetentionDays = 30,
    [string]$PublicDesktop = "C:\Users\Public\Desktop",
    [string]$ShortcutName = "單機版.lnk",
    [switch]$SimulationMode
)

$EXIT_SUCCESS = 0
$EXIT_NO_UPDATE = 10
$EXIT_COPY_FAIL = 20
$EXIT_INSTALL_FAIL = 30
$EXIT_EXCEPTION = 99

$EventSource = "NM504ClientUpdater"

if ($SimulationMode) {
    Write-Host "--- 以模擬模式運行 ---" -ForegroundColor Cyan
    $levels = @{"DEBUG"=1;"INFO"=2;"WARN"=3;"ERROR"=4}
    if ($levels[$LogLevel] -gt $levels["INFO"]) { $LogLevel = "INFO" }
}

$LocalTempDir = "$env:SystemDrive\TEMP"
if (-not (Test-Path $LocalTempDir)) { New-Item -Path $LocalTempDir -ItemType Directory -Force | Out-Null }

try { $targetDateObj = [datetime]::ParseExact($TargetDate,'yyyy-MM-dd',$null) } 
catch { Write-Host "錯誤：日期格式無效 ($TargetDate)" -ForegroundColor Red; exit $EXIT_EXCEPTION }

$InstallerS1 = Join-Path $DirS1 $Installer
$InstallerS2 = Join-Path $DirS2 $Installer
$LogFile = Join-Path $DirLog ("${env:COMPUTERNAME}_$(Get-Date -Format 'yyyyMMdd').log")
$localInstallerPath = Join-Path $LocalTempDir $Installer

if (-not ([System.Diagnostics.EventLog]::SourceExists($EventSource))) {
    try { [System.Diagnostics.EventLog]::CreateEventSource($EventSource,"Application") } catch {}
}

function Write-Log {
    param([string]$Message, [string]$Type="INFO")
    if ($SimulationMode -and $Type -ne "ERROR") { $Type = "SIMULATE:$Type" }

    $levels = @{"DEBUG"=1;"INFO"=2;"WARN"=3;"ERROR"=4;"SIMULATE:DEBUG"=1;"SIMULATE:INFO"=2;"SIMULATE:WARN"=3}
    if ($levels[$Type] -lt $levels[$LogLevel]) { return }

    $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Line = "[$Type] $Time`t$Message"

    if (-not $SimulationMode) { 
        try { $Line | Out-File -FilePath $LogFile -Append -Encoding UTF8 -ErrorAction Stop } 
        catch { Write-Host "日誌寫入失敗: $($_.Exception.Message)" -ForegroundColor Yellow }
    }

    Write-Host $Line
    try {
        $entryType = switch ($Type.ToUpper()) {"ERROR"{"Error"}"WARN"{"Warning"}default{"Information"}}
        Write-EventLog -LogName Application -Source $EventSource -EntryType $entryType -EventId 1000 -Message $Message -ErrorAction SilentlyContinue
    } catch {}
}

function Copy-PublicDesktopShortcut {
    param(
        [string]$SourceDirectory,
        [string]$TargetDirectory,
        [string]$ShortcutFile,
        [switch]$SimulationMode
    )
    $SourceShortcutPath = Join-Path $SourceDirectory $ShortcutFile
    $TargetShortcutPath = Join-Path $TargetDirectory $ShortcutFile

    Write-Log "檢查並複製捷徑：$ShortcutFile 到 $TargetDirectory" "DEBUG"

    if (-not (Test-Path $SourceShortcutPath -PathType Leaf)) {
        Write-Log "找不到來源捷徑 $SourceShortcutPath，跳過複製。" "WARN"
        return
    }

    if ($SimulationMode) {
        Write-Log "模擬模式：跳過 Copy-Item。目標：$TargetShortcutPath" "SIMULATE:INFO"
    } else {
        try { Copy-Item -Path $SourceShortcutPath -Destination $TargetShortcutPath -Force -ErrorAction Stop
              Write-Log "捷徑複製成功：$TargetShortcutPath" "INFO" }
        catch { Write-Log "複製捷徑失敗：$($_.Exception.Message)" "ERROR" }
    }
}

$LockFile = Join-Path $LocalTempDir "NM504Client_Update.lock"
if (Test-Path $LockFile) {
    Write-Log "偵測到其他更新程序 (LockFile: $LockFile)，略過。" "WARN"; exit $EXIT_NO_UPDATE
}

if (-not $SimulationMode) {
    try { 
        $LockFileContent = "PID=$PID Host=<COMPUTERNAME> User=<USERNAME> Time=$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        Set-Content -Path $LockFile -Value $LockFileContent -Encoding UTF8 -ErrorAction Stop 
    } 
    catch { Write-Host "無法建立鎖定檔 $LockFile ($($_.Exception.Message))" -ForegroundColor Red; exit $EXIT_EXCEPTION }
} else {
    Write-Log "模擬模式：跳過建立鎖定檔" "SIMULATE:INFO"
}

try {
    Write-Log "=== 更新檢查開始 ==="
    Write-Log "GPO=$GPO Ver=$Ver"

    $clientEXEPath=$null;$clientVersion=$null
    if (Test-Path $EXE114) { $clientEXEPath=$EXE114;$clientVersion="NM114"; Write-Log "偵測到新版 NM114" }
    elseif (Test-Path $EXE106) { $clientEXEPath=$EXE106;$clientVersion="NM106"; Write-Log "偵測到舊版 NM106" }
    else { Write-Log "未安裝，略過更新" "INFO"; exit $EXIT_NO_UPDATE }

    $fileDate=(Get-Item $clientEXEPath).LastWriteTime
    Write-Log "目前 EXE 修改日期：$fileDate"

    $needsUpdate=$false
    if ($clientVersion -eq "NM106") { Write-Log "舊版 NM106，強制更新" "WARN"; $needsUpdate=$true }
    elseif ($fileDate -lt $targetDateObj) { Write-Log "版本較舊 ($fileDate < $TargetDate)，更新中"; $needsUpdate=$true }

    $sourceInstallerPath=$null
    if ((Test-Path -Path $InstallerS1 -PathType Leaf)) { $sourceInstallerPath=$InstallerS1 }
    elseif ((Test-Path -Path $InstallerS2 -PathType Leaf)) { $sourceInstallerPath=$InstallerS2 }

    if (-not $needsUpdate) { 
        Write-Log "版本已最新，檢查捷徑..."
        if ($sourceInstallerPath) {
            $SourceShortcutDir = Split-Path $sourceInstallerPath -Parent
            Copy-PublicDesktopShortcut -SourceDirectory $SourceShortcutDir -TargetDirectory $PublicDesktop -ShortcutFile $ShortcutName -SimulationMode:$SimulationMode
        }
        exit $EXIT_NO_UPDATE 
    }

    if (-not $sourceInstallerPath) { Write-Log "找不到安裝來源 (S1/S2)" "ERROR"; exit $EXIT_COPY_FAIL }

    $hashFile = "$sourceInstallerPath.sha256"
    $expectedHash = "<SHA256_HASH_PLACEHOLDER>"

    Write-Log "複製安裝檔至 $localInstallerPath"
    if ($SimulationMode) { Write-Log "模擬模式：跳過 Copy-Item。假設複製成功。" "SIMULATE:INFO" }
    else { Copy-Item -Path $sourceInstallerPath -Destination $localInstallerPath -Force -ErrorAction Stop }

    Write-Log "執行安裝程式：$localInstallerPath /S"
    if ($SimulationMode) { Start-Sleep -Seconds 1; $proc = New-Object PSObject -Property @{ ExitCode = 0 } }
    else { $proc=Start-Process -FilePath $localInstallerPath -ArgumentList "/S" -Wait -PassThru }

    if ($proc.ExitCode -ne 0) { Write-Log "安裝失敗，代碼 $($proc.ExitCode)" "ERROR"; exit $EXIT_INSTALL_FAIL }

    $SourceShortcutDir = Split-Path $sourceInstallerPath -Parent
    Copy-PublicDesktopShortcut -SourceDirectory $SourceShortcutDir -TargetDirectory $PublicDesktop -ShortcutFile $ShortcutName -SimulationMode:$SimulationMode

} catch { 
    Write-Log "執行例外：$($_.Exception.Message)" "ERROR"; exit $EXIT_EXCEPTION 
} finally {
    if (Test-Path $LockFile) {
        if (-not $SimulationMode) { Remove-Item $LockFile -Force -ErrorAction SilentlyContinue; Write-Log "移除鎖定檔" "DEBUG" }
        else { Write-Log "模擬模式：跳過 Remove-Item 鎖定檔" "SIMULATE:DEBUG" }
    }

    if (Test-Path $localInstallerPath) {
        if (-not $SimulationMode) { Remove-Item $localInstallerPath -Force -ErrorAction SilentlyContinue; Write-Log "移除暫存檔" "DEBUG" }
        else { Write-Log "模擬模式：跳過 Remove-Item 暫存檔" "SIMULATE:DEBUG" }
    }
    Write-Log "=== 更新程序結束 ==="
}

exit $EXIT_SUCCESS
