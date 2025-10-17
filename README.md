# NM504Client 自動更新 Updater (範例公開版)

## 綠起
NM504Client 程式一直沒有自動更新功能，今（114）年公司 NM504Client 大改版（106 -〉 114），一時興起就使用 powershell 寫個更新程式。
程式執行邏輯：
- 檢查電腦是否有安裝版本 NM504Client，未安裝表示無需求，略過。
- 有安裝 114 版本者，再檢查主執行檔（`C:\Program Files\NM114\NM114.exe`）的檔案日期，與最新版本日期（$TargetDate）比較，決定是否更新。
- 未安裝 114 版本但有 106 版本者，直接安裝 114 版本。
- 另外，因為使用 GPO Start Script 派送安裝，沒有在一般使用者桌面上建捷徑，因此複製捷徑到公用桌面。

## 概述
此 PowerShell 可自動判斷「NM504Client」是否需要更新，並執行安裝、捷徑維護、日誌紀錄等操作。設計上支援模擬模式（SimulationMode），可用於測試部署流程而不實際修改系統。

主要功能：
- 檢查已安裝版本 (NM106 / NM114)
- 檢查安裝檔來源（NAS / AD）並自動選擇
- 安裝檔複製與 SHA256 雜湊驗證
- 自動靜默安裝 (`/S`)
- 公用桌面捷徑補救與複製
- 日誌與 EventLog 記錄
- 舊日誌自動清理
- 鎖定檔保護，避免重複執行
- 模擬模式支援 (SimulationMode)

---

## 系統需求
- Windows PowerShell 5.1 或以上
- 需以 **系統管理員權限**執行
- 可透過 GPO 或排程任務部署

---

## 使用方式

### 1. 基本執行
```powershell
.\NM504Client_Update.ps1
```

### 2. 指定模擬模式
模擬模式下不會實際安裝軟體或修改捷徑：
```powershell
.\NM504Client_Update.ps1 -SimulationMode
```

### 3. 自訂參數範例
```powershell
.\NM504Client_Update.ps1 `
    -Ver "2025-01-01.001" `
    -EXE114 "C:\Path\To\NM114.exe" `
    -EXE106 "C:\Path\To\NM106.exe" `
    -DirS1 "\\NAS\Software" `
    -DirS2 "\\AD\SYSVOL\scripts" `
    -TargetDate "2025-01-01" `
    -DirLog "\\LOGSERVER\Logs" `
    -LogLevel "INFO" `
    -LogRetentionDays 30 `
    -Installer "單機版_20250101_x64.exe" `
    -PublicDesktop "C:\Users\Public\Desktop" `
    -ShortcutName "單機版.lnk" `
    -LocalTempDir "$env:SystemDrive\TEMP"
```

---

## 參數說明

| 參數 | 說明 | 預設值 |
|------|------|--------|
| `-Ver` | 版本 | `"2025-01-01.001"` |
| `-GPO` | 用於部署的 GPO 名稱 | `"Powershell"` |
| `-EXE114` | 新版主程式路徑 | `"C:\Path\To\NM114.exe"` |
| `-EXE106` | 舊版主程式路徑 | `"C:\Path\To\NM106.exe"` |
| `-DirS1` | 安裝來源 1（NAS） | `"\\<NAS_SERVER>\Software"` |
| `-DirS2` | 安裝來源 2（AD SYSVOL） | `"\\<AD_SERVER>\SYSVOL\scripts"` |
| `-TargetDate` | 更新判斷基準日期 | `"2025-01-01"` |
| `-DirLog` | 日誌存放路徑 | `"\\<LOG_SERVER>\LogFiles"` |
| `-LogLevel` | 日誌等級（DEBUG / INFO / WARN / ERROR） | `"INFO"` |
| `-LogRetentionDays` | 舊日誌保留天數 | `30` |
| `-Installer` | 安裝檔名稱 | `"單機版_YYYMMDD_x64.exe"` |
| `-PublicDesktop` | 公用桌面路徑 | `"C:\Users\Public\Desktop"` |
| `-ShortcutName` | 捷徑名稱 | `"單機版.lnk"` |
| `-LocalTempDir` | 本機暫存目錄 | `$env:SystemDrive\TEMP` |
| `-SimulationMode` | 模擬模式開關 | 無預設，模擬模式下不會實際複製或安裝，只會顯示日誌。正式執行時移除 `-SimulationMode` |

---

## 鎖定機制
- 執行時會在暫存目錄建立鎖定檔 `NM504Client_Update.lock`，避免同時有多個程序運行。
- 執行完成後會自動刪除鎖定檔。

---

## 日誌管理
- 日誌會存放於指定的 `-DirLog` 目錄，檔名格式為 `<COMPUTERNAME>_YYYYMMDD.log`。
- 可透過 `-LogRetentionDays` 設定自動清理天數。
- 支援寫入 Windows EventLog，EventSource 為 `NM504ClientUpdater`。

---

## 注意事項
- 確保安裝來源可存取，否則會中止。
- 原始檔為 RAR ，先解壓為 EXE 執行檔，並在同一目錄下建立 SHA256 檢查檔。
  `$InstallerPath = ".\1141001_x64.exe"
$Hash = (Get-FileHash $InstallerPath -Algorithm SHA256).Hash
$HashFile = "$InstallerPath.sha256"
Set-Content -Path $HashFile -Value $Hash -Encoding UTF8
Write-Host "SHA256 已生成: $HashFile"
`
- SHA256 需依實際安裝檔生成，或在模擬模式下忽略
- 確認磁碟空間至少 500MB，因為安裝檔有 200MB，原本有寫檢查確認磁碟空間的，應該是特例，所以刪除檢查磁碟空間程式碼（ `if ($drive.Free -lt 500MB) { Write-Log "磁碟空間不足 (剩餘：$($drive.Free/1MB)MB)" "ERROR"; exit $EXIT_EXCEPTION }` ）
- 更新過程中會建立鎖定檔避免重複執行
- 公用桌面捷徑需對所有使用者可寫（這是系統預設）
- 模擬模式下不會進行實際檔案操作或安裝，僅做流程驗證。
- 確保 PowerShell 執行政策允許執行 (`RemoteSigned` 或 `Bypass`)。
- 推薦透過 GPO 或排程任務自動部署。

---

## Exit Code
| Exit Code | 說明 |
|-----------|------|
| 0 | 成功完成 |
| 10 | 無需更新或已有更新程序在執行 |
| 20 | 複製檔案失敗 |
| 30 | 安裝程式執行失敗 |
| 99 | 執行階段例外 |

---

## 更新歷史
- **2025-10-01.001**: 初版發布
- **2025-10-17.011**: 支援自動更新與捷徑維護

