# NM504 Client Updater (範例公開版)

## 綠起
NM504Client 程式一直沒有自動更新功能，今（114）年公司 NM504Client 大改版（106 -〉 114），一時興起就使用 powershell 寫個更新程式。
程式執行邏輯：
- 檢查電腦是否有安裝版本 NM504Client，未安裝表示無需求，略過。
- 有安裝 114 版本者，再檢查主執行檔（`C:\Program Files\NM114\NM114.exe`）的檔案日期，與最新版本日期（$TargetDate）比較，決定是否更新。
- 未安裝 114 版本但有 106 版本者，直接安裝 114 版本。
- 另外，因為使用 GPO Start Script 派送安裝，沒有在一般使用者桌面上建捷徑，因此複製捷徑到公用桌面。

## 說明
這是 NM504 Client 更新安裝腳本的公開範例版本，功能包括：

- 檢查已安裝版本 (NM106 / NM114)
- 安裝檔複製與 SHA256 驗證
- 模擬模式 (SimulationMode)
- 鎖定檔保護，避免多重執行
- 公用桌面捷徑複製
- 日誌與 EventLog 記錄

## 使用方法

### 1. 下載腳本
將 `NM504ClientUpdater_public.ps1` 放置於可執行目錄。

### 2. 修改參數
編輯腳本開頭的 `param()` 區塊，根據環境修改：

| 參數 | 說明 | 範例 |
|------|------|------|
| `EXE114` | 新版 NM114 檔案路徑 | `C:\Program Files\NM114\NM114.exe` |
| `EXE106` | 舊版 NM106 檔案路徑 | `C:\Program Files\NM106\NM106.exe` |
| `DirS1` | 安裝檔來源 NAS | `\\NAS_SERVER\Software` |
| `DirS2` | 安裝檔來源 AD | `\\AD_SERVER\SYSVOL\scripts` |
| `DirLog` | 日誌目錄 | `\\LOG_SERVER\LogFiles\NM114-504Client` |
| `PublicDesktop` | 公用桌面路徑 | `C:\Users\Public\Desktop` |
| `ShortcutName` | 捷徑檔名 | `單機版.lnk` |
| `SimulationMode` | 模擬模式 | `-SimulationMode` |

### 3. 執行腳本
以系統管理員權限開啟 PowerShell：
```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass 
-File "\\AD_SERVER\SYSVOL\script\NM504ClientUpdater_public.ps1" 
-EXE114 "C:\Program Files\NM114\NM114.exe" 
-EXE106 "C:\Program Files\NM106\NM106.exe" 
-DirS1 "\\NAS_SERVER\Software" 
-DirS2 "\\AD_SERVER\SYSVOL\scripts" 
-TargetDate "2025-09-29" 
-DirLog "\\LOG_SERVER\LogFiles" 
-Installer "1141001_x64.exe" 
-SimulationMode
```
模擬模式下不會實際複製或安裝，只會顯示日誌。正式執行時移除 `-SimulationMode`。

### 4. 日誌與檢查
日誌會寫入到 `DirLog` 目錄下，檔名格式為 `<COMPUTERNAME>_YYYYMMDD.log`。

EventLog 也會產生資訊，EventSource 為 `NM504ClientUpdater`。

### 5. 注意事項

- 原始檔為 RAR ，先解壓為 EXE 執行檔，並在同一目錄下建立 SHA256 檢查檔。
- SHA256 需依實際安裝檔生成，或在模擬模式下忽略
- 確認磁碟空間至少 500MB，因為安裝檔有 200MB，原本有寫檢查確認磁碟空間的，應該是特例，所以刪除檢查磁碟空間程式碼（ `if ($drive.Free -lt 500MB) { Write-Log "磁碟空間不足 (剩餘：$($drive.Free/1MB)MB)" "ERROR"; exit $EXIT_EXCEPTION }` ）
- 更新過程中會建立鎖定檔避免重複執行
- 公用桌面捷徑需對所有使用者可寫（這是系統預設）

## 範例
模擬檢查更新：
```powershell
.\NM504ClientUpdater_public.ps1 -SimulationMode
```

正式安裝更新：
```powershell
.\NM504ClientUpdater_public.ps1
```

## 授權
此範例腳本為公開示範，請依公司政策使用。
