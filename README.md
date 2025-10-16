# NM504 Client Updater (範例公開版)

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
.\NM504ClientUpdater_public.ps1 -SimulationMode
```
模擬模式下不會實際複製或安裝，只會顯示日誌。正式執行時移除 `-SimulationMode`。

### 4. 日誌與檢查
日誌會寫入到 `DirLog` 目錄下，檔名格式為 `<COMPUTERNAME>_YYYYMMDD.log`。

EventLog 也會產生資訊，EventSource 為 `NM504ClientUpdater`。

### 5. 注意事項

- 確認磁碟空間至少 500MB
- 更新過程中會建立鎖定檔避免重複執行
- SHA256 需依實際安裝檔生成，或在模擬模式下忽略
- 公用桌面捷徑需對所有使用者可寫

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
