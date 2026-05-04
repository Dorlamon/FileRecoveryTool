# Build and Release Guide

這份文件整理 `FileRecoveryTool` 的建置、測試與發佈方式。

---

## 1. 專案組成

本專案主要分成兩部分：

### A. PowerShell 主工具
- `OfficeRecoveryToolkit.ps1`
- `OfficeRecoveryToolkit.cmd`

這部分負責：
- 主選單 UI
- 掃描流程
- 分析結果
- 報表輸出
- 模擬改名 / 實際改名
- 檔案整理

### B. C# 偵測器
- `OfficeEncryptionProbe.Program.cs`
- `OfficeEncryptionProbe.csproj`

這部分負責：
- 判斷 Office / PDF / RTF 是否加密或受保護
- 回傳結構化結果與 exit code

---

## 2. 建置環境

### 執行環境
- Windows 10 / 11
- PowerShell 5.1

### C# 建置環境
- .NET 10 SDK

### 建議附加環境
- Microsoft Office
- 或 Office 相容轉換元件

---

## 3. 日常開發方式

如果只是修改主工具邏輯：
- 直接編輯 `OfficeRecoveryToolkit.ps1`
- 用 Windows PowerShell 執行 `OfficeRecoveryToolkit.cmd` 測試

如果修改了加密/保護判斷邏輯：
- 編輯 `OfficeEncryptionProbe.Program.cs`
- 重新 build `OfficeEncryptionProbe.exe`
- 放回工具目錄供 PowerShell 主程式呼叫

---

## 4. 建置 OfficeEncryptionProbe

### 方法 A：使用批次檔

直接執行：

```bat
OfficeEncryptionProbe_BUILD.bat
```

其內容等同於：

```bat
dotnet publish OfficeEncryptionProbe.csproj -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true /p:IncludeNativeLibrariesForSelfExtract=true
```

### 方法 B：手動執行

在專案目錄中執行：

```bat
dotnet publish .\OfficeEncryptionProbe.csproj -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
```

### 輸出位置

```text
.\bin\Release\net10.0\win-x64\publish\OfficeEncryptionProbe.exe
```

---

## 5. OfficeEncryptionProbe 測試方式

可用命令列測試：

```bat
OfficeEncryptionProbe.exe --json --path "C:\Test\sample.xls"
```

### Exit Codes

- `0` = NotProtected
- `1` = Encrypted
- `2` = WriteProtectedOnly
- `3` = PossiblyProtected
- `4` = Corrupt
- `5` = Unsupported
- `6` = Error

---

## 6. 主工具執行方式

在 Windows 直接執行：

```bat
OfficeRecoveryToolkit.cmd
```

它會呼叫：

```bat
powershell -ExecutionPolicy Bypass -File .\OfficeRecoveryToolkit.ps1
```

---

## 7. 建議測試清單

每次 release 前，建議至少做以下測試：

### 基本流程
- [ ] 可正常開啟主選單
- [ ] 可切換語言
- [ ] 可設定掃描資料夾
- [ ] 可完成掃描
- [ ] 可匯出 HTML 報表
- [ ] 可開啟最新報表

### 改名流程
- [ ] 模擬改名正常
- [ ] 實際改名正常
- [ ] ESC 中止流程正常
- [ ] 中止報表可輸出

### 整理流程
- [ ] Copy 模式正常
- [ ] Move 模式正常
- [ ] Primary only 模式正常
- [ ] 依副檔名分類整理正常

### 檔案判定
- [ ] 可辨識正常 Office 檔
- [ ] 可辨識重複檔
- [ ] 可標示損毀檔
- [ ] 可偵測加密或受保護檔案
- [ ] `.xls` / `.xlsx` 名冊規則正常

---

## 8. 建議測試資料集

建議準備一組固定測試資料夾，至少包含：

- 正常 `docx / xlsx / pptx`
- 舊格式 `doc / xls / ppt`
- `pdf / rtf / txt`
- 重複檔案
- 損毀檔案
- 加密檔案
- 寫入保護檔案
- 名冊類型 Excel 檔案
- 檔名遺失但內容可辨識的樣本

這樣每次改版都能快速回歸測試。

---

## 9. 發佈建議

### 建議發佈包內容

```text
FileRecoveryTool/
├── OfficeRecoveryToolkit.cmd
├── OfficeRecoveryToolkit.ps1
├── OfficeEncryptionProbe.exe
├── README.md
├── GUIDE.txt
├── WHAT'S NEW.txt
└── (必要的設定或樣本說明)
```

### 發佈前檢查
- [ ] `OfficeEncryptionProbe.exe` 已更新
- [ ] 版本號一致
- [ ] `README.md` 已更新
- [ ] `CHANGELOG.md` 已更新
- [ ] `WHAT'S NEW.txt` 已更新
- [ ] 用乾淨機器或乾淨資料夾實測一次

### 建議壓縮格式
- ZIP

### 建議命名格式

目前版本建議：

```text
FileRecoveryTool_v5.8.5.8_win-x64.zip
```

---

## 10. 版本管理建議

如果後面要持續維護，建議固定一套 release 流程：

1. 修改程式
2. 本機測試
3. Build `OfficeEncryptionProbe.exe`
4. 跑回歸測試
5. 更新：
   - `README.md`
   - `CHANGELOG.md`
   - `WHAT'S NEW.txt`
6. 打包 release ZIP
7. 保存對應版本原始碼與輸出檔

---

## 11. 未來可再補強的方向

- 新增一鍵 release script
- 自動驗證 `OfficeEncryptionProbe.exe` 是否存在
- 將版本號集中到單一設定來源
- 補測試樣本與測試記錄模板
- 規劃 Git tag / release notes 流程

---

## 12. 備註

本專案目前屬於實用型 Windows 工具。若未來想提升可維護性，最值得優先補強的是：

- 版本管理一致性
- build / release 流程固定化
- 測試資料集標準化
