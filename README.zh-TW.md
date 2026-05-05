# FileRecoveryTool

Office 文件救援後智慧整理工具。

這個工具主要用在檔案救援後，原始檔名遺失、內容混雜、難以人工辨識的情境。它會透過檔案內容分析、雜湊比對、格式解析與命名規則，協助使用者快速辨識主檔、重複檔、損毀檔，並產生建議檔名與分析報表。

![screenshot](https://github.com/Dorlamon/FileRecoveryTool/blob/main/screenshot.zh-tw.png)

---

## 核心用途

當磁碟救援後失去原始檔名資訊（例如 MFT 遺失）時，本工具可作為輔助整理工具，協助：

- 從文件內容推測合理檔名
- 找出重複檔案
- 區分主檔、重複檔、唯一檔、損毀檔
- 預覽改名結果
- 執行實際改名
- 匯出 HTML 分析報表
- 依規則整理輸出檔案

---

## 主要功能

- 支援掃描格式：
  - `docx` `xlsx` `pptx`
  - `doc` `xls` `ppt`
  - `rtf` `pdf` `txt`
- SHA256 雜湊去重複
- 檔案內容指紋比對
- Excel 全工作表解析
- Excel 智慧命名
- 主檔 / 重複檔 / 唯一檔 / 損毀檔分類
- 模擬改名
- 實際改名
- HTML 報表輸出
- 開啟最新分析報表
- 安全整理模式：Copy / Move
- Primary files only 模式
- 依副檔名分類整理
- Quality Score 評分
- 中英文介面切換
- 加密 / 保護檔案偵測（Office / PDF / RTF）

---

## 專案結構

```text
FileRecoveryTool/
├── OfficeRecoveryToolkit.cmd
├── OfficeRecoveryToolkit.ps1
├── OfficeEncryptionProbe.Program.cs
├── OfficeEncryptionProbe.csproj
├── OfficeEncryptionProbe_BUILD.bat
├── OfficeEncryptionProbe_BUILD.txt
├── GUIDE.txt
├── WHAT'S NEW.txt
├── README.txt
└── README.md
```

### 各檔案用途

- `OfficeRecoveryToolkit.cmd`
  - Windows 啟動入口
- `OfficeRecoveryToolkit.ps1`
  - 主工具程式，包含掃描、分析、報表、改名、整理等主要功能
- `OfficeEncryptionProbe.Program.cs`
  - C# 偵測器，用來判斷 Office / PDF / RTF 是否加密、保護或損毀
- `OfficeEncryptionProbe.csproj`
  - C# 專案設定
- `GUIDE.txt`
  - 快速操作說明
- `WHAT'S NEW.txt`
  - 版本更新摘要

---

## 系統需求

- Windows 10 / 11
- PowerShell 5.1
- 建議具備 Microsoft Office / 相容轉換元件環境

### 必要安裝

Microsoft Office Word、Excel、PowerPoint 2007 File Format Compatibility Pack：

<(https://driver.uch.edu.tw/?path=old_ftp%2F01_Microsoft_%E5%BE%AE%E8%BB%9F%E6%A0%A1%E5%9C%92%E8%BB%9F%E9%AB%94%2FOffice%2FMS_Office_2007_Enterprise%2F%E7%9B%B8%E5%AE%B9%E6%80%A7%E5%A5%97%E4%BB%B6)>

> 備註：本專案主要設計給 Windows 環境使用。在 macOS 上可閱讀、編輯、整理原始碼，但完整執行流程仍建議在 Windows 上進行。

---

## 執行方式

### 一般使用

解壓縮後直接執行：

```bat
OfficeRecoveryToolkit.cmd
```

該指令會呼叫：

```bat
powershell -ExecutionPolicy Bypass -File .\OfficeRecoveryToolkit.ps1
```

---

## 基本操作流程

建議操作順序：

1. 執行 `OfficeRecoveryToolkit.cmd`
2. 設定掃描資料夾
3. 開始掃描
4. 檢視分析結果
5. 匯出 HTML 報表
6. 先做模擬改名
7. 確認結果後再執行實際改名
8. 視需要整理主檔 / 重複檔

---

## 快速鍵 / 主選單操作

根據目前說明文件：

- `L`：切換語言
- `0`：切換 Primary files only 模式
- `4`：設定掃描資料夾
- `1`：開始掃描
- `2`：匯出分析報表
- `7`：開啟最新分析報表
- `5`：模擬自動改名
- `6`：執行實際自動改名
- `8`：整理主檔 / 重複檔到資料夾
- `3`：開啟輸出資料夾

---

## 掃描與判斷邏輯概要

本工具會綜合使用以下資訊進行分析：

- 檔案副檔名
- 檔案內容文字
- SHA256 雜湊值
- 內容指紋
- Excel 工作表名稱與內容
- 檔案保護 / 加密狀態
- 可疑損毀特徵

### 檔案角色分類

- **Primary**：推定為主要保留檔案
- **Duplicate**：與其他檔案重複
- **Unique**：可辨識但無重複對象
- **Corrupted**：檔案損毀或無法正確解析

---

## OfficeEncryptionProbe 說明

專案內含一個 C# 偵測器元件，用來判斷下列格式是否受保護：

- `docx`
- `xlsx`
- `pptx`
- `doc`
- `xls`
- `ppt`
- `pdf`
- `rtf`

可判斷狀態包括：

- `NotProtected`
- `Encrypted`
- `WriteProtectedOnly`
- `PossiblyProtected`
- `Corrupt`
- `Unsupported`
- `Error`

這個元件可降低工具在面對加密或受保護檔案時的誤判與例外中斷風險。

---

## 輸出內容

工具可產生以下類型輸出：

- HTML 分析報表
- CSV / 改名預覽資料
- 實際改名結果報表
- 任務中止報表
- 檔案整理結果

輸出路徑預設由工具內設定控制，通常會建立在專案目錄下的 `Output` 資料夾。

---

## 安全使用建議

為避免誤改重要檔案，建議遵循以下流程：

1. 永遠先執行掃描
2. 先看 HTML 報表
3. 先做模擬改名
4. 確認命名邏輯正確後，再做實際改名
5. 若資料非常重要，先建立完整備份
6. 優先使用 `Copy` 模式，再考慮 `Move` 模式

---

## 已知限制

- 主要執行環境為 Windows，不是跨平台 GUI 工具
- 部分舊 Office 格式可能依賴 Office 相容元件或轉換流程
- 真正救援品質仍取決於原始檔案損毀程度
- 若檔案內容過少、過度破損或為加密檔，智慧命名準確率會下降

---

## 適用情境

- 磁碟救援後檔名遺失
- 大量 Office 文件混雜難以人工整理
- 想快速辨識重複檔與主檔
- 想先模擬整理結果再決定是否批次改名
- 需要可視化分析報表協助檢查結果

---

## 版本說明

目前統一版本為：`v5.8.5.12`

歷史版本演進可見於：

- `v5.8.5.3`
- `v5.8.5.6`
- `v5.8.5.10`
- `v5.8.5.9`
- `v5.8.5.8`

目前版本重點包含：

- 修正掃描遇到加密 Word 2003 等舊版 Office 檔時，互動式密碼提示可能卡住整個掃描流程的問題
- 掃描等待 helper 轉檔時支援 ESC 取消
- 新增 `P` 選單：已加密檔案跳過（預設開啟）
- 關閉跳過時，Legacy 轉檔流程可改為嘗試互動式密碼提示
- `.pdf` / `.rtf` 加密判斷支援
- 改名前檔案計算提示
- 實際改名完整進度列
- ESC 中止與 UI 確認
- 中止任務報表
- 彩色 HTML 改名報表
- Dashboard PRO 整合
- 名冊類型 Excel 命名規則強化

---

## 作者備註

這是一套面向實務檔案救援整理場景的工具，不只是單純改名腳本，而是結合內容解析、規則判斷、報表輸出與安全整理流程的文件分析工具。
