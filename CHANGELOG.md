# Changelog

本文件整理 `FileRecoveryTool` / `Office Recovery Toolkit` 目前可從專案內文件確認的版本變更。

> 目前專案主版本已統一為 `v5.8.5.12`。以下保留舊版本紀錄，作為歷史變更參考。

---

## v5.8.5.12

### Fixed
- 改成預設跳過有加密碼的檔案。 按P鍵時 改輸入密碼是多少 然後批次有密碼的檔案都直接帶入這個密碼嘗試是否可以解開密碼 不行一樣標注 
- 按P鍵時如沒輸入密碼維持跳過加密碼的檔案
- 當有設定密碼時再按一次P鍵可以清除密碼。

---

## v5.8.5.11

### Fixed
- 修正 `P=False` 時，若 probe 發生錯誤仍直接略過加密檔案的邏輯問題
- `ProbeError` 現在會跟著 `已加密檔案跳過 / Skip encrypted files` 開關走，不再無視設定

### Notes
- 這版主要修正加密舊檔在 probe 非成功狀態下的跳過判斷
- 若 `P=False`，系統現在會繼續嘗試後續轉檔 / 互動流程，而不是先被 `ProbeError` 擋掉

---

## v5.8.5.10

### Fixed
- 修正掃描加密的舊版 Word / Office 檔時，互動式密碼提示可能卡住主掃描流程
- 掃描等待 helper 轉檔期間，支援按 `ESC` 取消當前轉檔工作
- `P=否` 時改走 helper 子程序，避免主掃描執行緒被 Office COM 鎖住

### Notes
- 這版重點是降低加密舊檔互動解密時的卡死風險
- 是否能成功跳出密碼框與轉檔，仍取決於 Windows / Office 實際行為

---

## v5.8.5.9

### Added
- 新增 `P` 選單：`已加密檔案跳過 / Skip encrypted files`
- 預設為開啟，維持既有安全行為
- 關閉後，Legacy Office 轉檔流程可改為嘗試互動式密碼提示
- 右側設定面板新增目前開關狀態顯示

### Notes
- 這版主要讓加密檔案的處理更可控，方便需要人工輸入密碼再轉檔的情境
- 是否真的跳出密碼視窗仍取決於 Windows / Office 實際執行環境

---

## v5.8.5.8

來源：`README.md`（舊內容）

### Added
- 支援判斷加密的 `.pdf` 與 `.rtf` 格式
- 掃描進度列升級為 PRO 顯示
- 實際改名進度列套用同一套 PRO 顯示

### Highlights
- 掃描格式擴充為：
  - `docx`
  - `xlsx`
  - `pptx`
  - `doc`
  - `xls`
  - `ppt`
  - `rtf`
  - `pdf`
  - `txt`
- LightBar 互動式選單
- HTML 報表與主檔 ⭐ 顯示
- SHA256 去重與內容指紋比對
- Excel 智慧命名與全工作表解析
- 主檔 / 重複檔 / 唯一檔 / 損毀檔分類
- 模擬改名 / 實際改名
- 安全整理模式：Copy / Move
- Primary files only 模式
- 多語系介面（zh-TW / en-US）

---

## v5.8.5.6

來源：`WHAT'S NEW.txt`

### Added
- 實際改名前先顯示「檔案計算中，請稍後...」
- 實際改名加入完整進度列
- 支援按 `ESC` 中止流程並彈出 UI 確認
- 中止後可輸出中止任務報表
- 產生彩色的實際改名 HTML 報表
- `Dashboard PRO` 整合進主選單

### Notes
- 這版重點在於實際改名流程、可中止控制與報表視覺化強化

---

## v5.8.5.3

來源：`WHAT'S NEW.txt`

### Added
- 新增 `.xls` / `.xlsx` 的「名冊」命名規則

### Rule Details
若內容偵測到至少兩類名冊特徵欄位，例如：
- 姓名 / Name
- 日期 / Date
- 身分證號 / 身份證號 / 身分證字號
- ID No / ID Number / ID Card Number

則建議檔名優先判定為：
- `名冊`
- 若可辨識日期，則命名為：`名冊_YYYY-MM-DD`

### Improvements
- `.xlsx`：同時分析工作表名稱與工作表內容
- `.xls`：分析掃描階段抽出的 `PreviewText`
- 提高此命名規則的信心分數，讓報表更明確標示為高可信度命名

---

## v5.8.5.2

來源：`README.txt`

### Highlights
- LightBar 選單版本
- 支援掃描：
  - `docx`
  - `xlsx`
  - `pptx`
  - `doc`
  - `xls`
- 強化報表功能
- 美化報表樣式
- SHA256 雜湊去重
- 檔案內容指紋比對
- Excel 全工作表解析
- Excel 智慧命名
- 主檔 / 重複檔 / 唯一檔 / 損毀檔分類
- 模擬改名 / 實際改名
- HTML 報表
- 開啟最新報表
- 安全整理模式：Copy / Move
- Primary files only 模式
- 依副檔名分類整理
- Quality Score 品質分數
- HTML 顯示主檔 ⭐ Primary

---

## 後續建議

建議之後固定維護一個單一版本來源，例如：

1. 每次 release 同步更新：
   - `CHANGELOG.md`
   - `README.md`
   - 介面標題版本字串
2. 採用一致格式，例如：
   - `v5.8.5.9`
   - 日期
   - Added / Changed / Fixed
3. 若之後有 Git，可改用 tag 管理版本
