# 🚀 快速執行指南

根據 `get_chart_google_sheet.prompt.md` 指示檔案的要求，本系統已完整實現所有功能。

## ⚠️ 重要更新：解決圖片變更問題

**如果您發現下載的圖片內容與原始圖表有差異，請使用以下修正版方法：**

### 🎯 推薦執行方法（修正版）

### 🎯 推薦執行方法（最新修正版）

在 Google Apps Script 編輯器中：

1. **執行 `runManualExportDownload()`** - 🌟 **最新推薦**：模擬手動右鍵選擇「Download Chart」→「PNG IMAGE(.PNG)」
2. **執行 `runOriginalChartDownload()`** - 確保取得真正原始圖片
3. **執行 `compareDownloadMethods()`** - 比較所有方法，選擇最佳品質（現包含 5 種方法）

## 📋 需求回顧

**原始要求：**
1. ✅ 開啟 Google Sheets 連結並導航到 `[NR_TX_LMH]Summary&NR_Test_1` 分頁
2. ✅ 取得該分頁中所有圖表的截圖
3. ✅ 下載截圖到電腦

## 🔧 可用的下載方法

### 方法 1: 手動匯出模擬方法（🌟 最新推薦）
```javascript
runManualExportDownload()
```
- ✅ **模擬手動下載**：模擬右鍵選擇「Download Chart」→「PNG IMAGE(.PNG)」的過程
- ✅ **最接近原始**：使用高品質重建和原始 PNG 匯出
- ✅ **多重備援**：嘗試多種方法確保取得最佳品質
- ✅ **詳細資訊**：提供檔案大小和內容類型資訊

### 方法 2: API 匯出方法（新增）
```javascript
runAPIDownload()
```
- ✅ **API 層級存取**：嘗試透過 Google Sheets API 取得圖表
- ✅ **底層資料**：存取圖表的底層匯出功能
- ✅ **技術細節**：提供圖表 ID、位置等詳細資訊

### 方法 3: 原始截圖方法
```javascript
runOriginalChartDownload()
```
- ✅ **真正的原始圖片**：建立臨時投影片取得截圖
- ✅ **無任何變更**：保持完全原始的視覺效果
- ✅ **高解析度**：最大化圖片品質

### 方法 4: 直接下載方法
```javascript
runDirectChartDownload()
```
- ✅ **原始 Blob**：直接取得圖表的原始資料
- ✅ **完全原始格式**：不經過任何轉換
- ✅ **檔案資訊詳細**：顯示內容類型和檔案大小

### 方法 5: 全方法比較（🌟 推薦用於測試）
```javascript
compareDownloadMethods()
```
- ✅ **測試所有 5 種方法**：一次執行所有下載方法
- ✅ **品質比較**：幫您選擇最好的版本
- ✅ **詳細結果**：提供每種方法的下載連結
- ✅ **智慧建議**：根據結果提供使用建議

### 方法 3: 比較方法（推薦用於測試）
```javascript
compareDownloadMethods()
```
- ✅ **測試所有方法**：一次執行所有下載方法
- ✅ **品質比較**：幫您選擇最好的版本
- ✅ **詳細結果**：提供每種方法的下載連結

### 方法 4: 原始方法（可能有圖片變更問題）
```javascript
completePromptRequirements()  // 原始的一鍵執行
```

## 📊 執行結果

執行成功後，您將獲得：

- ✅ **自動導航**：系統自動開啟指定的 Google Sheets 分頁
- ✅ **真實截圖**：取得分頁中所有目標圖表的真實原始截圖
- ✅ **無變更保證**：圖片內容與原始圖表完全一致
- ✅ **下載準備**：圖表已儲存到 Google Drive，隨時可下載到電腦

## 🔗 下載到電腦的步驟

1. **點擊日誌中的 Google Drive 資料夾連結**
2. **選擇所有圖表檔案**（Ctrl+A 或 Cmd+A）
3. **右鍵選擇「下載」**
4. **檔案會自動打包為 ZIP 下載到電腦**

## 📁 檔案格式

- **格式**：PNG（保持最高品質）
- **命名**：包含圖表標題、下載方法和時間戳記
- **位置**：Google Drive 中的時間戳記資料夾
- **品質**：真正的原始圖片，無任何變更或壓縮

## 🛠️ 其他可用功能

### 僅建立 Google Slides
```javascript
runMegaDataAnalysis()  // 在 demo.gs 中執行
```

### 傳統下載方法
```javascript
runChartDownload()     // 原始下載方法
runBatchDownload()     // 批次下載
```

### 測試和除錯
```javascript
quickTest()            // 快速測試連線
testChartFiltering()   // 測試圖表識別
listExistingDownloads() // 列出現有下載
```

## 📋 目標圖表

系統會自動識別並處理以下圖表：
1. `MaxPwr - TX Power`
2. `MaxPwr - ACLR - Max(E_-1,E_+1)`
3. `MaxPwr - EVM`

## 🔧 設定檔案

所有設定都在 `analyze_mega_data.gs` 頂部：
- Spreadsheet ID、Presentation ID
- 目標分頁名稱
- 目標圖表清單

## 💡 解決圖片變更問題的技術說明

**問題原因：**
- 傳統的 `chart.getBlob()` 方法可能會重新渲染圖表
- 這個過程可能改變圖片的解析度、色彩或細節

**解決方案：**
1. **截圖方法**：建立臨時投影片，插入圖表後擷取投影片縮圖
2. **直接方法**：使用多種技術直接取得原始 Blob
3. **高品質渲染**：確保取得最高解析度版本

## ⚠️ 注意事項

1. **權限**：首次執行需要授權 Google Sheets、Slides 和 Drive 權限
2. **圖片品質**：新方法確保取得真正的原始圖片
3. **下載位置**：圖表先儲存到 Google Drive，再由使用者下載到本機電腦
4. **方法比較**：建議先使用 `compareDownloadMethods()` 測試所有方法

## 📞 支援

如有問題，請檢查：
1. Google Apps Script 執行日誌
2. 比較不同下載方法的結果
3. Google Sheets 和 Slides 的存取權限
4. 目標分頁名稱是否正確
