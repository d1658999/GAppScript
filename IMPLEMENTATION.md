# MEGA 資料分析器 - 完整實作說明

## 📋 專案概述

此專案完全按照 `analyze_mega_data.prompt.md` 指示建立，實現了以下功能：

### ✅ 已實現的需求

1. **自動取得 Google Sheets 圖表**
   - 連結：https://docs.google.com/spreadsheets/d/1nnzLzALFvtwW0Lz8aVGCw7O0iaPXwN3_wRUF3BC2iak/edit
   - 目標分頁：`[NR_TX_LMH]Summary&NR_Test_1`
   - 智慧識別三個特定圖表

2. **自動建立 Google Slides**
   - 連結：https://docs.google.com/presentation/d/1_zV6grTwFms2a0jE2i9Tz0X6mjbnCpZk8VDKIyBLxi8/edit
   - 自動建立新投影片
   - 專業格式化標題

3. **保持原始比例插入圖片**
   - 圖表比例：5:1（符合指示要求）
   - 避免任何變形或拉伸
   - 智慧居中對齊

4. **優化排版設計**
   - 垂直排列
   - 每張投影片最多 3 張圖片
   - 圖表固定高度 1.4 英寸，提供緊湊佈局
   - 不顯示圖表標題，保持乾淨視覺效果

5. **自動分頁處理**
   - 超過 3 個圖表時自動建立新投影片
   - 投影片標題包含頁碼

## 📁 檔案結構

```
mega_data/
├── analyze_mega_data.gs     # 主要功能實作
├── demo.gs                  # 示範和測試函式
├── test_runner.gs          # 測試執行器
├── appsscript.json         # 專案設定檔
└── README.md               # 詳細技術文件

根目錄/
├── QUICK_START.md          # 快速執行指南
└── IMPLEMENTATION.md       # 本檔案
```

## 🚀 立即執行

### 方法 1: 一鍵執行（推薦）

```javascript
// 在 demo.gs 中執行
runMegaDataAnalysis();
```

### 方法 2: 先測試再執行

```javascript
// 1. 執行完整測試
runCompleteTestSuite();

// 2. 乾運行測試（不建立投影片）
dryRunAnalysis();

// 3. 確認無誤後執行主功能
runMegaDataAnalysis();
```

## 🎯 核心功能說明

### 智慧圖表識別

系統會自動識別以下三個圖表：
- `MaxPwr - TX Power`
- `MaxPwr - ACLR - Max(E_-1,E_+1)`
- `MaxPwr - EVM`

識別方法：
1. **標題匹配** - 直接比對圖表標題
2. **內容分析** - 分析圖表資料範圍
3. **位置推測** - 根據圖表位置推測

### 5:1 比例保持

- 自動計算最適尺寸
- 保持原始比例，避免變形
- 智慧居中對齊
- 最大化利用投影片空間

### 多投影片支援

- 每張投影片最多 3 個圖表
- 自動計算所需投影片數量
- 投影片標題包含頁碼 (例: "標題 (1/2)")

### 垂直排列設計

根據圖表數量自動選擇最佳排版：
- **1 張圖表**: 居中顯示
- **2 張圖表**: 垂直平分
- **3 張圖表**: 緊湊垂直排列

## 🔧 技術特色

### 符合 Google Apps Script 規範

- ✅ 使用 V8 Runtime
- ✅ 嚴禁使用 `var`，全面使用 `const/let`
- ✅ camelCase 命名規則
- ✅ 完整 JSDoc 註解，包含權限標示
- ✅ 2 空格縮排，80 字元行長度限制
- ✅ 適當錯誤處理和日誌記錄

### 權限管理

程式碼中清楚標示所需權限：
```javascript
/**
 * @permission SpreadsheetApp, SlidesApp, DriveApp
 */
```

設定檔案 `appsscript.json` 包含所有必要權限：
- Google Sheets API
- Google Slides API
- Google Drive API

### 錯誤處理

- 完整的 try/catch 結構
- 詳細的錯誤日誌
- 友好的錯誤訊息
- 自動恢復機制

## 📊 執行結果

成功執行後會：

1. **在 Google Slides 中建立新投影片**
   - 包含格式化標題
   - 保持 5:1 比例的圖表
   - 專業的垂直排列佈局

2. **回傳詳細結果資訊**
   ```javascript
   {
     success: true,
     slidesCreated: 1,        // 建立的投影片數量
     slideIds: ["slide_id"],  // 投影片 ID 陣列
     chartsProcessed: 3       // 處理的圖表數量
   }
   ```

3. **輸出執行日誌**
   - 詳細的處理步驟
   - 圖表識別結果
   - 排版計算過程
   - 成功/錯誤訊息

## 🧪 測試驗證

### 基本測試

```javascript
// 連線測試
testConnections();

// 圖表篩選測試
testChartFiltering();
```

### 進階測試

```javascript
// 完整測試套件
runCompleteTestSuite();

// 乾運行測試（不建立投影片）
dryRunAnalysis();
```

### 功能驗證

```javascript
// 排版計算測試
demoProportionalLayout();

// 截圖樣式測試
testScreenshotStyleLayout();
```

## 📈 效能特色

- **智慧圖表識別**: 多層次識別演算法，確保找到正確圖表
- **原始比例保持**: 精確的寬高比計算，避免圖表變形
- **自動排版**: 根據圖表數量智慧選擇最佳佈局
- **錯誤恢復**: 單一圖表處理失敗不影響其他圖表
- **日誌追蹤**: 完整的執行過程記錄，便於偵錯

## 🛠 客製化選項

### 修改目標圖表

```javascript
const TARGET_CHART_TITLES = [
    'MaxPwr - TX Power',
    'MaxPwr - ACLR - Max(E_-1,E_+1)',
    'MaxPwr - EVM',
    // 可新增更多圖表
];
```

### 調整投影片配置

```javascript
const MAX_CHARTS_PER_SLIDE = 3; // 每張投影片最多圖表數
```

### 更改比例設定

```javascript
const chartAspectRatio = 5 / 1; // 5:1 寬高比
```

## 🎯 品質保證

- **無語法錯誤**: 所有檔案通過語法檢查
- **符合規範**: 遵循 Google Apps Script 編碼規範
- **完整測試**: 包含單元測試和整合測試
- **詳細文件**: 完整的使用說明和技術文件
- **錯誤處理**: 全面的例外處理機制

## 📞 支援資源

- 📖 **詳細文件**: `mega_data/README.md`
- 🚀 **快速指南**: `QUICK_START.md`
- 🧪 **測試工具**: `test_runner.gs`
- 🎯 **示範功能**: `demo.gs`

---

**專案狀態**: ✅ 完全實作完成  
**版本**: v3.1.1  
**最後更新**: 2025-08-12  
**作者**: GitHub Copilot  
**相容性**: Google Apps Script V8 Runtime
