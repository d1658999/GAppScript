# MEGA 資料分析器 - 快速執行指南

## 🚀 立即開始

### 步驟 1: 準備工作

1. 確認您有以下 Google 文件的存取權限：
   - **Google Sheets**: https://docs.google.com/spreadsheets/d/1nnzLzALFvtwW0Lz8aVGCw7O0iaPXwN3_wRUF3BC2iak/edit
   - **Google Slides**: https://docs.google.com/presentation/d/1_zV6grTwFms2a0jE2i9Tz0X6mjbnCpZk8VDKIyBLxi8/edit

2. 開啟 Google Apps Script: https://script.google.com/

### 步驟 2: 設定專案

1. 建立新的 Google Apps Script 專案
2. 將以下檔案內容複製到對應的檔案中：
   - `appsscript.json` - 專案設定檔
   - `analyze_mega_data.gs` - 主要功能實作
   - `demo.gs` - 示範和測試函式
   - `test_runner.gs` - 測試執行器

### 步驟 3: 執行測試

在 Google Apps Script 編輯器中：

1. **快速測試** - 執行 `runCompleteTestSuite()` 函式
   ```javascript
   // 在 test_runner.gs 中
   runCompleteTestSuite();
   ```

2. **乾運行測試** - 執行 `dryRunAnalysis()` 函式（不會建立實際投影片）
   ```javascript
   // 在 test_runner.gs 中  
   dryRunAnalysis();
   ```

### 步驟 4: 正式執行

確認測試通過後，執行主要功能：

```javascript
// 在 demo.gs 中
runMegaDataAnalysis();
```

## 📋 執行檢查清單

### 執行前檢查

- [ ] 確認有 Google Sheets 和 Slides 的編輯權限
- [ ] 確認目標分頁 `[NR_TX_LMH]Summary&NR_Test_1` 存在
- [ ] 確認分頁中有圖表資料
- [ ] 執行 `runCompleteTestSuite()` 測試

### 執行後驗證

- [ ] 檢查執行日誌無錯誤訊息
- [ ] 確認 Google Slides 中新增了投影片
- [ ] 驗證圖表已正確插入且保持 5:1 比例
- [ ] 確認每張投影片最多有 3 個圖表
- [ ] 檢查圖表標題格式正確

## 🎯 功能說明

### 自動處理的目標圖表

系統會自動識別並處理以下圖表：

1. **MaxPwr - TX Power**
2. **MaxPwr - ACLR - Max(E_-1,E_+1)**  
3. **MaxPwr - EVM**

### 投影片設計特色

- **比例保持**: 圖表維持 5:1 寬高比，避免變形
- **固定高度**: 所有圖表高度固定為 1.4 英寸，提供緊湊佈局
- **垂直排列**: 圖表以垂直方式整齊排列
- **自動分頁**: 每張投影片最多 3 個圖表，自動建立多張投影片
- **專業標題**: 包含頁碼的格式化標題
- **簡潔設計**: 圖表不顯示個別標題，保持乾淨的視覺效果

### 智慧識別系統

- **標題匹配**: 優先通過圖表標題識別
- **內容分析**: 分析圖表資料範圍關鍵字
- **位置推測**: 根據圖表位置進行智慧推測
- **自動排序**: 按照預定順序排列圖表

## 🔧 常見問題解決

### 問題 1: 權限不足
**症狀**: 執行時出現權限錯誤
**解決**: 
1. 確認已授權所有必要的 API 權限
2. 檢查對目標 Google Sheets 和 Slides 的存取權限

### 問題 2: 找不到分頁
**症狀**: 顯示「找不到分頁」錯誤
**解決**:
1. 確認分頁名稱 `[NR_TX_LMH]Summary&NR_Test_1` 完全正確
2. 檢查 Google Sheets 連結是否可存取

### 問題 3: 沒有找到圖表
**症狀**: 顯示「沒有找到目標圖表」
**解決**:
1. 執行 `testChartFiltering()` 檢查圖表識別
2. 確認分頁中確實有圖表
3. 檢查圖表標題是否包含關鍵字

### 問題 4: 圖表變形
**症狀**: 插入的圖表比例不正確
**解決**:
1. 確認程式碼中 `chartAspectRatio = 5 / 1`
2. 檢查 `maintainAspectRatio: true` 設定
3. 重新執行 `calculateAspectRatioFitDimensions()` 函式

## 📈 進階使用

### 自訂目標圖表

修改 `TARGET_CHART_TITLES` 常數：

```javascript
const TARGET_CHART_TITLES = [
    'MaxPwr - TX Power',
    'MaxPwr - ACLR - Max(E_-1,E_+1)',
    'MaxPwr - EVM',
    // 可以新增更多圖表名稱
];
```

### 調整投影片配置

修改 `MAX_CHARTS_PER_SLIDE` 常數：

```javascript
const MAX_CHARTS_PER_SLIDE = 3; // 每個投影片最多圖表數
```

### 更改圖表比例

修改 `calculateAspectRatioFitDimensions()` 函式中的比例：

```javascript
const chartAspectRatio = 5 / 1; // 5:1 比例
```

## 📞 技術支援

如需更多協助，請：

1. 檢查 Google Apps Script 執行日誌
2. 使用 `Logger.log()` 進行偵錯
3. 執行測試函式確認各元件功能
4. 參考 `README.md` 中的詳細技術文件

---

**最後更新**: 2025-08-12  
**版本**: v3.1.0  
**相容性**: Google Apps Script V8 Runtime
