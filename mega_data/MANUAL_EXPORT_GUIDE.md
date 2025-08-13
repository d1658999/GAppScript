# 🔧 手動匯出模擬功能詳細說明

## 🎯 功能概述

為了解決圖片下載後內容被改變的問題，我們開發了「手動匯出模擬功能」，這個功能嘗試模擬在 Google Sheets 中手動右鍵點擊圖表，選擇「Download Chart」→「PNG IMAGE(.PNG)」的過程。

## 🚀 推薦使用方法

### 主要執行函式
```javascript
runManualExportDownload()  // 在 demo.gs 中執行
```

### 完整比較測試
```javascript
compareDownloadMethods()   // 測試所有 5 種方法
```

## 🛠️ 技術實作原理

### 手動匯出模擬過程

我們的 `downloadChartsLikeManualExport()` 函式使用了多層次的方法來模擬手動匯出：

#### 方法 1: 高品質重建
```javascript
// 嘗試重建圖表以取得最高品質版本
const builder = chart.modify();
builder.setOption('width', 1200);   // 設定較高的寬度
builder.setOption('height', 800);   // 設定較高的高度
builder.setOption('backgroundColor', '#ffffff');  // 設定白色背景
const rebuiltChart = builder.build();
const highQualityBlob = rebuiltChart.getBlob();
```

#### 方法 2: 原始 PNG 取得
```javascript
// 直接取得圖表的原始 PNG 資料
const originalBlob = chart.getBlob();
if (originalBlob.getContentType() === 'image/png') {
    return originalBlob;  // 如果已經是 PNG，直接使用
}
```

#### 方法 3: 試算表匯出功能
```javascript
// 使用試算表的內建匯出功能
const containerInfo = chart.getContainerInfo();
const chartBlob = chart.getBlob();
const finalBlob = chartBlob.setName(`chart_export_${Date.now()}.png`);
```

## 📊 與其他方法的比較

| 方法         | 優點                                               | 缺點           | 適用情況     |
| ------------ | -------------------------------------------------- | -------------- | ------------ |
| 手動匯出模擬 | 🌟 最接近手動下載<br>🌟 多重備援方法<br>🌟 高品質重建 | 可能較慢       | **推薦首選** |
| API 方法     | 底層存取<br>技術細節完整                           | 複雜度較高     | 技術用戶     |
| 截圖方法     | 視覺完全一致<br>真實截圖                           | 需要臨時投影片 | 視覺品質優先 |
| 直接方法     | 完全原始資料<br>無轉換                             | 可能有格式限制 | 原始資料需求 |
| 傳統方法     | 簡單快速                                           | 可能有品質損失 | 快速測試     |

## 🔍 品質檢查機制

### 自動品質驗證
```javascript
// 檢查 Blob 的資訊
Logger.log(`原始 Blob 資訊:`);
Logger.log(`  - 內容類型: ${originalBlob.getContentType()}`);
Logger.log(`  - 檔案大小: ${Math.round(originalBlob.getBytes().length / 1024)} KB`);
Logger.log(`  - 原始名稱: ${originalBlob.getName()}`);
```

### 檔案命名規則
- 格式：`{索引}_{圖表名稱}_export_{時間戳記}.png`
- 範例：`1_MaxPwr_TX_Power_export_20250812_143022.png`

## 📝 使用步驟

### 1. 執行手動匯出模擬
```javascript
// 在 Google Apps Script 編輯器中執行
function testManualExport() {
    const result = runManualExportDownload();
    Logger.log('執行結果:', result);
}
```

### 2. 檢查執行結果
查看執行日誌中的：
- ✅ 成功下載的檔案數量
- 📁 Google Drive 資料夾連結
- 📊 每個檔案的品質資訊

### 3. 下載到本機電腦
1. 點擊日誌中的 Google Drive 連結
2. 選擇所有檔案（Ctrl+A）
3. 右鍵選擇「下載」
4. 檔案自動打包為 ZIP

## 🧪 測試和比較

### 全方法比較測試
```javascript
// 執行所有 5 種方法的比較
function runFullComparison() {
    const results = compareDownloadMethods();
    
    // 檢查每個方法的結果
    results.forEach(result => {
        if (result.success) {
            Logger.log(`${result.method}: ✅ 成功`);
            Logger.log(`資料夾: ${result.folderUrl}`);
        } else {
            Logger.log(`${result.method}: ❌ 失敗`);
        }
    });
}
```

### 品質評估指標
1. **檔案大小**：較大通常表示較高品質
2. **內容類型**：PNG 格式最佳
3. **視覺比較**：與原始圖表對照
4. **細節保持**：文字、線條、色彩的清晰度

## ⚠️ 注意事項

### 權限要求
- SpreadsheetApp（讀取圖表）
- DriveApp（儲存檔案）
- UrlFetchApp（API 呼叫，可選）

### 效能考量
- 手動匯出模擬可能需要較長時間
- 建議先用單一方法測試，再使用比較功能
- 大量圖表建議分批處理

### 故障排除
1. **權限錯誤**：確認已授權所有必要權限
2. **圖表不存在**：檢查分頁名稱和圖表識別
3. **下載失敗**：嘗試不同的方法
4. **品質不佳**：使用比較功能找最佳方法

## 🎯 最佳實踐建議

### 首次使用
1. 執行 `quickTest()` 確認連線
2. 執行 `testChartFiltering()` 確認圖表識別
3. 執行 `runManualExportDownload()` 下載圖表
4. 檢查下載品質

### 品質不滿意時
1. 執行 `compareDownloadMethods()` 比較所有方法
2. 檢查每個資料夾的圖片品質
3. 選擇最符合需求的版本
4. 如有需要，調整圖表選項後重新下載

### 批量處理
1. 先測試少量圖表確認方法有效
2. 使用最佳方法處理所有圖表
3. 定期檢查下載進度和品質

## 📞 技術支援

如果手動匯出模擬方法仍無法滿足需求：

1. **檢查原始圖表設定**：確認 Google Sheets 中的圖表設定正確
2. **嘗試其他方法**：使用比較功能測試所有可用方法
3. **調整參數**：修改高品質重建的寬度、高度參數
4. **手動驗證**：與真實手動下載的結果對比

這個功能代表我們最接近真實手動下載過程的技術實現，應該能夠解決大部分圖片品質問題。
