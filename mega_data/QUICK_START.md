# 🚀 MEGA 資料分析器 - 快速執行指南

## 一鍵執行（最簡單的方式）

### 步驟 1: 開啟 Google Apps Script 編輯器
1. 前往您的 Google Apps Script 專案
2. 開啟 `analyze_mega_data.gs` 檔案

### 步驟 2: 執行主要函式
```javascript
// 執行這個函式即可完成所有工作！
executeAnalysis()
```

### 步驟 3: 檢查結果
執行完成後，您會看到類似的日誌：

```
==================================================
開始執行 MEGA 資料分析
==================================================
成功開啟分頁: [NR_TX_LMH]Summary&NR_Test_1
找到 X 個圖表
篩選出 3 個目標圖表
建立了 1 張投影片
==================================================
✅ 分析完成！
✅ 共處理 3 個圖表
✅ 建立 1 張投影片
✅ 請檢查目標 Google Slides 文件：
   https://docs.google.com/presentation/d/1_zV6grTwFms2a0jE2i9Tz0X6mjbnCpZk8VDKIyBLxi8/edit
==================================================
```

## 🔧 疑難排解

### 如果遇到權限問題
第一次執行時，Google Apps Script 會要求授權：
1. 點擊「授權」
2. 選擇您的 Google 帳戶
3. 點擊「允許」授權所有權限

### 如果找不到圖表
執行測試函式檢查：
```javascript
testConnections()  // 測試連線
testChartFiltering()  // 測試圖表識別
```

### 如果圖片品質不佳
程式已自動使用 `insertSheetsChartAsImage` 方法確保最高品質。如果還有問題，檢查：
1. 原始 Google Sheets 中的圖表是否清晰
2. 圖表類型是否為 Google Sheets 內建圖表

## 🎯 功能特色

### ✅ 自動完成所有步驟
- 開啟指定的 Google Sheets 分頁
- 自動識別目標圖表
- 轉換為高品質圖片
- 插入到 Google Slides

### ✅ 精確自訂樣式排版
- 固定高度：1.36 英寸（鎖定寬高比）
- 寬高比：5:1（符合規格要求）
- 精確位置定位：
  - 第一張圖片：X: 0.21", Y: 1.11"
  - 第二張圖片：X: 2.94", Y: 2.55"
  - 第三張圖片：X: 0.21", Y: 3.98"
- 無變形，嚴格保持比例

### ✅ 高品質圖片
- 使用 `insertSheetsChartAsImage` 方法
- 確保圖片與原始圖表完全一致
- 自動清理臨時檔案

### ✅ 零設定需求
- 所有 ID 和設定都已預設
- 不需要修改任何程式碼
- 一鍵執行即可

## 📊 處理的圖表

系統會自動找到並處理這三個圖表：
1. **MaxPwr - TX Power**
2. **MaxPwr - ACLR - Max(E_-1,E_+1)**
3. **MaxPwr - EVM**

## 🔗 連結資訊

- **來源 Google Sheets**: https://docs.google.com/spreadsheets/d/1nnzLzALFvtwW0Lz8aVGCw7O0iaPXwN3_wRUF3BC2iak/edit
- **目標 Google Slides**: https://docs.google.com/presentation/d/1_zV6grTwFms2a0jE2i9Tz0X6mjbnCpZk8VDKIyBLxi8/edit
- **目標分頁**: `[NR_TX_LMH]Summary&NR_Test_1`

## ⚡ 進階選項

如果需要更多控制，可以使用其他函式：

```javascript
// 僅測試連線（不建立投影片）
testConnections()

// 測試自訂樣式排版設定
testCustomStyleLayout()

// 僅測試圖表識別（不建立投影片）
testChartFiltering()

// 完整分析但獲得詳細回傳資訊
const result = analyzeMegaData()
Logger.log(result)
```

## 📝 注意事項

1. **確保有權限存取**指定的 Google Sheets 和 Slides
2. **第一次執行**會需要授權，這是正常的
3. **執行時間**大約 30-60 秒，取決於圖表數量
4. **圖表必須是** Google Sheets 內建圖表，不支援插入的圖片
5. **會建立臨時檔案**但會自動清理
6. **圖片尺寸固定**為 1.36 英寸高，5:1 寬高比
7. **位置精確定位**，不會自動居中或調整

---

**🎉 就這麼簡單！只需要執行 `executeAnalysis()` 就能完成所有工作。**
