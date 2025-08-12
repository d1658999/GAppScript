---
applyTo: "./mega_data/*.gs"
---

# Chart 專用規範

* 圖表類型：預設 `LINE`，多資料疊加請一次 `addRange`。
* 圖表標題格式：`<SheetName> – <ChartPurpose> Line Chart`
* 生成後必須 `.setMergeStrategy(Charts.ChartMergeStrategy.MERGE_ROWS)`
* 若 Y 軸單位不同，需使用 `setOption('series', {1: {targetAxisIndex: 1}})`。
* X 軸格式： 參考。 `mega_data/MaxPwr - TX Power.png`