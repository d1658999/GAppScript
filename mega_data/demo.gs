/**
 * MEGA 資料分析器 - 示範檔案
 * @author GitHub Copilot
 * @permission SpreadsheetApp, SlidesApp
 */

/**
 * 示範函式 - 執行 MEGA 資料分析
 * 這是主要入口點，會呼叫 analyze_mega_data.gs 中的函式
 * @permission SpreadsheetApp, SlidesApp, DriveApp
 */
function runMegaDataAnalysis() {
    try {
        Logger.log('=== MEGA 資料分析器啟動 ===');

        // 先測試連線
        Logger.log('測試連線中...');
        const connectionTest = testConnections();

        if (!connectionTest) {
            throw new Error('連線測試失敗，請檢查 ID 是否正確');
        }

        Logger.log('連線測試成功，開始分析資料...');

        // 執行主要分析
        const result = analyzeMegaData();

        if (result.success) {
            Logger.log('=== 分析完成 ===');
            Logger.log(`建立了 ${result.slidesCreated} 張投影片`);
            Logger.log(`處理了 ${result.chartsProcessed} 個圖表`);

            if (result.slideIds && result.slideIds.length > 0) {
                Logger.log('投影片 ID:');
                result.slideIds.forEach((slideId, index) => {
                    Logger.log(`  投影片 ${index + 1}: ${slideId}`);
                });
            }
        }

        return result;

    } catch (error) {
        Logger.log(`執行失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 快速測試函式
 * @permission SpreadsheetApp, SlidesApp
 */
function quickTest() {
    Logger.log('執行快速測試...');
    return testConnections();
}

/**
 * 測試圖表篩選功能
 * 僅測試圖表識別，不建立投影片
 * @permission SpreadsheetApp
 */
function testChartFiltering() {
    try {
        Logger.log('=== 測試圖表篩選功能 ===');

        // 開啟試算表
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const targetSheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);

        if (!targetSheet) {
            throw new Error(`找不到分頁: ${TARGET_SHEET_NAME}`);
        }

        // 取得所有圖表
        const allCharts = targetSheet.getCharts();
        Logger.log(`總共找到 ${allCharts.length} 個圖表`);

        // 逐一檢查圖表資訊
        for (let i = 0; i < allCharts.length; i++) {
            Logger.log(`--- 圖表 ${i + 1} ---`);
            const chart = allCharts[i];

            try {
                const title = getChartTitle(chart);
                Logger.log(`標題: "${title}"`);

                const ranges = chart.getRanges();
                if (ranges && ranges.length > 0) {
                    Logger.log(`資料範圍: ${ranges.map(r => r.getA1Notation()).join(', ')}`);
                }

                const containerInfo = chart.getContainerInfo();
                Logger.log(`位置: 行 ${containerInfo.getAnchorRow()}, 列 ${containerInfo.getAnchorColumn()}`);

            } catch (error) {
                Logger.log(`檢查圖表 ${i + 1} 時發生錯誤: ${error.message}`);
            }
        }

        // 測試篩選
        const targetCharts = filterTargetCharts(allCharts);
        Logger.log(`=== 篩選結果 ===`);
        Logger.log(`找到 ${targetCharts.length} 個目標圖表:`);

        targetCharts.forEach((chartInfo, index) => {
            Logger.log(`${index + 1}. ${chartInfo.title}`);
        });

        return {
            totalCharts: allCharts.length,
            targetCharts: targetCharts.length,
            chartTitles: targetCharts.map(c => c.title)
        };

    } catch (error) {
        Logger.log(`測試失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 測試多投影片功能
 * 展示當圖表超過 3 個時的處理方式
 * @permission SpreadsheetApp, SlidesApp, DriveApp
 */
function testMultiSlideFeature() {
    try {
        Logger.log('=== 測試多投影片功能 ===');

        // 先測試圖表篩選
        const filterResult = testChartFiltering();

        Logger.log(`目標圖表數量: ${filterResult.targetCharts}`);
        Logger.log(`每張投影片最多圖表數: ${MAX_CHARTS_PER_SLIDE}`);

        const expectedSlides = Math.ceil(filterResult.targetCharts / MAX_CHARTS_PER_SLIDE);
        Logger.log(`預期建立投影片數: ${expectedSlides}`);

        // 如果要實際建立投影片，取消下面的註解
        // const result = analyzeMegaData();
        // Logger.log(`實際建立投影片數: ${result.slidesCreated}`);

        return {
            targetChartsFound: filterResult.targetCharts,
            maxChartsPerSlide: MAX_CHARTS_PER_SLIDE,
            expectedSlides: expectedSlides
        };

    } catch (error) {
        Logger.log(`測試多投影片功能失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 演示原始比例保持功能（固定高度 1.4 英寸）
 * 展示如何計算保持比例的佈局
 */
function demoProportionalLayout() {
    Logger.log('=== 演示原始比例保持功能（固定高度 1.4 英寸）===');

    // 測試不同數量圖表的排版
    for (let chartCount = 1; chartCount <= 3; chartCount++) {
        Logger.log(`--- ${chartCount} 張圖表的排版 ---`);

        const layout = calculateProportionalLayout(chartCount);
        Logger.log(`排版類型: ${layout.arrangement}`);
        Logger.log(`可用區域: ${layout.availableWidth} x ${layout.availableHeight}`);
        Logger.log(`固定圖表高度: ${layout.fixedChartHeight}px (1.4 英寸)`);

        layout.charts.forEach((chart, index) => {
            Logger.log(`圖表 ${index + 1}: 位置(${chart.x}, ${chart.y}), 尺寸${chart.width} x ${chart.height}, 保持比例: ${chart.maintainAspectRatio}`);
        });
    }
}

/**
 * 測試截圖樣式的排版效果
 * 根據提供的截圖來驗證排版是否正確（固定高度 1.4 英寸）
 */
function testScreenshotStyleLayout() {
    Logger.log('=== 測試截圖樣式排版（固定高度 1.4 英寸）===');

    // 模擬三個圖表的排版
    const layout = calculateProportionalLayout(3);

    Logger.log('截圖樣式排版參數:');
    Logger.log(`投影片尺寸: ${layout.slideWidth} x ${layout.slideHeight}`);
    Logger.log(`可用區域: ${layout.availableWidth} x ${layout.availableHeight}`);
    Logger.log(`排版類型: ${layout.arrangement}`);
    Logger.log(`邊距: ${layout.margin}px`);
    Logger.log(`間距: ${layout.spacing}px`);
    Logger.log(`固定圖表高度: ${layout.fixedChartHeight}px (1.4 英寸)`);

    Logger.log('\n三個圖表的位置和尺寸:');
    layout.charts.forEach((chart, index) => {
        const chartTitle = TARGET_CHART_TITLES[index] || `圖表 ${index + 1}`;
        Logger.log(`${index + 1}. ${chartTitle}:`);
        Logger.log(`   位置: (${chart.x}, ${chart.y})`);
        Logger.log(`   尺寸: ${chart.width} x ${chart.height}`);
        Logger.log(`   保持比例: ${chart.maintainAspectRatio}`);
    });

    // 計算實際的圖表尺寸（考慮寬高比）
    Logger.log('\n考慮 5:1 寬高比後的實際尺寸:');
    layout.charts.forEach((chart, index) => {
        const dimensions = calculateAspectRatioFitDimensions(chart.width, chart.height, null);
        const chartTitle = TARGET_CHART_TITLES[index] || `圖表 ${index + 1}`;
        Logger.log(`${index + 1}. ${chartTitle}: ${dimensions.width} x ${dimensions.height}`);
    });

    return layout;
}

/**
 * 演示截圖樣式的標題格式
 */
function demoScreenshotTitleStyle() {
    Logger.log('=== 演示截圖樣式標題 ===');

    Logger.log('主標題樣式:');
    Logger.log('- 字體: Arial, 18px, 粗體');
    Logger.log('- 顏色: #1c4587 (藍色)');
    Logger.log('- 位置: 居中，距頂部 8px');
    Logger.log('- 內容: "MEGA 資料分析 - NR TX LMH 測試結果 (1/2)"');

    Logger.log('\n圖表標題樣式:');
    TARGET_CHART_TITLES.forEach((title, index) => {
        Logger.log(`${index + 1}. "${title}"`);
        Logger.log('   - 字體: Arial, 11px, 粗體');
        Logger.log('   - 顏色: #000000 (黑色)');
        Logger.log('   - 位置: 圖表內部左上角');
    });
}