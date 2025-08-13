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

/**
 * 執行圖表下載功能
 * 根據 get_chart_google_sheet.prompt.md 要求下載圖表
 * @permission SpreadsheetApp, DriveApp
 */
function runChartDownload() {
    try {
        Logger.log('=== 執行圖表下載功能 ===');
        Logger.log('根據 get_chart_google_sheet.prompt.md 指示執行...');

        // 先測試連線
        Logger.log('測試連線中...');
        const connectionTest = testConnections();

        if (!connectionTest) {
            throw new Error('連線測試失敗，請檢查 ID 是否正確');
        }

        Logger.log('連線測試成功，開始下載圖表...');

        // 執行下載功能
        const result = downloadChartsFromSheet();

        if (result.success) {
            Logger.log('=== 下載完成 ===');
            Logger.log(`下載了 ${result.filesDownloaded} 個圖表檔案`);
            Logger.log(`下載資料夾: ${result.folderName}`);
            Logger.log(`資料夾連結: ${result.folderUrl}`);

            Logger.log('\n下載的檔案:');
            result.files.forEach((file, index) => {
                Logger.log(`${index + 1}. ${file.name}`);
                Logger.log(`   檢視連結: ${file.url}`);
                Logger.log(`   下載連結: ${file.downloadUrl}`);
            });

            Logger.log('\n使用說明:');
            Logger.log('1. 點擊上方資料夾連結開啟 Google Drive');
            Logger.log('2. 選擇所有圖表檔案');
            Logger.log('3. 右鍵選擇「下載」來下載到本機電腦');
        }

        return result;

    } catch (error) {
        Logger.log(`下載失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 批次下載圖表功能
 * 提供更便利的下載方式
 * @permission SpreadsheetApp, DriveApp
 */
function runBatchDownload() {
    try {
        Logger.log('=== 批次下載圖表 ===');

        const result = batchDownloadCharts();

        if (result.success) {
            Logger.log('\n下載完成！請按照以下步驟下載到電腦:');
            result.instructions.forEach((instruction, index) => {
                Logger.log(`${index + 1}. ${instruction}`);
            });
        }

        return result;

    } catch (error) {
        Logger.log(`批次下載失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 列出現有的下載資料夾
 * @permission DriveApp
 */
function listExistingDownloads() {
    try {
        Logger.log('=== 列出現有下載資料夾 ===');

        const folders = listDownloadFolders();

        if (folders.length === 0) {
            Logger.log('沒有找到任何下載資料夾');
        } else {
            Logger.log(`找到 ${folders.length} 個下載資料夾:`);
            folders.forEach((folder, index) => {
                Logger.log(`${index + 1}. ${folder.name}`);
                Logger.log(`   檔案數量: ${folder.fileCount}`);
                Logger.log(`   最後修改: ${folder.lastModified}`);
                Logger.log(`   連結: ${folder.url}`);
            });
        }

        return folders;

    } catch (error) {
        Logger.log(`列出下載資料夾失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 快速完成指示檔案要求的所有步驟
 * 這是主要執行函式，會完成 get_chart_google_sheet.prompt.md 中的所有要求
 * @permission SpreadsheetApp, SlidesApp, DriveApp
 */
function completePromptRequirements() {
    try {
        Logger.log('=== 執行 get_chart_google_sheet.prompt.md 所有要求 ===');

        Logger.log('步驟 1: 開啟 Google Sheets 並導航到指定分頁...');

        // 步驟 1: 開啟 Google Sheets 連結並導航到 [NR_TX_LMH]Summary&NR_Test_1 分頁
        const connectionTest = testConnections();
        if (!connectionTest) {
            throw new Error('無法連接到 Google Sheets');
        }

        Logger.log('步驟 2: 取得分頁中所有圖表的截圖...');

        // 步驟 2: 取得該分頁中所有圖表的截圖
        const chartFilterResult = testChartFiltering();
        Logger.log(`找到 ${chartFilterResult.targetCharts} 個目標圖表`);

        Logger.log('步驟 3: 下載截圖到 Google Drive（可下載到電腦）...');

        // 步驟 3: 下載截圖（到 Google Drive，使用者可再下載到電腦）
        const downloadResult = downloadChartsFromSheet();

        if (downloadResult.success) {
            Logger.log('✅ 所有步驟完成成功！');
            Logger.log(`✅ 已開啟 Google Sheets: ${TARGET_SHEET_NAME} 分頁`);
            Logger.log(`✅ 已取得 ${downloadResult.filesDownloaded} 個圖表截圖`);
            Logger.log(`✅ 已下載到 Google Drive: ${downloadResult.folderName}`);

            Logger.log('\n🔗 Google Drive 資料夾連結:');
            Logger.log(downloadResult.folderUrl);

            Logger.log('\n📥 如何下載到電腦:');
            Logger.log('1. 點擊上方 Google Drive 連結');
            Logger.log('2. 選擇所有圖表檔案（Ctrl+A 或 Cmd+A）');
            Logger.log('3. 右鍵選擇「下載」');
            Logger.log('4. 檔案會自動打包為 ZIP 下載到電腦');

            return {
                success: true,
                step1_completed: true,
                step2_completed: true,
                step3_completed: true,
                chartsFound: chartFilterResult.targetCharts,
                filesDownloaded: downloadResult.filesDownloaded,
                driveFolder: downloadResult.folderUrl,
                downloadInstructions: [
                    '點擊 Google Drive 資料夾連結',
                    '選擇所有圖表檔案',
                    '右鍵選擇「下載」',
                    '檔案將自動下載到電腦'
                ]
            };
        }

    } catch (error) {
        Logger.log(`❌ 執行失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 使用原始截圖方法下載圖表（修正版本）
 * 解決圖片內容被改變的問題
 * @permission SpreadsheetApp, SlidesApp, DriveApp
 */
function runOriginalChartDownload() {
    try {
        Logger.log('=== 使用原始截圖方法下載圖表 ===');
        Logger.log('此方法可確保取得真正的原始圖表，無任何變更...');

        // 先測試連線
        const connectionTest = testConnections();
        if (!connectionTest) {
            throw new Error('連線測試失敗');
        }

        // 使用截圖方法下載
        const result = downloadOriginalChartsViaScreenshot();

        if (result.success) {
            Logger.log('=== 原始圖表下載完成 ===');
            Logger.log(`下載了 ${result.filesDownloaded} 個原始圖表檔案`);
            Logger.log(`下載方法: ${result.method}`);
            Logger.log(`資料夾連結: ${result.folderUrl}`);

            Logger.log('\n下載的檔案:');
            result.files.forEach((file, index) => {
                Logger.log(`${index + 1}. ${file.name}`);
                Logger.log(`   連結: ${file.url}`);
            });

            Logger.log('\n✨ 這些圖片是真正的原始截圖，沒有經過任何轉換或壓縮');
        }

        return result;

    } catch (error) {
        Logger.log(`原始圖表下載失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 使用直接方法下載圖表
 * @permission SpreadsheetApp, DriveApp
 */
function runDirectChartDownload() {
    try {
        Logger.log('=== 使用直接方法下載圖表 ===');
        Logger.log('此方法直接取得圖表的原始 Blob，完全不變更...');

        // 先測試連線
        const connectionTest = testConnections();
        if (!connectionTest) {
            throw new Error('連線測試失敗');
        }

        // 使用直接方法下載
        const result = downloadDirectChartScreenshots();

        if (result.success) {
            Logger.log('=== 直接圖表下載完成 ===');
            Logger.log(`下載了 ${result.filesDownloaded} 個直接圖表檔案`);
            Logger.log(`下載方法: ${result.method}`);
            Logger.log(`資料夾連結: ${result.folderUrl}`);

            Logger.log('\n下載的檔案:');
            result.files.forEach((file, index) => {
                Logger.log(`${index + 1}. ${file.name}`);
                Logger.log(`   連結: ${file.url}`);
            });

            Logger.log('\n✨ 這些是圖表的原始 Blob，保持完全原始格式');
        }

        return result;

    } catch (error) {
        Logger.log(`直接圖表下載失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 比較不同下載方法的差異
 * 這個函式會測試多種下載方法，讓您選擇品質最好的
 * @permission SpreadsheetApp, SlidesApp, DriveApp
 */
function compareDownloadMethods() {
    try {
        Logger.log('=== 比較不同下載方法 ===');
        Logger.log('將測試多種下載方法，幫您找到最好的原始圖片...');

        const results = [];

        // 方法 1: 原始下載方法
        try {
            Logger.log('\n--- 測試方法 1: 原始下載方法 ---');
            const result1 = downloadChartsFromSheet();
            results.push({
                method: '原始下載',
                success: result1.success,
                filesCount: result1.filesDownloaded,
                folderUrl: result1.folderUrl
            });
            Logger.log(`方法 1 完成: ${result1.filesDownloaded} 個檔案`);
        } catch (error) {
            Logger.log(`方法 1 失敗: ${error.message}`);
            results.push({ method: '原始下載', success: false, error: error.message });
        }

        // 方法 2: 截圖方法
        try {
            Logger.log('\n--- 測試方法 2: 截圖方法 ---');
            const result2 = downloadOriginalChartsViaScreenshot();
            results.push({
                method: '截圖方法',
                success: result2.success,
                filesCount: result2.filesDownloaded,
                folderUrl: result2.folderUrl
            });
            Logger.log(`方法 2 完成: ${result2.filesDownloaded} 個檔案`);
        } catch (error) {
            Logger.log(`方法 2 失敗: ${error.message}`);
            results.push({ method: '截圖方法', success: false, error: error.message });
        }

        // 方法 3: 直接方法
        try {
            Logger.log('\n--- 測試方法 3: 直接方法 ---');
            const result3 = downloadDirectChartScreenshots();
            results.push({
                method: '直接方法',
                success: result3.success,
                filesCount: result3.filesDownloaded,
                folderUrl: result3.folderUrl
            });
            Logger.log(`方法 3 完成: ${result3.filesDownloaded} 個檔案`);
        } catch (error) {
            Logger.log(`方法 3 失敗: ${error.message}`);
            results.push({ method: '直接方法', success: false, error: error.message });
        }

        // 方法 4: 手動匯出模擬方法（新增）
        try {
            Logger.log('\n--- 測試方法 4: 手動匯出模擬方法 ---');
            const result4 = downloadChartsLikeManualExport();
            results.push({
                method: '手動匯出模擬',
                success: result4.success,
                filesCount: result4.filesDownloaded,
                folderUrl: result4.folderUrl
            });
            Logger.log(`方法 4 完成: ${result4.filesDownloaded} 個檔案`);
        } catch (error) {
            Logger.log(`方法 4 失敗: ${error.message}`);
            results.push({ method: '手動匯出模擬', success: false, error: error.message });
        }

        // 方法 5: API 方法（新增）
        try {
            Logger.log('\n--- 測試方法 5: API 方法 ---');
            const result5 = downloadChartsViaAPI();
            results.push({
                method: 'API 方法',
                success: result5.success,
                filesCount: result5.filesDownloaded,
                folderUrl: result5.folderUrl
            });
            Logger.log(`方法 5 完成: ${result5.filesDownloaded} 個檔案`);
        } catch (error) {
            Logger.log(`方法 5 失敗: ${error.message}`);
            results.push({ method: 'API 方法', success: false, error: error.message });
        }

        Logger.log('\n=== 下載方法比較結果 ===');
        results.forEach((result, index) => {
            Logger.log(`${index + 1}. ${result.method}:`);
            if (result.success) {
                Logger.log(`   ✅ 成功，下載 ${result.filesCount} 個檔案`);
                Logger.log(`   🔗 資料夾: ${result.folderUrl}`);
            } else {
                Logger.log(`   ❌ 失敗: ${result.error || '未知錯誤'}`);
            }
        });

        Logger.log('\n📝 建議:');
        Logger.log('1. 檢查每個資料夾中的圖片品質');
        Logger.log('2. 選擇最符合原始圖表的版本');
        Logger.log('3. 手動匯出模擬方法最接近您要求的手動下載過程');
        Logger.log('4. 如果圖片仍有差異，請檢查 Google Sheets 中的原始圖表設定');

        return results;

    } catch (error) {
        Logger.log(`比較下載方法失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 使用手動匯出模擬方法下載圖表（新增）
 * 模擬手動右鍵選擇 "Download Chart" → "PNG IMAGE(.PNG)" 的過程
 * @permission SpreadsheetApp, DriveApp, UrlFetchApp
 */
function runManualExportDownload() {
    try {
        Logger.log('=== 使用手動匯出模擬方法下載圖表 ===');
        Logger.log('模擬手動右鍵選擇 "Download Chart" → "PNG IMAGE(.PNG)" 的過程...');

        // 先測試連線
        const connectionTest = testConnections();
        if (!connectionTest) {
            throw new Error('連線測試失敗');
        }

        // 使用手動匯出模擬方法下載
        const result = downloadChartsLikeManualExport();

        if (result.success) {
            Logger.log('=== 手動匯出模擬下載完成 ===');
            Logger.log(`下載了 ${result.filesDownloaded} 個模擬手動匯出的圖表檔案`);
            Logger.log(`下載方法: ${result.method}`);
            Logger.log(`資料夾連結: ${result.folderUrl}`);

            Logger.log('\n下載的檔案:');
            result.files.forEach((file, index) => {
                Logger.log(`${index + 1}. ${file.name}`);
                Logger.log(`   連結: ${file.url}`);
            });

            Logger.log('\n✨ 這些圖片使用了模擬手動匯出的方法，應該最接近您要求的效果');
            Logger.log('📝 如果這個方法的結果仍不滿意，建議使用比較功能測試所有方法');
        }

        return result;

    } catch (error) {
        Logger.log(`手動匯出模擬下載失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 使用 API 方法下載圖表（新增）
 * @permission SpreadsheetApp, DriveApp, UrlFetchApp
 */
function runAPIDownload() {
    try {
        Logger.log('=== 使用 API 方法下載圖表 ===');
        Logger.log('嘗試透過 Google Sheets API 取得圖表...');

        // 先測試連線
        const connectionTest = testConnections();
        if (!connectionTest) {
            throw new Error('連線測試失敗');
        }

        // 使用 API 方法下載
        const result = downloadChartsViaAPI();

        if (result.success) {
            Logger.log('=== API 下載完成 ===');
            Logger.log(`下載了 ${result.filesDownloaded} 個 API 圖表檔案`);
            Logger.log(`下載方法: ${result.method}`);
            Logger.log(`資料夾連結: ${result.folderUrl}`);

            Logger.log('\n下載的檔案:');
            result.files.forEach((file, index) => {
                Logger.log(`${index + 1}. ${file.name}`);
                Logger.log(`   連結: ${file.url}`);
            });

            Logger.log('\n✨ 這些圖片使用了 API 方法取得');
        }

        return result;

    } catch (error) {
        Logger.log(`API 下載失敗: ${error.message}`);
        throw error;
    }
}