/**
 * MEGA 資料分析器 - 測試執行器
 * 用於驗證程式功能和執行流程
 * @author GitHub Copilot
 * @permission SpreadsheetApp, SlidesApp, DriveApp
 */

/**
 * 完整測試套件 - 按順序執行所有測試
 * @permission SpreadsheetApp, SlidesApp, DriveApp
 */
function runCompleteTestSuite() {
    Logger.log('=== MEGA 資料分析器完整測試套件 ===');
    Logger.log(`執行時間: ${new Date().toLocaleString('zh-TW')}`);
    
    const testResults = {};
    
    try {
        // 測試 1: 基本連線測試
        Logger.log('\n🔗 測試 1: 基本連線測試');
        testResults.connectionTest = runConnectionTest();
        
        // 測試 2: 圖表篩選測試
        Logger.log('\n🎯 測試 2: 圖表篩選測試');
        testResults.filteringTest = runFilteringTest();
        
        // 測試 3: 排版計算測試
        Logger.log('\n📐 測試 3: 排版計算測試');
        testResults.layoutTest = runLayoutTest();
        
        // 測試 4: 配置驗證測試
        Logger.log('\n⚙️ 測試 4: 配置驗證測試');
        testResults.configTest = runConfigTest();
        
        // 總結測試結果
        Logger.log('\n=== 測試結果總結 ===');
        const passedTests = Object.values(testResults).filter(result => result.success).length;
        const totalTests = Object.keys(testResults).length;
        
        Logger.log(`通過測試: ${passedTests}/${totalTests}`);
        if (passedTests === totalTests) {
            Logger.log('✅ 所有測試通過！程式已準備就緒');
        } else {
            Logger.log('❌ 部分測試失敗，請檢查錯誤日誌');
        }
        
        return {
            success: passedTests === totalTests,
            testResults: testResults,
            summary: {
                passed: passedTests,
                total: totalTests
            }
        };
        
    } catch (error) {
        Logger.log(`測試套件執行失敗: ${error.message}`);
        return {
            success: false,
            error: error.message,
            testResults: testResults
        };
    }
}

/**
 * 測試 1: 基本連線測試
 * @permission SpreadsheetApp, SlidesApp
 */
function runConnectionTest() {
    try {
        Logger.log('檢查 Google Sheets 連線...');
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        Logger.log(`✅ Sheets 連線成功: ${spreadsheet.getName()}`);
        
        Logger.log('檢查目標分頁...');
        const targetSheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);
        if (!targetSheet) {
            throw new Error(`找不到分頁: ${TARGET_SHEET_NAME}`);
        }
        Logger.log(`✅ 目標分頁存在: ${TARGET_SHEET_NAME}`);
        
        Logger.log('檢查 Google Slides 連線...');
        const presentation = SlidesApp.openById(PRESENTATION_ID);
        Logger.log(`✅ Slides 連線成功: ${presentation.getName()}`);
        
        return {
            success: true,
            spreadsheetName: spreadsheet.getName(),
            presentationName: presentation.getName(),
            targetSheetExists: true
        };
        
    } catch (error) {
        Logger.log(`❌ 連線測試失敗: ${error.message}`);
        return {
            success: false,
            error: error.message
        };
    }
}

/**
 * 測試 2: 圖表篩選測試
 * @permission SpreadsheetApp
 */
function runFilteringTest() {
    try {
        Logger.log('執行圖表篩選測試...');
        
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const targetSheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);
        const allCharts = targetSheet.getCharts();
        
        Logger.log(`總圖表數量: ${allCharts.length}`);
        
        const targetCharts = filterTargetCharts(allCharts);
        Logger.log(`目標圖表數量: ${targetCharts.length}`);
        
        Logger.log('目標圖表清單:');
        targetCharts.forEach((chart, index) => {
            Logger.log(`  ${index + 1}. ${chart.title}`);
        });
        
        // 驗證是否找到了預期的圖表
        const expectedCharts = TARGET_CHART_TITLES.length;
        const foundCharts = targetCharts.length;
        
        if (foundCharts === 0) {
            Logger.log('⚠️ 警告: 沒有找到任何目標圖表');
        } else if (foundCharts < expectedCharts) {
            Logger.log(`⚠️ 警告: 只找到 ${foundCharts}/${expectedCharts} 個預期圖表`);
        } else {
            Logger.log(`✅ 圖表識別成功`);
        }
        
        return {
            success: foundCharts > 0,
            totalCharts: allCharts.length,
            targetCharts: foundCharts,
            expectedCharts: expectedCharts,
            chartTitles: targetCharts.map(c => c.title)
        };
        
    } catch (error) {
        Logger.log(`❌ 圖表篩選測試失敗: ${error.message}`);
        return {
            success: false,
            error: error.message
        };
    }
}

/**
 * 測試 3: 排版計算測試（固定高度 1.4 英寸）
 */
function runLayoutTest() {
    try {
        Logger.log('執行排版計算測試（固定高度 1.4 英寸）...');
        
        const testResults = [];
        
        // 測試不同數量圖表的排版
        for (let chartCount = 1; chartCount <= 3; chartCount++) {
            Logger.log(`測試 ${chartCount} 張圖表的排版...`);
            
            const layout = calculateProportionalLayout(chartCount);
            
            // 驗證排版合理性
            const isValid = validateLayout(layout, chartCount);
            
            Logger.log(`  排版類型: ${layout.arrangement}`);
            Logger.log(`  可用區域: ${layout.availableWidth} x ${layout.availableHeight}`);
            Logger.log(`  固定圖表高度: ${layout.fixedChartHeight}px (1.4 英寸)`);
            Logger.log(`  圖表配置: ${layout.charts.length} 個`);
            Logger.log(`  驗證結果: ${isValid ? '✅ 通過' : '❌ 失敗'}`);
            
            testResults.push({
                chartCount: chartCount,
                layout: layout,
                isValid: isValid
            });
        }
        
        // 測試 5:1 寬高比計算與固定高度的配合
        Logger.log('測試 5:1 寬高比計算與固定高度配合...');
        const fixedHeight = 100.8; // 1.4 英寸
        const testDimensions = calculateAspectRatioFitDimensions(600, fixedHeight, null);
        const expectedRatio = 5 / 1;
        const actualRatio = testDimensions.width / testDimensions.height;
        const ratioMatch = Math.abs(actualRatio - expectedRatio) < 0.1;
        
        Logger.log(`  固定高度: ${fixedHeight}px (1.4 英寸)`);
        Logger.log(`  期望比例: ${expectedRatio}`);
        Logger.log(`  實際比例: ${actualRatio.toFixed(2)}`);
        Logger.log(`  比例匹配: ${ratioMatch ? '✅ 通過' : '❌ 失敗'}`);
        
        const allLayoutsValid = testResults.every(result => result.isValid);
        
        return {
            success: allLayoutsValid && ratioMatch,
            layoutTests: testResults,
            aspectRatioTest: {
                fixedHeight: fixedHeight,
                expected: expectedRatio,
                actual: actualRatio,
                match: ratioMatch
            }
        };
        
    } catch (error) {
        Logger.log(`❌ 排版計算測試失敗: ${error.message}`);
        return {
            success: false,
            error: error.message
        };
    }
}

/**
 * 測試 4: 配置驗證測試
 */
function runConfigTest() {
    try {
        Logger.log('執行配置驗證測試...');
        
        const configTests = [];
        
        // 檢查必要常數
        const requiredConstants = [
            { name: 'SPREADSHEET_ID', value: SPREADSHEET_ID },
            { name: 'PRESENTATION_ID', value: PRESENTATION_ID },
            { name: 'TARGET_SHEET_NAME', value: TARGET_SHEET_NAME },
            { name: 'TARGET_CHART_TITLES', value: TARGET_CHART_TITLES },
            { name: 'MAX_CHARTS_PER_SLIDE', value: MAX_CHARTS_PER_SLIDE }
        ];
        
        requiredConstants.forEach(constant => {
            const isDefined = constant.value !== undefined && constant.value !== null;
            const isValid = isDefined && (typeof constant.value === 'string' ? constant.value.length > 0 : true);
            
            Logger.log(`  ${constant.name}: ${isValid ? '✅' : '❌'} ${isDefined ? '已定義' : '未定義'}`);
            
            configTests.push({
                name: constant.name,
                isDefined: isDefined,
                isValid: isValid
            });
        });
        
        // 檢查目標圖表標題格式
        Logger.log('檢查目標圖表標題...');
        TARGET_CHART_TITLES.forEach((title, index) => {
            const isValidFormat = title && typeof title === 'string' && title.length > 0;
            Logger.log(`  ${index + 1}. "${title}": ${isValidFormat ? '✅' : '❌'}`);
        });
        
        // 檢查投影片限制
        const validSlideLimit = MAX_CHARTS_PER_SLIDE >= 1 && MAX_CHARTS_PER_SLIDE <= 5;
        Logger.log(`投影片圖表限制 (${MAX_CHARTS_PER_SLIDE}): ${validSlideLimit ? '✅' : '❌'}`);
        
        const allConfigValid = configTests.every(test => test.isValid) && validSlideLimit;
        
        return {
            success: allConfigValid,
            configTests: configTests,
            slideLimit: {
                value: MAX_CHARTS_PER_SLIDE,
                isValid: validSlideLimit
            }
        };
        
    } catch (error) {
        Logger.log(`❌ 配置驗證測試失敗: ${error.message}`);
        return {
            success: false,
            error: error.message
        };
    }
}

/**
 * 驗證排版配置的合理性
 * @param {Object} layout 排版物件
 * @param {number} chartCount 圖表數量
 * @return {boolean} 是否合理
 */
function validateLayout(layout, chartCount) {
    try {
        // 檢查基本屬性
        if (!layout || !layout.charts || !Array.isArray(layout.charts)) {
            return false;
        }
        
        // 檢查圖表數量匹配
        if (layout.charts.length !== chartCount) {
            return false;
        }
        
        // 檢查每個圖表的配置
        for (let chart of layout.charts) {
            // 檢查必要屬性
            if (typeof chart.x !== 'number' || typeof chart.y !== 'number' ||
                typeof chart.width !== 'number' || typeof chart.height !== 'number') {
                return false;
            }
            
            // 檢查尺寸合理性
            if (chart.width <= 0 || chart.height <= 0) {
                return false;
            }
            
            // 檢查位置合理性（不能為負數）
            if (chart.x < 0 || chart.y < 0) {
                return false;
            }
        }
        
        return true;
        
    } catch (error) {
        Logger.log(`排版驗證錯誤: ${error.message}`);
        return false;
    }
}

/**
 * 快速執行主要功能（不建立實際投影片）
 * 用於在不影響現有簡報的情況下測試主要邏輯
 * @permission SpreadsheetApp
 */
function dryRunAnalysis() {
    try {
        Logger.log('=== 乾運行測試（不建立投影片）===');
        
        // 執行所有步驟，除了實際建立投影片
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const targetSheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);
        
        if (!targetSheet) {
            throw new Error(`找不到分頁: ${TARGET_SHEET_NAME}`);
        }
        
        const allCharts = targetSheet.getCharts();
        const targetCharts = filterTargetCharts(allCharts);
        
        Logger.log(`找到 ${targetCharts.length} 個目標圖表`);
        
        if (targetCharts.length > 0) {
            // 模擬圖片轉換（不實際轉換）
            const mockChartData = targetCharts.map(chart => ({
                title: chart.title,
                blob: null, // 不實際產生 blob
                originalIndex: chart.index
            }));
            
            // 計算排版
            const totalSlides = Math.ceil(targetCharts.length / MAX_CHARTS_PER_SLIDE);
            Logger.log(`將建立 ${totalSlides} 張投影片`);
            
            for (let slideIndex = 0; slideIndex < totalSlides; slideIndex++) {
                const startIndex = slideIndex * MAX_CHARTS_PER_SLIDE;
                const endIndex = Math.min(startIndex + MAX_CHARTS_PER_SLIDE, targetCharts.length);
                const chartsForSlide = mockChartData.slice(startIndex, endIndex);
                
                Logger.log(`投影片 ${slideIndex + 1}: ${chartsForSlide.length} 個圖表`);
                chartsForSlide.forEach((chart, index) => {
                    Logger.log(`  ${index + 1}. ${chart.title}`);
                });
                
                // 計算排版
                const layout = calculateProportionalLayout(chartsForSlide.length);
                Logger.log(`  排版: ${layout.arrangement}`);
            }
            
            Logger.log('✅ 乾運行測試完成 - 所有邏輯正常');
            return {
                success: true,
                chartsFound: targetCharts.length,
                slidesNeeded: totalSlides
            };
        } else {
            Logger.log('⚠️ 沒有找到目標圖表');
            return {
                success: false,
                chartsFound: 0,
                slidesNeeded: 0
            };
        }
        
    } catch (error) {
        Logger.log(`❌ 乾運行測試失敗: ${error.message}`);
        return {
            success: false,
            error: error.message
        };
    }
}
