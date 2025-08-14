/**
 * MEGA 資料分析器 - 自動從 Google Sheets 擷取圖表並建立 Slides 簡報
 * @author GitHub Copilot
 * @permission SpreadsheetApp, SlidesApp, DriveApp, Utilities
 */

// 常數定義
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const PRESENTATION_ID = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Upload_Link')
    .getRange('B2')
    .getValue();
const TARGET_SHEET_NAME = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Upload_Link')
    .getRange('B1')
    .getValue();
const SLIDE_TITLE = 'MEGA 資料分析 - 測試結果'; // 備用標題
const TITLE_CELL = 'X2'; // 儲存格位置，用於取得投影片標題（根據 prompt 第12點）

// 目標圖表名稱（只處理這些特定圖表）
const TARGET_CHART_TITLES = [
    'MaxPwr - TX Power',
    'MaxPwr - ACLR - Max(E_-1,E_+1)',
    'MaxPwr - EVM'
];

// 投影片配置
const MAX_CHARTS_PER_SLIDE = 3; // 每個投影片最多放三張圖片

// 圖表位置和尺寸配置（根據 prompt 第11點的自訂樣式）
const CHART_STYLE_CONFIG = {
    // 固定高度：1.36 英寸，轉換為點數 (1 英寸 = 72 點)
    FIXED_HEIGHT: 1.36 * 72, // 97.92 點

    // 寬高比：5:1（根據 prompt 第7點）
    ASPECT_RATIO: 5 / 1,

    // 具體位置座標（英寸轉點數）
    POSITIONS: [
        { x: 0.21 * 72, y: 1.11 * 72 }, // 第一張圖片：X: 0.21", Y: 1.11"
        { x: 2.94 * 72, y: 2.55 * 72 }, // 第二張圖片：X: 2.94", Y: 2.55"  
        { x: 0.21 * 72, y: 3.98 * 72 }  // 第三張圖片：X: 0.21", Y: 3.98"
    ]
};

/**
 * 簡化執行函式 - 一鍵執行所有步驟
 * @permission SpreadsheetApp, SlidesApp, DriveApp
 */
function executeAnalysis() {
    Logger.log('='.repeat(50));
    Logger.log('開始執行 MEGA 資料分析');
    Logger.log('='.repeat(50));

    try {
        const result = analyzeMegaData();

        Logger.log('='.repeat(50));
        Logger.log('✅ 分析完成！');
        Logger.log(`✅ 共處理 ${result.chartsProcessed} 個圖表`);
        Logger.log(`✅ 建立 ${result.slidesCreated} 張投影片`);
        Logger.log('✅ 請檢查目標 Google Slides 文件：');
        Logger.log(`   https://docs.google.com/presentation/d/${PRESENTATION_ID}/edit`);
        Logger.log('='.repeat(50));

        return result;

    } catch (error) {
        Logger.log('='.repeat(50));
        Logger.log('❌ 執行失敗！');
        Logger.log(`❌ 錯誤: ${error.message}`);
        Logger.log('='.repeat(50));
        throw error;
    }
}

/**
 * 讀取指定儲存格的內容作為投影片標題（根據 prompt 第12點）
 * @param {Sheet} sheet Google Sheets 工作表物件
 * @param {string} cellAddress 儲存格地址（如 'X2'）
 * @return {string} 儲存格內容，如果為空則使用備用標題
 * @permission SpreadsheetApp
 */
function getSlideTitle(sheet, cellAddress = TITLE_CELL) {
    try {
        const cellValue = sheet.getRange(cellAddress).getValue();

        // 檢查儲存格內容
        if (cellValue && cellValue.toString().trim() !== '') {
            const title = cellValue.toString().trim();
            Logger.log(`從儲存格 ${cellAddress} 讀取標題: "${title}"`);
            return title;
        } else {
            Logger.log(`儲存格 ${cellAddress} 為空，使用備用標題`);
            return SLIDE_TITLE;
        }
    } catch (error) {
        Logger.log(`讀取儲存格 ${cellAddress} 失敗: ${error.message}，使用備用標題`);
        return SLIDE_TITLE;
    }
}

/**
 * 主要執行函式 - 分析 MEGA 資料並建立投影片
 * @permission SpreadsheetApp, SlidesApp, DriveApp
 */
function analyzeMegaData() {
    try {
        Logger.log('開始執行 MEGA 資料分析...');

        // 步驟 1: 開啟 Google Sheets 並取得目標分頁
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const targetSheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);

        if (!targetSheet) {
            throw new Error(`找不到分頁: ${TARGET_SHEET_NAME}`);
        }

        Logger.log(`成功開啟分頁: ${TARGET_SHEET_NAME}`);

        // 步驟 1.5: 讀取儲存格 X2 的內容作為投影片標題（根據 prompt 第12點）
        const slideTitle = getSlideTitle(targetSheet);

        // 步驟 2: 取得分頁中的所有圖表並篩選目標圖表
        const allCharts = targetSheet.getCharts();
        Logger.log(`找到 ${allCharts.length} 個圖表`);

        if (allCharts.length === 0) {
            Logger.log('警告: 分頁中沒有發現任何圖表');
            return;
        }

        // 篩選出目標圖表
        const targetCharts = filterTargetCharts(allCharts);
        Logger.log(`篩選出 ${targetCharts.length} 個目標圖表`);

        if (targetCharts.length === 0) {
            Logger.log('警告: 沒有找到指定的目標圖表');
            Logger.log(`目標圖表名稱: ${TARGET_CHART_TITLES.join(', ')}`);
            return;
        }

        // 步驟 3: 將目標圖表轉換為圖片
        const chartBlobs = exportChartsAsImages(targetCharts);

        // 步驟 4: 開啟 Google Slides 並建立新投影片
        const presentation = SlidesApp.openById(PRESENTATION_ID);
        const slides = createSlidesWithCharts(presentation, chartBlobs, slideTitle);

        Logger.log('MEGA 資料分析完成！');
        Logger.log(`建立了 ${slides.length} 張投影片`);
        slides.forEach((slide, index) => {
            Logger.log(`投影片 ${index + 1} ID: ${slide.getObjectId()}`);
        });

        return {
            success: true,
            slidesCreated: slides.length,
            slideIds: slides.map(slide => slide.getObjectId()),
            chartsProcessed: targetCharts.length
        };

    } catch (error) {
        Logger.log(`錯誤: ${error.message}`);
        throw error;
    }
}

/**
 * 篩選出目標圖表
 * @param {EmbeddedChart[]} allCharts 所有圖表陣列
 * @return {Object[]} 包含圖表和標題的物件陣列
 * @permission SpreadsheetApp
 */
function filterTargetCharts(allCharts) {
    const targetCharts = [];

    for (let i = 0; i < allCharts.length; i++) {
        try {
            const chart = allCharts[i];
            const chartTitle = getChartTitle(chart);

            Logger.log(`檢查圖表 ${i + 1}: "${chartTitle}"`);

            // 檢查是否為目標圖表
            let isTargetChart = TARGET_CHART_TITLES.some(targetTitle =>
                chartTitle && chartTitle.includes(targetTitle)
            );

            // 如果通過標題無法識別，嘗試通過位置和內容識別
            if (!isTargetChart) {
                isTargetChart = identifyChartByContent(chart, i);
            }

            if (isTargetChart) {
                // 嘗試取得更準確的標題
                const refinedTitle = refineChartTitle(chart, chartTitle, i);

                targetCharts.push({
                    chart: chart,
                    title: refinedTitle,
                    index: i
                });
                Logger.log(`✓ 找到目標圖表: "${refinedTitle}"`);
            }

        } catch (error) {
            Logger.log(`檢查圖表 ${i + 1} 時發生錯誤: ${error.message}`);
        }
    }

    // 按照TARGET_CHART_TITLES的順序排序
    targetCharts.sort((a, b) => {
        const indexA = TARGET_CHART_TITLES.findIndex(title => a.title.includes(title));
        const indexB = TARGET_CHART_TITLES.findIndex(title => b.title.includes(title));
        if (indexA === -1 && indexB === -1) return 0;
        if (indexA === -1) return 1;
        if (indexB === -1) return -1;
        return indexA - indexB;
    });

    return targetCharts;
}

/**
 * 通過內容識別圖表
 * @param {EmbeddedChart} chart 圖表物件
 * @param {number} index 圖表索引
 * @return {boolean} 是否為目標圖表
 * @permission SpreadsheetApp
 */
function identifyChartByContent(chart, index) {
    try {
        // 嘗試通過圖表的資料範圍來識別
        const ranges = chart.getRanges();
        if (ranges && ranges.length > 0) {
            for (let range of ranges) {
                const values = range.getValues();
                const displayValues = range.getDisplayValues();

                // 檢查範圍內是否包含目標關鍵字
                const allText = values.concat(displayValues).flat().join(' ').toLowerCase();

                const hasTargetKeywords = TARGET_CHART_TITLES.some(title => {
                    const keywords = title.toLowerCase().split(/[\s-]+/);
                    return keywords.some(keyword => allText.includes(keyword));
                });

                if (hasTargetKeywords) {
                    Logger.log(`通過內容識別到目標圖表 ${index + 1}`);
                    return true;
                }
            }
        }

        // 如果前三個圖表都識別不出來，假設它們就是我們要的
        if (index < 3) {
            Logger.log(`假設圖表 ${index + 1} 為目標圖表（位置推測）`);
            return true;
        }

    } catch (error) {
        Logger.log(`識別圖表內容時發生錯誤: ${error.message}`);
    }

    return false;
}

/**
 * 精確化圖表標題（根據截圖確保正確標題）
 * @param {EmbeddedChart} chart 圖表物件
 * @param {string} originalTitle 原始標題
 * @param {number} index 圖表索引
 * @return {string} 精確化的標題
 * @permission SpreadsheetApp
 */
function refineChartTitle(chart, originalTitle, index) {
    // 如果原始標題已經包含目標關鍵字，直接使用
    const matchedTarget = TARGET_CHART_TITLES.find(target =>
        originalTitle.includes(target)
    );

    if (matchedTarget) {
        return originalTitle;
    }

    // 根據位置推測標題，確保與截圖中的標題一致
    if (index < TARGET_CHART_TITLES.length) {
        return TARGET_CHART_TITLES[index];
    }

    // 嘗試通過資料範圍推測
    try {
        const ranges = chart.getRanges();
        if (ranges && ranges.length > 0) {
            const range = ranges[0];
            const headers = range.offset(0, 0, 1, range.getNumColumns()).getDisplayValues()[0];

            for (let target of TARGET_CHART_TITLES) {
                const keywords = target.toLowerCase().split(/[\s-]+/);
                const headerText = headers.join(' ').toLowerCase();

                if (keywords.some(keyword => headerText.includes(keyword))) {
                    return target;
                }
            }
        }
    } catch (error) {
        Logger.log(`推測圖表標題時發生錯誤: ${error.message}`);
    }

    // 如果無法確定，使用預設標題格式（與截圖一致）
    const fallbackTitles = [
        'MaxPwr - TX Power',
        'MaxPwr - ACLR - Max(E_-1,E_+1)',
        'MaxPwr - EVM'
    ];

    return fallbackTitles[index] || originalTitle || `圖表 ${index + 1}`;
}/**
 * 取得圖表標題
 * @param {EmbeddedChart} chart 圖表物件
 * @return {string} 圖表標題
 * @permission SpreadsheetApp
 */
function getChartTitle(chart) {
    try {
        // 方法 1: 嘗試從圖表選項中取得標題
        const options = chart.getOptions();
        if (options && options.title) {
            return options.title;
        }

        // 方法 2: 嘗試從圖表建構器取得標題
        try {
            const chartBuilder = chart.modify();
            const builtChart = chartBuilder.build();
            const title = builtChart.getOptions().title;
            if (title) {
                return title;
            }
        } catch (e) {
            // 忽略這個錯誤，嘗試下一個方法
        }

        // 方法 3: 嘗試從圖表範圍推測標題
        try {
            const ranges = chart.getRanges();
            if (ranges && ranges.length > 0) {
                const range = ranges[0];
                const sheet = range.getSheet();
                const sheetName = sheet.getName();
                return `${sheetName} 圖表`;
            }
        } catch (e) {
            // 忽略這個錯誤
        }

        // 方法 4: 使用圖表位置資訊
        try {
            const containerInfo = chart.getContainerInfo();
            const anchorRow = containerInfo.getAnchorRow();
            const anchorCol = containerInfo.getAnchorColumn();
            return `圖表_R${anchorRow}C${anchorCol}`;
        } catch (e) {
            // 忽略這個錯誤
        }

        // 如果所有方法都失敗，回傳預設名稱
        return `未知圖表_${Date.now()}`;

    } catch (error) {
        Logger.log(`取得圖表標題失敗: ${error.message}`);
        return `錯誤圖表_${Date.now()}`;
    }
}

/**
 * 將圖表匯出為圖片檔（使用 insertSheetsChartAsImage 方法確保品質）
 * @param {Object[]} targetCharts 目標圖表物件陣列
 * @return {Object[]} 包含圖片 Blob 和標題的物件陣列
 * @permission SlidesApp, DriveApp
 */
function exportChartsAsImages(targetCharts) {
    const chartData = [];

    for (let i = 0; i < targetCharts.length; i++) {
        try {
            const chartInfo = targetCharts[i];
            Logger.log(`處理圖表: "${chartInfo.title}"`);

            // 使用建議的方法：建立臨時 Slides 來匯出高品質圖片
            const chartBlob = exportChartAsImageCorrectly(chartInfo.chart, chartInfo.title);

            if (chartBlob) {
                chartData.push({
                    blob: chartBlob,
                    title: chartInfo.title,
                    originalIndex: chartInfo.index
                });

                Logger.log(`圖表 "${chartInfo.title}" 轉換完成`);
            } else {
                Logger.log(`圖表 "${chartInfo.title}" 轉換失敗 - 無法取得圖片`);
            }

        } catch (error) {
            Logger.log(`圖表 "${targetCharts[i].title}" 處理失敗: ${error.message}`);

            // 如果高品質方法失敗，回退到基本方法
            try {
                const fallbackBlob = targetCharts[i].chart.getBlob();
                chartData.push({
                    blob: fallbackBlob,
                    title: targetCharts[i].title,
                    originalIndex: targetCharts[i].index
                });
                Logger.log(`圖表 "${targetCharts[i].title}" 使用回退方法轉換完成`);
            } catch (fallbackError) {
                Logger.log(`圖表 "${targetCharts[i].title}" 回退方法也失敗: ${fallbackError.message}`);
            }
        }
    }

    return chartData;
}

/**
 * 使用 insertSheetsChartAsImage 方法正確匯出圖表
 * @param {EmbeddedChart} chart 圖表物件
 * @param {string} chartTitle 圖表標題
 * @return {Blob} 圖表的圖片 Blob
 * @permission SlidesApp, DriveApp
 */
function exportChartAsImageCorrectly(chart, chartTitle) {
    let tempSlides = null;

    try {
        // 建立臨時的 Google Slides
        const tempName = `temp_chart_export_${Date.now()}`;
        tempSlides = SlidesApp.create(tempName);
        Logger.log(`建立臨時 Slides: ${tempName}`);

        // 將圖表插入到 Slides 中並轉換為圖片
        const slide = tempSlides.getSlides()[0];
        const chartImage = slide.insertSheetsChartAsImage(chart);

        // 取得圖片 Blob
        const imageBlob = chartImage.getAs("image/png");

        Logger.log(`圖表 "${chartTitle}" 成功透過 insertSheetsChartAsImage 匯出`);

        return imageBlob;

    } catch (error) {
        Logger.log(`使用 insertSheetsChartAsImage 匯出圖表失敗: ${error.message}`);
        return null;
    } finally {
        // 清理臨時 Slides 檔案
        if (tempSlides) {
            try {
                DriveApp.getFileById(tempSlides.getId()).setTrashed(true);
                Logger.log(`已刪除臨時 Slides 檔案`);
            } catch (cleanupError) {
                Logger.log(`刪除臨時檔案失敗: ${cleanupError.message}`);
            }
        }
    }
}

/**
 * 建立投影片並插入圖表（支援多張投影片）
 * @param {Presentation} presentation Google Slides 簡報物件
 * @param {Object[]} chartData 包含圖片 Blob 和標題的物件陣列
 * @param {string} customTitle 自訂投影片標題（從儲存格 X2 讀取）
 * @return {Slide[]} 建立的投影片陣列
 * @permission SlidesApp
 */
function createSlidesWithCharts(presentation, chartData, customTitle = null) {
    const slides = [];
    const totalCharts = chartData.length;

    // 計算需要多少張投影片
    const totalSlides = Math.ceil(totalCharts / MAX_CHARTS_PER_SLIDE);
    Logger.log(`需要建立 ${totalSlides} 張投影片來容納 ${totalCharts} 個圖表`);

    for (let slideIndex = 0; slideIndex < totalSlides; slideIndex++) {
        const startIndex = slideIndex * MAX_CHARTS_PER_SLIDE;
        const endIndex = Math.min(startIndex + MAX_CHARTS_PER_SLIDE, totalCharts);
        const chartsForThisSlide = chartData.slice(startIndex, endIndex);

        Logger.log(`建立投影片 ${slideIndex + 1}，包含圖表 ${startIndex + 1} 到 ${endIndex}`);

        // 建立新投影片
        const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

        // 添加標題（使用自訂標題或預設標題）
        addSlideTitle(slide, slideIndex + 1, totalSlides, customTitle);

        // 計算保持原始比例的排版
        const layout = calculateProportionalLayout(chartsForThisSlide.length);

        // 插入圖表並保持原始比例
        insertChartsWithOriginalProportions(slide, chartsForThisSlide, layout);

        slides.push(slide);
    }

    return slides;
}

/**
 * 在簡報中建立新投影片並插入圖表（舊版相容性）
 * @param {Presentation} presentation Google Slides 簡報物件
 * @param {Object[]} chartData 包含圖片 Blob 和標題的物件陣列
 * @return {Slide} 新建立的投影片
 * @permission SlidesApp
 */
function createNewSlideWithCharts(presentation, chartData) {
    const slides = createSlidesWithCharts(presentation, chartData);
    return slides[0]; // 回傳第一張投影片以保持相容性
}

/**
 * 為投影片添加標題（使用儲存格 X2 的內容，根據 prompt 第12點）
 * @param {Slide} slide 投影片物件
 * @param {number} slideNumber 投影片編號（可選）
 * @param {number} totalSlides 總投影片數（可選）
 * @param {string} customTitle 自訂標題（從儲存格 X2 讀取）
 * @permission SlidesApp
 */
function addSlideTitle(slide, slideNumber = 1, totalSlides = 1, customTitle = null) {
    // 使用自訂標題或預設標題
    let baseTitle = customTitle || SLIDE_TITLE;
    let title = baseTitle;

    // 如果有多張投影片，在標題中加入頁碼
    if (totalSlides > 1) {
        title = `${baseTitle} (${slideNumber}/${totalSlides})`;
    }

    Logger.log(`設定投影片標題: "${title}"`);

    const titleBox = slide.insertTextBox(title);
    titleBox.setLeft(20);          // 稍微縮進
    titleBox.setTop(8);            // 更靠近頂部
    titleBox.setWidth(680);        // 調整寬度
    titleBox.setHeight(35);        // 縮小高度

    const titleText = titleBox.getText();
    titleText.getTextStyle()
        .setFontSize(18)             // 根據截圖調整字體大小
        .setFontFamily('Arial')
        .setBold(true)
        .setForegroundColor('#1c4587'); // 保持藍色，如截圖

    // 文字置中
    titleText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
}

/**
 * 計算自訂樣式排版（根據 prompt 第11點的具體要求）
 * @param {number} imageCount 圖片數量（最多3張）
 * @return {Object} 排版資訊
 */
function calculateProportionalLayout(imageCount) {
    // 使用固定高度和寬高比計算寬度
    const chartHeight = CHART_STYLE_CONFIG.FIXED_HEIGHT;
    const chartWidth = chartHeight * CHART_STYLE_CONFIG.ASPECT_RATIO;

    const charts = [];

    // 根據圖片數量建立對應的位置配置
    for (let i = 0; i < imageCount && i < CHART_STYLE_CONFIG.POSITIONS.length; i++) {
        const position = CHART_STYLE_CONFIG.POSITIONS[i];

        charts.push({
            x: position.x,
            y: position.y,
            width: chartWidth,
            height: chartHeight,
            maintainAspectRatio: true,
            index: i + 1
        });
    }

    const layout = {
        arrangement: `custom_${imageCount}`,
        charts: charts,
        slideWidth: 720,        // 標準投影片寬度
        slideHeight: 540,       // 標準投影片高度
        chartWidth: chartWidth,
        chartHeight: chartHeight,
        aspectRatio: CHART_STYLE_CONFIG.ASPECT_RATIO
    };

    Logger.log(`自訂樣式排版 - 圖片數量: ${imageCount}`);
    Logger.log(`圖表尺寸: ${chartWidth.toFixed(1)}x${chartHeight.toFixed(1)} 點 (${(chartWidth / 72).toFixed(2)}"x${(chartHeight / 72).toFixed(2)}")`);
    Logger.log(`寬高比: ${CHART_STYLE_CONFIG.ASPECT_RATIO}:1`);

    charts.forEach((chart, index) => {
        Logger.log(`圖片 ${index + 1} 位置: (${(chart.x / 72).toFixed(2)}", ${(chart.y / 72).toFixed(2)}")`);
    });

    return layout;
}/**
 * 插入圖表並使用自訂樣式排版（根據 prompt 第11點）
 * @param {Slide} slide 投影片物件
 * @param {Object[]} chartData 包含圖片 Blob 和標題的物件陣列
 * @param {Object} layout 排版資訊
 * @permission SlidesApp
 */
function insertChartsWithOriginalProportions(slide, chartData, layout) {
    for (let i = 0; i < chartData.length && i < layout.charts.length; i++) {
        const chartInfo = chartData[i];
        const chartLayout = layout.charts[i];

        try {
            Logger.log(`插入圖表 "${chartInfo.title}"`);

            // 插入圖片
            const image = slide.insertImage(chartInfo.blob);

            // 使用自訂樣式設定：固定尺寸和位置
            image.setLeft(chartLayout.x);
            image.setTop(chartLayout.y);
            image.setWidth(chartLayout.width);
            image.setHeight(chartLayout.height);

            Logger.log(`圖表 "${chartInfo.title}" 已插入`);
            Logger.log(`  ➤ 位置: (${(chartLayout.x / 72).toFixed(2)}", ${(chartLayout.y / 72).toFixed(2)}")`);
            Logger.log(`  ➤ 尺寸: ${(chartLayout.width / 72).toFixed(2)}"x${(chartLayout.height / 72).toFixed(2)}" (${chartLayout.width.toFixed(1)}x${chartLayout.height.toFixed(1)} 點)`);
            Logger.log(`  ➤ 寬高比: ${(chartLayout.width / chartLayout.height).toFixed(2)}:1`);

            // 不顯示圖表標題（根據需求第8點）
            // addChartLabelOptimized(slide, chartInfo.title, chartLayout, i);

        } catch (error) {
            Logger.log(`插入圖表 "${chartInfo.title}" 失敗: ${error.message}`);
        }
    }
}

/**
 * 計算保持寬高比的適配尺寸（已棄用 - 現在使用固定 5:1 比例）
 * @deprecated 不再使用，改為使用固定 5:1 寬高比
 * @param {number} maxWidth 最大寬度
 * @param {number} maxHeight 最大高度
 * @param {Blob} imageBlob 圖片 Blob
 * @return {Object} 包含 width 和 height 的物件
 */
function calculateAspectRatioFitDimensions(maxWidth, maxHeight, imageBlob) {
    Logger.log('警告: calculateAspectRatioFitDimensions 已棄用，現在使用固定 5:1 比例');

    // 固定使用 5:1 寬高比
    const chartAspectRatio = 5 / 1;

    const widthBasedHeight = maxWidth / chartAspectRatio;
    const heightBasedWidth = maxHeight * chartAspectRatio;

    if (widthBasedHeight <= maxHeight) {
        return {
            width: maxWidth,
            height: widthBasedHeight
        };
    } else {
        return {
            width: heightBasedWidth,
            height: maxHeight
        };
    }
}/**
 * 優化的圖表標籤添加（已停用 - 不顯示圖表標題）
 * @param {Slide} slide 投影片物件
 * @param {string} title 圖表標題
 * @param {Object} chartLayout 圖表排版資訊
 * @param {number} index 圖表索引
 * @permission SlidesApp
 */
function addChartLabelOptimized(slide, title, chartLayout, index) {
    // 功能已停用 - 不在投影片中顯示圖表標題
    Logger.log(`圖表標籤功能已停用 - 跳過 "${title}" 標題添加`);
    return;
}

/**
 * 計算垂直排版的佈局（已棄用，保留以維持相容性）
 * @deprecated 請使用 calculateProportionalLayout
 * @param {number} imageCount 圖片數量
 * @return {Object} 垂直排版資訊
 */
function calculateVerticalLayout(imageCount) {
    Logger.log('警告: calculateVerticalLayout 已棄用，請使用 calculateProportionalLayout');

    const slideWidth = 720;
    const slideHeight = 540;
    const titleHeight = 60;
    const margin = 15;
    const verticalSpacing = 10;

    // 可用區域
    const availableWidth = slideWidth - (2 * margin);
    const availableHeight = slideHeight - titleHeight - (2 * margin);

    // 垂直排列：一欄多列
    const cols = 1;
    const rows = imageCount;

    // 計算每個圖表的尺寸（保持寬高比）
    const chartHeight = (availableHeight - ((rows - 1) * verticalSpacing)) / rows;
    const chartWidth = availableWidth;

    // 如果圖表太高，調整尺寸保持比例
    const maxChartHeight = availableHeight / rows * 0.8; // 留一些間距
    const finalHeight = Math.min(chartHeight, maxChartHeight);
    const finalWidth = chartWidth;

    return {
        cols: cols,
        rows: rows,
        imageWidth: finalWidth,
        imageHeight: finalHeight,
        startX: margin,
        startY: titleHeight + margin,
        verticalSpacing: verticalSpacing,
        availableWidth: availableWidth,
        availableHeight: availableHeight
    };
}

/**
 * 垂直插入圖表到投影片（已棄用，保留以維持相容性）
 * @deprecated 請使用 insertChartsWithOriginalProportions
 * @param {Slide} slide 投影片物件
 * @param {Object[]} chartData 包含圖片 Blob 和標題的物件陣列
 * @param {Object} layout 排版資訊
 * @permission SlidesApp
 */
function insertChartsVertically(slide, chartData, layout) {
    Logger.log('警告: insertChartsVertically 已棄用，請使用 insertChartsWithOriginalProportions');

    for (let i = 0; i < chartData.length; i++) {
        const chartInfo = chartData[i];

        // 計算垂直位置
        const y = layout.startY + (i * (layout.imageHeight + layout.verticalSpacing));
        const x = layout.startX;

        try {
            // 插入圖片
            const image = slide.insertImage(chartInfo.blob);
            image.setLeft(x);
            image.setTop(y);
            image.setWidth(layout.imageWidth);
            image.setHeight(layout.imageHeight);

            // 添加圖表標題（在圖片下方或旁邊）
            addChartLabel(slide, chartInfo.title, x, y, layout);

            Logger.log(`圖表 "${chartInfo.title}" 已插入投影片 (位置: ${x}, ${y})`);

        } catch (error) {
            Logger.log(`插入圖表 "${chartInfo.title}" 失敗: ${error.message}`);
        }
    }
}

/**
 * 為圖表添加標籤（已棄用，保留以維持相容性）
 * @deprecated 請使用 addChartLabelOptimized
 * @param {Slide} slide 投影片物件
 * @param {string} title 圖表標題
 * @param {number} x 圖表 X 位置
 * @param {number} y 圖表 Y 位置
 * @param {Object} layout 排版資訊
 * @permission SlidesApp
 */
function addChartLabel(slide, title, x, y, layout) {
    Logger.log('警告: addChartLabel 已棄用，請使用 addChartLabelOptimized');

    try {
        // 在圖表上方添加小標籤
        const labelY = y - 25;
        const labelHeight = 20;

        if (labelY >= layout.startY - 25) { // 確保標籤不會超出邊界
            const labelBox = slide.insertTextBox(title);
            labelBox.setLeft(x);
            labelBox.setTop(labelY);
            labelBox.setWidth(layout.imageWidth);
            labelBox.setHeight(labelHeight);

            const labelText = labelBox.getText();
            labelText.getTextStyle()
                .setFontSize(10)
                .setFontFamily('Arial')
                .setBold(false)
                .setForegroundColor('#333333');

            // 文字置中
            labelText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
        }

    } catch (error) {
        Logger.log(`添加圖表標籤失敗: ${error.message}`);
    }
}

/**
 * 測試儲存格 X2 標題讀取功能（根據 prompt 第12點）
 * @permission SpreadsheetApp
 */
function testSlideTitle() {
    Logger.log('='.repeat(50));
    Logger.log('測試投影片標題功能（儲存格 X2）');
    Logger.log('='.repeat(50));

    try {
        // 開啟 Google Sheets 並取得目標分頁
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const targetSheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);

        if (!targetSheet) {
            throw new Error(`找不到分頁: ${TARGET_SHEET_NAME}`);
        }

        Logger.log(`✅ 成功開啟分頁: ${TARGET_SHEET_NAME}`);

        // 測試讀取儲存格 X2
        Logger.log(`🎯 嘗試讀取儲存格 ${TITLE_CELL} 的內容...`);

        const slideTitle = getSlideTitle(targetSheet);

        Logger.log(`📝 標題結果: "${slideTitle}"`);
        Logger.log(`📋 備用標題: "${SLIDE_TITLE}"`);

        // 測試不同投影片數量的標題格式
        Logger.log('\n📄 測試標題格式：');
        Logger.log(`單張投影片: "${slideTitle}"`);
        Logger.log(`多張投影片 (1/2): "${slideTitle} (1/2)"`);
        Logger.log(`多張投影片 (2/3): "${slideTitle} (2/3)"`);

        Logger.log('\n✅ 投影片標題測試完成');
        return slideTitle;

    } catch (error) {
        Logger.log(`❌ 測試失敗: ${error.message}`);
        throw error;
    } finally {
        Logger.log('='.repeat(50));
    }
}

/**
 * 測試自訂樣式排版 - 驗證新的位置和尺寸設定
 * @permission None
 */
function testCustomStyleLayout() {
    Logger.log('='.repeat(50));
    Logger.log('測試自訂樣式排版設定');
    Logger.log('='.repeat(50));

    Logger.log('🎯 自訂樣式配置：');
    Logger.log(`固定高度: ${CHART_STYLE_CONFIG.FIXED_HEIGHT.toFixed(1)} 點 (${(CHART_STYLE_CONFIG.FIXED_HEIGHT / 72).toFixed(2)} 英寸)`);
    Logger.log(`寬高比: ${CHART_STYLE_CONFIG.ASPECT_RATIO}:1`);
    Logger.log(`計算寬度: ${(CHART_STYLE_CONFIG.FIXED_HEIGHT * CHART_STYLE_CONFIG.ASPECT_RATIO).toFixed(1)} 點 (${((CHART_STYLE_CONFIG.FIXED_HEIGHT * CHART_STYLE_CONFIG.ASPECT_RATIO) / 72).toFixed(2)} 英寸)`);

    Logger.log('\n📍 圖片位置設定：');
    CHART_STYLE_CONFIG.POSITIONS.forEach((pos, index) => {
        Logger.log(`圖片 ${index + 1}: X=${(pos.x / 72).toFixed(2)}" (${pos.x.toFixed(1)}點), Y=${(pos.y / 72).toFixed(2)}" (${pos.y.toFixed(1)}點)`);
    });

    Logger.log('\n🧪 測試不同圖片數量的排版：');

    for (let imageCount = 1; imageCount <= 3; imageCount++) {
        Logger.log(`\n--- ${imageCount} 張圖片的排版 ---`);
        const layout = calculateProportionalLayout(imageCount);

        Logger.log(`排版類型: ${layout.arrangement}`);
        Logger.log(`使用位置數量: ${layout.charts.length}`);

        layout.charts.forEach((chart, index) => {
            Logger.log(`  圖片 ${index + 1}:`);
            Logger.log(`    位置: (${(chart.x / 72).toFixed(2)}", ${(chart.y / 72).toFixed(2)}")`);
            Logger.log(`    尺寸: ${(chart.width / 72).toFixed(2)}"x${(chart.height / 72).toFixed(2)}"`);
        });
    }

    Logger.log('\n✅ 自訂樣式排版測試完成');
    Logger.log('='.repeat(50));
}

/**
 * 測試函式 - 檢查連線和權限
 * @permission SpreadsheetApp, SlidesApp
 */
function testConnections() {
    try {
        // 測試 Spreadsheet 連線
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        Logger.log(`成功連接到試算表: ${spreadsheet.getName()}`);

        // 測試目標分頁
        const targetSheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);
        if (targetSheet) {
            Logger.log(`找到目標分頁: ${TARGET_SHEET_NAME}`);

            const allCharts = targetSheet.getCharts();
            Logger.log(`分頁中有 ${allCharts.length} 個圖表`);

            // 測試圖表篩選
            const targetCharts = filterTargetCharts(allCharts);
            Logger.log(`找到 ${targetCharts.length} 個目標圖表:`);

            targetCharts.forEach((chartInfo, index) => {
                Logger.log(`  ${index + 1}. ${chartInfo.title}`);
            });

        } else {
            Logger.log(`警告: 找不到分頁 ${TARGET_SHEET_NAME}`);
        }

        // 測試 Slides 連線
        const presentation = SlidesApp.openById(PRESENTATION_ID);
        Logger.log(`成功連接到簡報: ${presentation.getName()}`);
        Logger.log(`目前有 ${presentation.getSlides().length} 張投影片`);

        return true;

    } catch (error) {
        Logger.log(`連線測試失敗: ${error.message}`);
        return false;
    }
}

/**
 * 清理函式 - 刪除測試用的投影片
 * @param {number} maxSlides 保留的最大投影片數量
 * @permission SlidesApp
 */
function cleanupTestSlides(maxSlides = 10) {
    try {
        const presentation = SlidesApp.openById(PRESENTATION_ID);
        const slides = presentation.getSlides();

        if (slides.length > maxSlides) {
            const slidesToDelete = slides.length - maxSlides;

            for (let i = 0; i < slidesToDelete; i++) {
                slides[slides.length - 1 - i].remove();
            }

            Logger.log(`已清理 ${slidesToDelete} 張測試投影片`);
        }

    } catch (error) {
        Logger.log(`清理失敗: ${error.message}`);
    }
}
