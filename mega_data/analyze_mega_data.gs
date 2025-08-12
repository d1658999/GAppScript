/**
 * MEGA 資料分析器 - 自動從 Google Sheets 擷取圖表並建立 Slides 簡報
 * @author GitHub Copilot
 * @permission SpreadsheetApp, SlidesApp, DriveApp, Utilities
 */

// 常數定義
const SPREADSHEET_ID = '1nnzLzALFvtwW0Lz8aVGCw7O0iaPXwN3_wRUF3BC2iak';
const PRESENTATION_ID = '1_zV6grTwFms2a0jE2i9Tz0X6mjbnCpZk8VDKIyBLxi8';
const TARGET_SHEET_NAME = '[NR_TX_LMH]Summary&NR_Test_1';
const SLIDE_TITLE = 'MEGA 資料分析 - NR TX LMH 測試結果';

// 目標圖表名稱（只處理這些特定圖表）
const TARGET_CHART_TITLES = [
    'MaxPwr - TX Power',
    'MaxPwr - ACLR - Max(E_-1,E_+1)',
    'MaxPwr - EVM'
];

// 投影片配置
const MAX_CHARTS_PER_SLIDE = 3; // 每個投影片最多放三張圖片

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
        const slides = createSlidesWithCharts(presentation, chartBlobs);

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
 * 將圖表匯出為圖片檔
 * @param {Object[]} targetCharts 目標圖表物件陣列
 * @return {Object[]} 包含圖片 Blob 和標題的物件陣列
 * @permission DriveApp
 */
function exportChartsAsImages(targetCharts) {
    const chartData = [];

    for (let i = 0; i < targetCharts.length; i++) {
        try {
            const chartInfo = targetCharts[i];
            Logger.log(`處理圖表: "${chartInfo.title}"`);

            // 取得圖表的圖片
            const chartBlob = chartInfo.chart.getBlob();
            chartData.push({
                blob: chartBlob,
                title: chartInfo.title,
                originalIndex: chartInfo.index
            });

            Logger.log(`圖表 "${chartInfo.title}" 轉換完成`);

        } catch (error) {
            Logger.log(`圖表 "${targetCharts[i].title}" 處理失敗: ${error.message}`);
        }
    }

    return chartData;
}

/**
 * 建立投影片並插入圖表（支援多張投影片）
 * @param {Presentation} presentation Google Slides 簡報物件
 * @param {Object[]} chartData 包含圖片 Blob 和標題的物件陣列
 * @return {Slide[]} 建立的投影片陣列
 * @permission SlidesApp
 */
function createSlidesWithCharts(presentation, chartData) {
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

        // 添加標題
        addSlideTitle(slide, slideIndex + 1, totalSlides);

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
 * 為投影片添加標題（根據截圖樣式調整）
 * @param {Slide} slide 投影片物件
 * @param {number} slideNumber 投影片編號（可選）
 * @param {number} totalSlides 總投影片數（可選）
 * @permission SlidesApp
 */
function addSlideTitle(slide, slideNumber = 1, totalSlides = 1) {
    let title = SLIDE_TITLE;

    // 如果有多張投影片，在標題中加入頁碼
    if (totalSlides > 1) {
        title = `${SLIDE_TITLE} (${slideNumber}/${totalSlides})`;
    }

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
}/**
 * 計算保持原始比例的排版（固定圖表高度為 1.4 英寸）
 * @param {number} imageCount 圖片數量（最多3張）
 * @return {Object} 排版資訊
 */
function calculateProportionalLayout(imageCount) {
    const slideWidth = 720;
    const slideHeight = 540;
    const titleHeight = 50;  // 標題高度
    const margin = 15;       // 邊距
    const spacing = 8;       // 圖表間距

    // 固定圖表高度為 1.4 英寸 (1.4 * 72 = 100.8 點)
    const fixedChartHeight = 100.8;

    // 可用區域
    const availableWidth = slideWidth - (2 * margin);
    const availableHeight = slideHeight - titleHeight - (2 * margin);

    let layout;

    if (imageCount === 1) {
        // 單張圖片：居中顯示，使用固定高度
        layout = {
            arrangement: 'single',
            charts: [{
                x: margin,
                y: titleHeight + margin,
                width: availableWidth,
                height: fixedChartHeight,
                maintainAspectRatio: true
            }]
        };
    } else if (imageCount === 2) {
        // 兩張圖片：垂直排列，使用固定高度
        layout = {
            arrangement: 'vertical',
            charts: [
                {
                    x: margin,
                    y: titleHeight + margin,
                    width: availableWidth,
                    height: fixedChartHeight,
                    maintainAspectRatio: true
                },
                {
                    x: margin,
                    y: titleHeight + margin + fixedChartHeight + spacing,
                    width: availableWidth,
                    height: fixedChartHeight,
                    maintainAspectRatio: true
                }
            ]
        };
    } else if (imageCount === 3) {
        // 三張圖片：垂直排列，使用固定高度
        layout = {
            arrangement: 'triple_vertical',
            charts: [
                {
                    x: margin,
                    y: titleHeight + margin,
                    width: availableWidth,
                    height: fixedChartHeight,
                    maintainAspectRatio: true
                },
                {
                    x: margin,
                    y: titleHeight + margin + fixedChartHeight + spacing,
                    width: availableWidth,
                    height: fixedChartHeight,
                    maintainAspectRatio: true
                },
                {
                    x: margin,
                    y: titleHeight + margin + (2 * fixedChartHeight) + (2 * spacing),
                    width: availableWidth,
                    height: fixedChartHeight,
                    maintainAspectRatio: true
                }
            ]
        };
    }

    return {
        ...layout,
        slideWidth: slideWidth,
        slideHeight: slideHeight,
        availableWidth: availableWidth,
        availableHeight: availableHeight,
        margin: margin,
        spacing: spacing,
        fixedChartHeight: fixedChartHeight  // 新增固定圖表高度資訊
    };
}

/**
 * 插入圖表並保持原始比例
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

            if (chartLayout.maintainAspectRatio) {
                // 保持原始比例
                const dimensions = calculateAspectRatioFitDimensions(
                    chartLayout.width,
                    chartLayout.height,
                    chartInfo.blob
                );

                // 計算居中位置
                const centeredX = chartLayout.x + (chartLayout.width - dimensions.width) / 2;
                const centeredY = chartLayout.y + (chartLayout.height - dimensions.height) / 2;

                image.setLeft(centeredX);
                image.setTop(centeredY);
                image.setWidth(dimensions.width);
                image.setHeight(dimensions.height);

                Logger.log(`圖表 "${chartInfo.title}" 已插入 (${centeredX}, ${centeredY}, ${dimensions.width}x${dimensions.height})`);
            } else {
                // 不保持比例（拉伸填滿）
                image.setLeft(chartLayout.x);
                image.setTop(chartLayout.y);
                image.setWidth(chartLayout.width);
                image.setHeight(chartLayout.height);

                Logger.log(`圖表 "${chartInfo.title}" 已插入 (拉伸模式)`);
            }

            // 添加圖表標籤 - 已移除圖表標題顯示
            // addChartLabelOptimized(slide, chartInfo.title, chartLayout, i);

        } catch (error) {
            Logger.log(`插入圖表 "${chartInfo.title}" 失敗: ${error.message}`);
        }
    }
}

/**
 * 計算保持寬高比的適配尺寸（根據截圖優化）
 * @param {number} maxWidth 最大寬度
 * @param {number} maxHeight 最大高度
 * @param {Blob} imageBlob 圖片 Blob
 * @return {Object} 包含 width 和 height 的物件
 */
function calculateAspectRatioFitDimensions(maxWidth, maxHeight, imageBlob) {
    try {
        // 根據指示要求，圖表寬高比設定為 5:1
        // 這能確保圖表保持橫向寬版面的比例
        const chartAspectRatio = 5 / 1; // 寬高比 5:1，符合需求規格

        // 根據可用空間計算最適尺寸
        const widthBasedHeight = maxWidth / chartAspectRatio;
        const heightBasedWidth = maxHeight * chartAspectRatio;

        if (widthBasedHeight <= maxHeight) {
            // 以寬度為準，高度自適應
            return {
                width: maxWidth,
                height: widthBasedHeight
            };
        } else {
            // 以高度為準，寬度自適應
            return {
                width: heightBasedWidth,
                height: maxHeight
            };
        }

    } catch (error) {
        Logger.log(`計算寬高比時發生錯誤: ${error.message}`);
        // 發生錯誤時，使用預設比例
        return {
            width: maxWidth,
            height: maxWidth / (5 / 1) // 使用 5:1 比例
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
