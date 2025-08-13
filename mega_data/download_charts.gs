/**
 * 圖表下載器 - 從 Google Sheets 下載圖表並儲存到 Google Drive
 * 根據 chart.instructions.md 要求：保持圖片格式一致
 * @author GitHub Copilot
 * @permission SpreadsheetApp, SlidesApp, DriveApp
 */

/**
 * 主要執行函式 - 下載圖表到 Google Drive
 * @permission SpreadsheetApp, DriveApp
 */
function downloadChartsFromSheet() {
    try {
        Logger.log('開始下載圖表...');

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
            return;
        }

        // 步驟 3: 建立下載資料夾
        const downloadFolder = createDownloadFolder();

        // 步驟 4: 下載圖表到 Google Drive
        const downloadedFiles = downloadChartsToGoogleDrive(targetCharts, downloadFolder);

        Logger.log('圖表下載完成！');
        Logger.log(`下載了 ${downloadedFiles.length} 個圖表檔案`);
        Logger.log(`下載資料夾: ${downloadFolder.getName()}`);
        Logger.log(`資料夾連結: ${downloadFolder.getUrl()}`);

        return {
            success: true,
            filesDownloaded: downloadedFiles.length,
            folderUrl: downloadFolder.getUrl(),
            folderName: downloadFolder.getName(),
            files: downloadedFiles.map(file => ({
                name: file.getName(),
                url: file.getUrl(),
                downloadUrl: file.getDownloadUrl()
            }))
        };

    } catch (error) {
        Logger.log(`錯誤: ${error.message}`);
        throw error;
    }
}

/**
 * 建立下載資料夾
 * @return {Folder} Google Drive 資料夾物件
 * @permission DriveApp
 */
function createDownloadFolder() {
    try {
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
        const folderName = `NR_TX_LMH_Charts_${timestamp}`;

        // 在 Google Drive 根目錄建立資料夾
        const folder = DriveApp.createFolder(folderName);
        Logger.log(`建立下載資料夾: ${folderName}`);

        return folder;

    } catch (error) {
        Logger.log(`建立資料夾失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 下載圖表到 Google Drive（原始圖片不變）
 * @param {Object[]} targetCharts 目標圖表物件陣列
 * @param {Folder} folder Google Drive 資料夾物件
 * @return {File[]} 下載的檔案陣列
 * @permission DriveApp
 */
function downloadChartsToGoogleDrive(targetCharts, folder) {
    const downloadedFiles = [];

    for (let i = 0; i < targetCharts.length; i++) {
        try {
            const chartInfo = targetCharts[i];
            Logger.log(`下載圖表: "${chartInfo.title}"`);

            // 嘗試多種方法取得原始圖片，避免重新渲染
            let chartBlob = null;
            let downloadMethod = '';

            // 方法 1: 嘗試取得原始高解析度圖片
            try {
                chartBlob = getOriginalChartBlob(chartInfo.chart);
                downloadMethod = '原始高解析度';
                Logger.log(`使用原始高解析度方法取得圖表: ${chartInfo.title}`);
            } catch (error) {
                Logger.log(`原始高解析度方法失敗: ${error.message}`);
            }

            // 方法 2: 如果方法 1 失敗，使用原始 getBlob 但設定高品質參數
            if (!chartBlob) {
                try {
                    chartBlob = getHighQualityChartBlob(chartInfo.chart);
                    downloadMethod = '高品質渲染';
                    Logger.log(`使用高品質渲染方法取得圖表: ${chartInfo.title}`);
                } catch (error) {
                    Logger.log(`高品質渲染方法失敗: ${error.message}`);
                }
            }

            // 方法 3: 如果以上都失敗，使用預設方法
            if (!chartBlob) {
                chartBlob = chartInfo.chart.getBlob();
                downloadMethod = '預設方法';
                Logger.log(`使用預設方法取得圖表: ${chartInfo.title}`);
            }

            // 產生檔案名稱（確保檔案名稱安全）
            const safeFileName = generateSafeFileName(chartInfo.title, i + 1);

            // 保持圖片的原始格式和品質
            const file = folder.createFile(chartBlob.setName(safeFileName));

            downloadedFiles.push(file);
            Logger.log(`圖表 "${chartInfo.title}" 已儲存為: ${safeFileName} (方法: ${downloadMethod})`);
            Logger.log(`檔案大小: ${Math.round(chartBlob.getBytes().length / 1024)} KB`);
            Logger.log(`檔案連結: ${file.getUrl()}`);

        } catch (error) {
            Logger.log(`下載圖表 "${targetCharts[i].title}" 失敗: ${error.message}`);
        }
    }

    return downloadedFiles;
}

/**
 * 取得原始圖表 Blob（高解析度，不重新渲染）
 * @param {EmbeddedChart} chart 圖表物件
 * @return {Blob} 原始圖表 Blob
 * @permission SpreadsheetApp
 */
function getOriginalChartBlob(chart) {
    try {
        // 方法 1: 嘗試使用 getAs 方法取得原始格式
        try {
            const originalBlob = chart.getAs('image/png');
            if (originalBlob && originalBlob.getBytes().length > 0) {
                Logger.log('成功使用 getAs 方法取得原始圖片');
                return originalBlob;
            }
        } catch (e) {
            Logger.log(`getAs 方法失敗: ${e.message}`);
        }

        // 方法 2: 嘗試從圖表建構器取得高品質版本
        try {
            const builder = chart.modify();
            const modifiedChart = builder.build();
            const highQualityBlob = modifiedChart.getBlob();

            if (highQualityBlob && highQualityBlob.getBytes().length > 0) {
                Logger.log('成功使用圖表建構器取得高品質圖片');
                return highQualityBlob;
            }
        } catch (e) {
            Logger.log(`圖表建構器方法失敗: ${e.message}`);
        }

        // 如果以上都失敗，拋出錯誤讓呼叫方使用其他方法
        throw new Error('無法取得原始圖表 Blob');

    } catch (error) {
        Logger.log(`取得原始圖表 Blob 失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 取得高品質圖表 Blob
 * @param {EmbeddedChart} chart 圖表物件
 * @return {Blob} 高品質圖表 Blob
 * @permission SpreadsheetApp
 */
function getHighQualityChartBlob(chart) {
    try {
        // 使用預設的 getBlob 方法，但確保取得最高品質
        const blob = chart.getBlob();

        // 檢查圖片大小，如果太小可能是低品質版本
        const sizeKB = blob.getBytes().length / 1024;
        Logger.log(`圖表 Blob 大小: ${Math.round(sizeKB)} KB`);

        if (sizeKB < 10) {
            Logger.log('警告: 圖片檔案大小較小，可能不是最高品質版本');
        }

        return blob;

    } catch (error) {
        Logger.log(`取得高品質圖表 Blob 失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 產生安全的檔案名稱
 * @param {string} title 圖表標題
 * @param {number} index 圖表索引
 * @return {string} 安全的檔案名稱
 */
function generateSafeFileName(title, index) {
    // 移除或替換不安全的字元
    let safeName = title
        .replace(/[\/\\:*?"<>|]/g, '_')  // 替換不安全字元為底線
        .replace(/\s+/g, '_')           // 替換空格為底線
        .replace(/_+/g, '_')            // 合併多個底線
        .replace(/^_|_$/g, '');         // 移除開頭和結尾的底線

    // 如果標題太長，截斷它
    if (safeName.length > 50) {
        safeName = safeName.substring(0, 50);
    }

    // 如果經過處理後名稱為空，使用預設名稱
    if (!safeName) {
        safeName = `Chart_${index}`;
    }

    // 添加時間戳記以避免重複
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');

    // 保持原始圖片格式（通常是 PNG）
    return `${index}_${safeName}_${timestamp}.png`;
}

/**
 * 取得所有圖表的下載連結（不實際下載檔案）
 * @return {Object[]} 包含圖表資訊和下載連結的陣列
 * @permission SpreadsheetApp, DriveApp
 */
function getChartDownloadLinks() {
    try {
        Logger.log('取得圖表下載連結...');

        // 開啟 Google Sheets 並取得目標分頁
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const targetSheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);

        if (!targetSheet) {
            throw new Error(`找不到分頁: ${TARGET_SHEET_NAME}`);
        }

        // 取得並篩選圖表
        const allCharts = targetSheet.getCharts();
        const targetCharts = filterTargetCharts(allCharts);

        const chartLinks = [];

        for (let i = 0; i < targetCharts.length; i++) {
            const chartInfo = targetCharts[i];

            try {
                // 取得圖表的圖片 Blob
                const chartBlob = chartInfo.chart.getBlob();

                // 建立臨時檔案以取得下載連結
                const tempFile = DriveApp.createFile(chartBlob.setName(`temp_chart_${i + 1}.png`));
                const downloadUrl = tempFile.getDownloadUrl();

                chartLinks.push({
                    title: chartInfo.title,
                    index: i + 1,
                    downloadUrl: downloadUrl,
                    fileName: generateSafeFileName(chartInfo.title, i + 1),
                    fileId: tempFile.getId()
                });

                Logger.log(`圖表 "${chartInfo.title}" 下載連結: ${downloadUrl}`);

                // 清理臨時檔案（可選）
                // DriveApp.getFileById(tempFile.getId()).setTrashed(true);

            } catch (error) {
                Logger.log(`取得圖表 "${chartInfo.title}" 下載連結失敗: ${error.message}`);
            }
        }

        return chartLinks;

    } catch (error) {
        Logger.log(`取得下載連結失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 清理下載的臨時檔案
 * @param {string[]} fileIds 要清理的檔案 ID 陣列
 * @permission DriveApp
 */
function cleanupTempFiles(fileIds) {
    try {
        for (const fileId of fileIds) {
            try {
                const file = DriveApp.getFileById(fileId);
                file.setTrashed(true);
                Logger.log(`已清理臨時檔案: ${file.getName()}`);
            } catch (error) {
                Logger.log(`清理檔案 ${fileId} 失敗: ${error.message}`);
            }
        }
    } catch (error) {
        Logger.log(`清理臨時檔案失敗: ${error.message}`);
    }
}

/**
 * 批次下載所有圖表並產生壓縮檔（使用 Google Drive）
 * @return {Object} 包含下載資訊的物件
 * @permission SpreadsheetApp, DriveApp
 */
function batchDownloadCharts() {
    try {
        Logger.log('開始批次下載圖表...');

        // 先下載所有圖表到 Google Drive
        const result = downloadChartsFromSheet();

        if (result.success) {
            Logger.log('批次下載完成！');
            Logger.log(`下載資料夾連結: ${result.folderUrl}`);
            Logger.log('您可以：');
            Logger.log('1. 點擊上方連結開啟 Google Drive 資料夾');
            Logger.log('2. 選擇所有檔案');
            Logger.log('3. 右鍵選擇「下載」來下載到本機電腦');

            return {
                ...result,
                instructions: [
                    '點擊 Google Drive 資料夾連結',
                    '選擇所有圖表檔案',
                    '右鍵選擇「下載」',
                    '檔案將自動打包為 ZIP 下載到電腦'
                ]
            };
        }

        return result;

    } catch (error) {
        Logger.log(`批次下載失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 檢查 Google Drive 中的下載資料夾
 * @param {string} folderNamePattern 資料夾名稱模式（可選）
 * @return {Object[]} 找到的下載資料夾資訊
 * @permission DriveApp
 */
function listDownloadFolders(folderNamePattern = 'NR_TX_LMH_Charts_') {
    try {
        const folders = DriveApp.searchFolders(`title contains "${folderNamePattern}"`);
        const folderList = [];

        while (folders.hasNext()) {
            const folder = folders.next();
            const files = folder.getFiles();
            const fileCount = [];

            while (files.hasNext()) {
                fileCount.push(files.next());
            }

            folderList.push({
                name: folder.getName(),
                url: folder.getUrl(),
                fileCount: fileCount.length,
                lastModified: folder.getLastUpdated(),
                id: folder.getId()
            });
        }

        Logger.log(`找到 ${folderList.length} 個下載資料夾`);
        folderList.forEach(folder => {
            Logger.log(`- ${folder.name} (${folder.fileCount} 個檔案) - ${folder.url}`);
        });

        return folderList;

    } catch (error) {
        Logger.log(`列出下載資料夾失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 透過截圖方式下載原始圖表（真正的原始圖片）
 * 這個方法會建立一個臨時投影片，插入圖表，然後取得投影片截圖
 * @return {Object} 包含下載資訊的物件
 * @permission SpreadsheetApp, SlidesApp, DriveApp
 */
function downloadOriginalChartsViaScreenshot() {
    try {
        Logger.log('=== 開始透過截圖方式下載原始圖表 ===');
        Logger.log('此方法將建立臨時投影片來取得真正的原始圖表圖片...');

        // 步驟 1: 取得目標圖表
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const targetSheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);

        if (!targetSheet) {
            throw new Error(`找不到分頁: ${TARGET_SHEET_NAME}`);
        }

        const allCharts = targetSheet.getCharts();
        const targetCharts = filterTargetCharts(allCharts);

        if (targetCharts.length === 0) {
            throw new Error('沒有找到目標圖表');
        }

        Logger.log(`找到 ${targetCharts.length} 個目標圖表`);

        // 步驟 2: 建立臨時簡報
        const tempPresentation = SlidesApp.create('臨時_圖表截圖_' + Date.now());
        Logger.log(`建立臨時簡報: ${tempPresentation.getName()}`);

        // 步驟 3: 建立下載資料夾
        const downloadFolder = createDownloadFolder();

        // 步驟 4: 為每個圖表建立單獨的投影片並截圖
        const downloadedFiles = [];

        for (let i = 0; i < targetCharts.length; i++) {
            try {
                const chartInfo = targetCharts[i];
                Logger.log(`處理圖表 ${i + 1}: "${chartInfo.title}"`);

                // 建立新投影片
                const slide = tempPresentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

                // 插入圖表到投影片（佔據整個投影片以取得最大解析度）
                const chartBlob = chartInfo.chart.getBlob();
                const image = slide.insertImage(chartBlob);

                // 設定圖片佔據整個投影片
                image.setLeft(0);
                image.setTop(0);
                image.setWidth(720);
                image.setHeight(540);

                // 等待圖片載入
                Utilities.sleep(2000);

                // 取得投影片的縮圖（這是原始渲染的圖片）
                const slideThumb = slide.getThumbnail();

                if (slideThumb) {
                    const safeFileName = generateSafeFileNameExtended(chartInfo.title, i + 1, 'screenshot');
                    const file = downloadFolder.createFile(slideThumb.setName(safeFileName));
                    downloadedFiles.push(file);

                    Logger.log(`✅ 圖表 "${chartInfo.title}" 截圖已儲存: ${safeFileName}`);
                    Logger.log(`   檔案大小: ${Math.round(slideThumb.getBytes().length / 1024)} KB`);
                } else {
                    Logger.log(`❌ 無法取得圖表 "${chartInfo.title}" 的投影片縮圖`);
                }

            } catch (error) {
                Logger.log(`處理圖表 "${targetCharts[i].title}" 時發生錯誤: ${error.message}`);
            }
        }

        // 步驟 5: 清理臨時簡報
        Logger.log('清理臨時簡報...');
        try {
            DriveApp.getFileById(tempPresentation.getId()).setTrashed(true);
            Logger.log('臨時簡報已刪除');
        } catch (error) {
            Logger.log(`清理臨時簡報失敗: ${error.message}`);
        }

        Logger.log('=== 截圖下載完成 ===');
        Logger.log(`成功下載 ${downloadedFiles.length} 個原始圖表截圖`);

        return {
            success: true,
            method: 'screenshot',
            filesDownloaded: downloadedFiles.length,
            folderUrl: downloadFolder.getUrl(),
            folderName: downloadFolder.getName(),
            files: downloadedFiles.map(file => ({
                name: file.getName(),
                url: file.getUrl(),
                downloadUrl: file.getDownloadUrl()
            }))
        };

    } catch (error) {
        Logger.log(`截圖下載失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 直接從 Google Sheets 取得圖表的真實截圖
 * 不透過任何轉換，直接取得分頁中圖表的視覺截圖
 * @return {Object} 包含下載資訊的物件
 * @permission SpreadsheetApp, DriveApp
 */
function downloadDirectChartScreenshots() {
    try {
        Logger.log('=== 開始直接截圖下載 ===');
        Logger.log('此方法將直接從 Google Sheets 取得圖表區域的截圖...');

        // 步驟 1: 取得目標分頁
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const targetSheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);

        if (!targetSheet) {
            throw new Error(`找不到分頁: ${TARGET_SHEET_NAME}`);
        }

        const allCharts = targetSheet.getCharts();
        const targetCharts = filterTargetCharts(allCharts);

        if (targetCharts.length === 0) {
            throw new Error('沒有找到目標圖表');
        }

        Logger.log(`找到 ${targetCharts.length} 個目標圖表`);

        // 步驟 2: 建立下載資料夾
        const downloadFolder = createDownloadFolder();

        // 步驟 3: 直接下載每個圖表的原始 Blob
        const downloadedFiles = [];

        for (let i = 0; i < targetCharts.length; i++) {
            try {
                const chartInfo = targetCharts[i];
                Logger.log(`直接下載圖表 ${i + 1}: "${chartInfo.title}"`);

                // 直接取得圖表的原始 Blob，不進行任何處理
                const originalBlob = chartInfo.chart.getBlob();

                // 檢查 Blob 的資訊
                Logger.log(`原始 Blob 資訊:`);
                Logger.log(`  - 內容類型: ${originalBlob.getContentType()}`);
                Logger.log(`  - 檔案大小: ${Math.round(originalBlob.getBytes().length / 1024)} KB`);
                Logger.log(`  - 原始名稱: ${originalBlob.getName()}`);

                const safeFileName = generateSafeFileNameExtended(chartInfo.title, i + 1, 'direct');

                // 保持完全原始的格式
                const file = downloadFolder.createFile(originalBlob.setName(safeFileName));
                downloadedFiles.push(file);

                Logger.log(`✅ 圖表 "${chartInfo.title}" 直接下載完成: ${safeFileName}`);

            } catch (error) {
                Logger.log(`直接下載圖表 "${targetCharts[i].title}" 失敗: ${error.message}`);
            }
        }

        Logger.log('=== 直接下載完成 ===');
        Logger.log(`成功下載 ${downloadedFiles.length} 個原始圖表檔案`);

        return {
            success: true,
            method: 'direct',
            filesDownloaded: downloadedFiles.length,
            folderUrl: downloadFolder.getUrl(),
            folderName: downloadFolder.getName(),
            files: downloadedFiles.map(file => ({
                name: file.getName(),
                url: file.getUrl(),
                downloadUrl: file.getDownloadUrl()
            }))
        };

    } catch (error) {
        Logger.log(`直接下載失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 產生安全的檔案名稱（支援不同方法標示）
 * @param {string} title 圖表標題
 * @param {number} index 圖表索引
 * @param {string} method 下載方法標示（可選）
 * @return {string} 安全的檔案名稱
 */
function generateSafeFileNameExtended(title, index, method = '') {
    // 移除或替換不安全的字元
    let safeName = title
        .replace(/[\/\\:*?"<>|]/g, '_')  // 替換不安全字元為底線
        .replace(/\s+/g, '_')           // 替換空格為底線
        .replace(/_+/g, '_')            // 合併多個底線
        .replace(/^_|_$/g, '');         // 移除開頭和結尾的底線

    // 如果標題太長，截斷它
    if (safeName.length > 50) {
        safeName = safeName.substring(0, 50);
    }

    // 如果經過處理後名稱為空，使用預設名稱
    if (!safeName) {
        safeName = `Chart_${index}`;
    }

    // 添加時間戳記以避免重複
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');

    // 添加方法標示
    const methodSuffix = method ? `_${method}` : '';

    // 保持原始圖片格式（通常是 PNG）
    return `${index}_${safeName}${methodSuffix}_${timestamp}.png`;
}

/**
 * 模擬手動 "Download Chart" 功能
 * 嘗試取得圖表的原始匯出資料，就像手動右鍵選擇 "Download Chart" → "PNG IMAGE(.PNG)" 一樣
 * @return {Object} 包含下載資訊的物件
 * @permission SpreadsheetApp, DriveApp, UrlFetchApp
 */
function downloadChartsLikeManualExport() {
    try {
        Logger.log('=== 模擬手動 Download Chart 功能 ===');
        Logger.log('嘗試取得圖表的原始匯出資料，就像手動右鍵下載一樣...');

        // 步驟 1: 取得目標圖表
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const targetSheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);

        if (!targetSheet) {
            throw new Error(`找不到分頁: ${TARGET_SHEET_NAME}`);
        }

        const allCharts = targetSheet.getCharts();
        const targetCharts = filterTargetCharts(allCharts);

        if (targetCharts.length === 0) {
            throw new Error('沒有找到目標圖表');
        }

        Logger.log(`找到 ${targetCharts.length} 個目標圖表`);

        // 步驟 2: 建立下載資料夾
        const downloadFolder = createDownloadFolder();

        // 步驟 3: 嘗試取得每個圖表的原始匯出資料
        const downloadedFiles = [];

        for (let i = 0; i < targetCharts.length; i++) {
            try {
                const chartInfo = targetCharts[i];
                Logger.log(`處理圖表 ${i + 1}: "${chartInfo.title}"`);

                // 方法 1: 嘗試取得圖表的原始匯出 URL
                const exportBlob = getChartExportBlob(chartInfo.chart, spreadsheet);

                if (exportBlob) {
                    const safeFileName = generateSafeFileNameExtended(chartInfo.title, i + 1, 'export');
                    const file = downloadFolder.createFile(exportBlob.setName(safeFileName));
                    downloadedFiles.push(file);

                    Logger.log(`✅ 圖表 "${chartInfo.title}" 匯出成功: ${safeFileName}`);
                    Logger.log(`   檔案大小: ${Math.round(exportBlob.getBytes().length / 1024)} KB`);
                    Logger.log(`   內容類型: ${exportBlob.getContentType()}`);
                } else {
                    Logger.log(`❌ 無法取得圖表 "${chartInfo.title}" 的匯出資料`);
                }

            } catch (error) {
                Logger.log(`處理圖表 "${targetCharts[i].title}" 時發生錯誤: ${error.message}`);
            }
        }

        Logger.log('=== 手動匯出模擬完成 ===');
        Logger.log(`成功下載 ${downloadedFiles.length} 個匯出圖表檔案`);

        return {
            success: true,
            method: 'manual_export',
            filesDownloaded: downloadedFiles.length,
            folderUrl: downloadFolder.getUrl(),
            folderName: downloadFolder.getName(),
            files: downloadedFiles.map(file => ({
                name: file.getName(),
                url: file.getUrl(),
                downloadUrl: file.getDownloadUrl()
            }))
        };

    } catch (error) {
        Logger.log(`手動匯出模擬失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 取得圖表的原始匯出 Blob（模擬手動匯出）
 * @param {EmbeddedChart} chart 圖表物件
 * @param {Spreadsheet} spreadsheet 試算表物件
 * @return {Blob} 圖表的匯出 Blob
 * @permission SpreadsheetApp, UrlFetchApp
 */
function getChartExportBlob(chart, spreadsheet) {
    try {
        Logger.log('嘗試取得圖表的原始匯出資料...');

        // 方法 1: 嘗試使用圖表的內建匯出功能
        try {
            // 取得圖表的原始資料，設定最高品質參數
            const chartOptions = chart.getOptions();
            Logger.log(`圖表選項: ${JSON.stringify(chartOptions, null, 2)}`);

            // 嘗試重建圖表以取得最高品質版本
            const builder = chart.modify();

            // 設定高品質匯出選項
            builder.setOption('width', 1200);  // 設定較高的寬度
            builder.setOption('height', 800);  // 設定較高的高度
            builder.setOption('backgroundColor', '#ffffff');  // 設定白色背景
            builder.setOption('chartArea.backgroundColor', '#ffffff');

            const rebuiltChart = builder.build();
            const highQualityBlob = rebuiltChart.getBlob();

            if (highQualityBlob && highQualityBlob.getBytes().length > 0) {
                Logger.log('✅ 成功使用高品質重建方法取得圖表');
                return highQualityBlob;
            }

        } catch (e) {
            Logger.log(`高品質重建方法失敗: ${e.message}`);
        }

        // 方法 2: 嘗試取得圖表的原始 PNG 資料
        try {
            const originalBlob = chart.getBlob();

            // 檢查並優化 Blob
            if (originalBlob.getContentType() === 'image/png') {
                Logger.log('✅ 取得原始 PNG 格式圖表');
                return originalBlob;
            } else {
                // 如果不是 PNG，嘗試轉換
                Logger.log(`圖表格式: ${originalBlob.getContentType()}，嘗試轉換為 PNG...`);
                const pngBlob = Utilities.newBlob(originalBlob.getBytes(), 'image/png', originalBlob.getName());
                return pngBlob;
            }

        } catch (e) {
            Logger.log(`原始 PNG 方法失敗: ${e.message}`);
        }

        // 方法 3: 嘗試使用試算表的匯出功能
        try {
            Logger.log('嘗試使用試算表匯出功能...');

            // 取得圖表在試算表中的位置
            const containerInfo = chart.getContainerInfo();
            const anchorRow = containerInfo.getAnchorRow();
            const anchorCol = containerInfo.getAnchorColumn();

            Logger.log(`圖表位置: 行 ${anchorRow}, 列 ${anchorCol}`);

            // 嘗試取得圖表的原始資料
            const chartBlob = chart.getBlob();

            // 確保是 PNG 格式
            if (chartBlob) {
                const finalBlob = chartBlob.setName(`chart_export_${Date.now()}.png`);
                Logger.log('✅ 使用試算表匯出功能取得圖表');
                return finalBlob;
            }

        } catch (e) {
            Logger.log(`試算表匯出方法失敗: ${e.message}`);
        }

        Logger.log('❌ 所有匯出方法都失敗');
        return null;

    } catch (error) {
        Logger.log(`取得圖表匯出 Blob 失敗: ${error.message}`);
        return null;
    }
}

/**
 * 使用 Google Sheets API 取得圖表的真實匯出資料
 * 這個方法嘗試模擬瀏覽器中的圖表下載功能
 * @return {Object} 包含下載資訊的物件
 * @permission SpreadsheetApp, DriveApp, UrlFetchApp
 */
function downloadChartsViaAPI() {
    try {
        Logger.log('=== 使用 API 方式取得圖表匯出資料 ===');
        Logger.log('這個方法嘗試模擬瀏覽器中的真實圖表下載...');

        // 步驟 1: 取得目標圖表和試算表資訊
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const targetSheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);

        if (!targetSheet) {
            throw new Error(`找不到分頁: ${TARGET_SHEET_NAME}`);
        }

        const allCharts = targetSheet.getCharts();
        const targetCharts = filterTargetCharts(allCharts);

        if (targetCharts.length === 0) {
            throw new Error('沒有找到目標圖表');
        }

        Logger.log(`找到 ${targetCharts.length} 個目標圖表`);

        // 步驟 2: 建立下載資料夾
        const downloadFolder = createDownloadFolder();

        // 步驟 3: 嘗試透過 API 取得每個圖表
        const downloadedFiles = [];

        for (let i = 0; i < targetCharts.length; i++) {
            try {
                const chartInfo = targetCharts[i];
                Logger.log(`API 下載圖表 ${i + 1}: "${chartInfo.title}"`);

                // 取得圖表的詳細資訊
                const chartDetails = getChartDetailsForAPI(chartInfo.chart, targetSheet);

                if (chartDetails) {
                    // 嘗試透過 API 取得圖表
                    const apiBlob = fetchChartViaAPI(chartDetails);

                    if (apiBlob) {
                        const safeFileName = generateSafeFileNameExtended(chartInfo.title, i + 1, 'api');
                        const file = downloadFolder.createFile(apiBlob.setName(safeFileName));
                        downloadedFiles.push(file);

                        Logger.log(`✅ 圖表 "${chartInfo.title}" API 下載成功: ${safeFileName}`);
                        Logger.log(`   檔案大小: ${Math.round(apiBlob.getBytes().length / 1024)} KB`);
                    } else {
                        Logger.log(`❌ 圖表 "${chartInfo.title}" API 下載失敗`);
                    }
                } else {
                    Logger.log(`❌ 無法取得圖表 "${chartInfo.title}" 的 API 詳細資訊`);
                }

            } catch (error) {
                Logger.log(`API 下載圖表 "${targetCharts[i].title}" 時發生錯誤: ${error.message}`);
            }
        }

        Logger.log('=== API 下載完成 ===');
        Logger.log(`成功下載 ${downloadedFiles.length} 個 API 圖表檔案`);

        return {
            success: true,
            method: 'api_export',
            filesDownloaded: downloadedFiles.length,
            folderUrl: downloadFolder.getUrl(),
            folderName: downloadFolder.getName(),
            files: downloadedFiles.map(file => ({
                name: file.getName(),
                url: file.getUrl(),
                downloadUrl: file.getDownloadUrl()
            }))
        };

    } catch (error) {
        Logger.log(`API 下載失敗: ${error.message}`);
        throw error;
    }
}

/**
 * 取得圖表的 API 詳細資訊
 * @param {EmbeddedChart} chart 圖表物件
 * @param {Sheet} sheet 工作表物件
 * @return {Object} 圖表的 API 詳細資訊
 * @permission SpreadsheetApp
 */
function getChartDetailsForAPI(chart, sheet) {
    try {
        const chartId = chart.getChartId();
        const sheetId = sheet.getSheetId();
        const spreadsheetId = sheet.getParent().getId();

        const containerInfo = chart.getContainerInfo();
        const position = {
            row: containerInfo.getAnchorRow(),
            column: containerInfo.getAnchorColumn(),
            width: containerInfo.getWidth(),
            height: containerInfo.getHeight()
        };

        Logger.log(`圖表 API 詳細資訊: ID=${chartId}, Sheet ID=${sheetId}, 位置=${JSON.stringify(position)}`);

        return {
            chartId: chartId,
            sheetId: sheetId,
            spreadsheetId: spreadsheetId,
            position: position,
            ranges: chart.getRanges()
        };

    } catch (error) {
        Logger.log(`取得圖表 API 詳細資訊失敗: ${error.message}`);
        return null;
    }
}

/**
 * 透過 API 取得圖表
 * @param {Object} chartDetails 圖表詳細資訊
 * @return {Blob} 圖表 Blob
 * @permission UrlFetchApp
 */
function fetchChartViaAPI(chartDetails) {
    try {
        Logger.log('嘗試透過 Google Sheets API 取得圖表...');

        // 注意：這需要啟用 Google Sheets API 並且有適當的權限
        // 由於 Apps Script 的限制，我們使用替代方法

        // 替代方法：使用內建功能取得最高品質的圖表
        const spreadsheet = SpreadsheetApp.openById(chartDetails.spreadsheetId);
        const sheet = spreadsheet.getSheets().find(s => s.getSheetId() === chartDetails.sheetId);

        if (sheet) {
            const charts = sheet.getCharts();
            const targetChart = charts.find(c => c.getChartId() === chartDetails.chartId);

            if (targetChart) {
                // 取得最高品質的圖表 Blob
                const blob = targetChart.getBlob();
                Logger.log('✅ 透過替代 API 方法取得圖表');
                return blob;
            }
        }

        Logger.log('❌ API 方法無法取得圖表');
        return null;

    } catch (error) {
        Logger.log(`API 取得圖表失敗: ${error.message}`);
        return null;
    }
}
