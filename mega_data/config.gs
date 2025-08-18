/**
 * Google Sheets 選單配置檔案 - 新增自訂選單按鍵
 * @author GitHub Copilot
 * @permission SpreadsheetApp
 */

/**
 * 當試算表開啟時自動執行，建立自訂選單
 * @permission SpreadsheetApp
 */
function onOpen() {
  try {
    Logger.log('建立自訂選單...');
    
    const ui = SpreadsheetApp.getUi();
    
    // 建立主選單
    ui.createMenu('🔬 MEGA 資料分析')
      .addItem('📊 執行完整分析', 'executeAnalysisFromMenu')
      .addSeparator()
      .addItem('🧪 測試連線', 'testConnectionsFromMenu')
      .addItem('🎯 測試投影片標題', 'testSlideTitleFromMenu')
      .addItem('🎨 測試自訂排版', 'testCustomStyleLayoutFromMenu')
      .addSeparator()
      .addItem('ℹ️ 關於此工具', 'showAboutDialog')
      .addToUi();
      
    Logger.log('✅ 自訂選單建立完成');
    
  } catch (error) {
    Logger.log(`❌ 建立選單失敗: ${error.message}`);
  }
}

/**
 * 從選單執行完整分析 - 主要功能
 * @permission SpreadsheetApp, SlidesApp, DriveApp
 */
function executeAnalysisFromMenu() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 顯示確認對話框
    const response = ui.alert(
      '🔬 MEGA 資料分析',
      '即將開始分析 MEGA 資料並建立投影片。\n\n這個過程可能需要幾分鐘時間，請確認：\n\n' +
      '1. 試算表包含目標圖表\n' +
      '2. Upload_Link 分頁的 B1、B2 儲存格已正確設定\n' +
      '3. 您有足夠的 Google Drive 空間\n\n' +
      '是否要繼續執行？',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      // 顯示執行中提示
      ui.alert(
        '⏳ 執行中',
        '分析正在進行中，請稍候...\n\n' +
        '請勿關閉試算表，執行完成後會有通知。',
        ui.ButtonSet.OK
      );
      
      // 執行主要功能
      const result = executeAnalysis();
      
      // 顯示執行結果
      if (result && result.success) {
        ui.alert(
          '✅ 分析完成！',
          `🎉 MEGA 資料分析已成功完成！\n\n` +
          `📊 處理圖表數量: ${result.chartsProcessed}\n` +
          `📄 建立投影片數量: ${result.slidesCreated}\n\n` +
          `請前往 Google Slides 查看結果：\n` +
          `https://docs.google.com/presentation/d/${PRESENTATION_ID}/edit`,
          ui.ButtonSet.OK
        );
      } else {
        ui.alert(
          '⚠️ 執行完成',
          '分析已執行，但可能遇到一些問題。\n\n請檢查 Apps Script 記錄檔以獲取詳細資訊。',
          ui.ButtonSet.OK
        );
      }
      
    } else {
      Logger.log('使用者取消執行');
    }
    
  } catch (error) {
    Logger.log(`❌ 選單執行失敗: ${error.message}`);
    
    ui.alert(
      '❌ 執行錯誤',
      `執行過程中發生錯誤：\n\n${error.message}\n\n` +
      `請檢查以下項目：\n` +
      `• Upload_Link 分頁的設定是否正確\n` +
      `• 目標分頁是否存在\n` +
      `• 網路連線是否正常\n\n` +
      `如需協助，請檢查 Apps Script 記錄檔。`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 從選單測試連線
 * @permission SpreadsheetApp, SlidesApp
 */
function testConnectionsFromMenu() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.alert(
      '🧪 測試連線',
      '正在測試與 Google Sheets 和 Google Slides 的連線...',
      ui.ButtonSet.OK
    );
    
    const success = testConnections();
    
    if (success) {
      ui.alert(
        '✅ 連線測試成功',
        '所有連線測試通過！\n\n✓ Google Sheets 連線正常\n✓ 目標分頁存在\n✓ Google Slides 連線正常\n\n系統已準備好執行分析。',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        '❌ 連線測試失敗',
        '連線測試遇到問題，請檢查 Apps Script 記錄檔以獲取詳細資訊。',
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    Logger.log(`❌ 連線測試失敗: ${error.message}`);
    ui.alert(
      '❌ 測試錯誤',
      `連線測試失敗：\n\n${error.message}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 從選單測試投影片標題功能
 * @permission SpreadsheetApp
 */
function testSlideTitleFromMenu() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.alert(
      '🎯 測試投影片標題',
      '正在測試從儲存格 X2 讀取投影片標題功能...',
      ui.ButtonSet.OK
    );
    
    const title = testSlideTitle();
    
    ui.alert(
      '✅ 標題測試完成',
      `投影片標題讀取測試成功！\n\n` +
      `📝 讀取的標題: "${title}"\n` +
      `📍 來源儲存格: ${TITLE_CELL}\n\n` +
      `如果標題不正確，請檢查目標分頁的 ${TITLE_CELL} 儲存格內容。`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log(`❌ 標題測試失敗: ${error.message}`);
    ui.alert(
      '❌ 測試錯誤',
      `投影片標題測試失敗：\n\n${error.message}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 從選單測試自訂排版功能
 * @permission None
 */
function testCustomStyleLayoutFromMenu() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.alert(
      '🎨 測試自訂排版',
      '正在測試自訂樣式排版設定...',
      ui.ButtonSet.OK
    );
    
    testCustomStyleLayout();
    
    ui.alert(
      '✅ 排版測試完成',
      `自訂樣式排版測試成功！\n\n` +
      `📏 固定高度: ${(CHART_STYLE_CONFIG.FIXED_HEIGHT / 72).toFixed(2)} 英寸\n` +
      `📐 寬高比: ${CHART_STYLE_CONFIG.ASPECT_RATIO}:1\n` +
      `📍 位置數量: ${CHART_STYLE_CONFIG.POSITIONS.length}\n\n` +
      `詳細測試結果請檢查 Apps Script 記錄檔。`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log(`❌ 排版測試失敗: ${error.message}`);
    ui.alert(
      '❌ 測試錯誤',
      `自訂排版測試失敗：\n\n${error.message}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * 顯示關於對話框
 * @permission SpreadsheetApp
 */
function showAboutDialog() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    'ℹ️ MEGA 資料分析工具',
    `🔬 MEGA 資料分析器 v5.1.0\n\n` +
    `📋 功能說明：\n` +
    `• 自動從 Google Sheets 擷取指定圖表\n` +
    `• 建立高品質的 Google Slides 簡報\n` +
    `• 支援自訂排版和動態標題\n` +
    `• 提供完整的測試和驗證功能\n\n` +
    `🎯 目標圖表：\n` +
    `• MaxPwr - TX Power\n` +
    `• MaxPwr - ACLR - Max(E_-1,E_+1)\n` +
    `• MaxPwr - EVM\n\n` +
    `👨‍💻 開發者: GitHub Copilot\n` +
    `📅 版本日期: 2025-08-13\n` +
    `🔧 相容性: Google Apps Script V8 Runtime`,
    ui.ButtonSet.OK
  );
}

/**
 * 安裝選單（手動呼叫用）
 * @permission SpreadsheetApp
 */
function installMenu() {
  onOpen();
  Logger.log('✅ 選單已手動安裝');
}

/**
 * 移除自訂選單（清理用）
 * @permission SpreadsheetApp
 */
function removeMenu() {
  try {
    // Google Apps Script 沒有直接移除選單的方法
    // 重新載入試算表後選單會自動重置
    Logger.log('ℹ️ 請重新載入試算表以移除自訂選單');
    
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'ℹ️ 移除選單',
      '請重新載入試算表以移除自訂選單。\n\n按 Ctrl+R (Windows) 或 Cmd+R (Mac) 重新載入。',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log(`❌ 移除選單失敗: ${error.message}`);
  }
}
