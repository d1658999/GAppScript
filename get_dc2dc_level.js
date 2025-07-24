
/**
 * get_dc2dc_level.js
 *
 * Google Apps Script to:
 * 1. Read the Google Drive link to the ini file from cell A1 of the first tab.
 * 2. Download and parse the ini file (DPD_Smart_Char_LTE_NR_Vpa_Search.ini).
 * 3. Extract for every band section the 12 specified dc2dc level keys.
 * 4. Output the result as a table to a new sheet named 'DC2DC Level' (one row per key per band).
 */

function getDc2dcLevel() {
    const DC2DC_KEYS = [
        'Pa control table0 dc2dc level Num0',
        'Pa control table1 dc2dc level Num0',
        'Pa control table0 dc2dc level Num1',
        'Pa control table1 dc2dc level Num1',
        'Pa control table0 dc2dc level Num2',
        'Pa control table1 dc2dc level Num2',
        'Pa control table0 dc2dc level Num3',
        'Pa control table1 dc2dc level Num3',
        'Pa control table0 dc2dc level Num4',
        'Pa control table1 dc2dc level Num4',
        'Pa control table0 dc2dc level Num5',
        'Pa control table1 dc2dc level Num5',
    ];

    // 1. Get the ini file link from cell A1 of the first sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    var iniFileUrl = sheet.getRange('A1').getValue();
    if (!iniFileUrl) {
        SpreadsheetApp.getUi().alert('Please provide the Google Drive link to the ini file in cell A1.');
        return;
    }

    // 2. Extract file ID from the URL
    var fileId = extractDriveFileId_(iniFileUrl);
    if (!fileId) {
        SpreadsheetApp.getUi().alert('Invalid Google Drive file link in cell A1.');
        return;
    }

    // 3. Download the ini file content
    var file = DriveApp.getFileById(fileId);
    var iniContent = file.getBlob().getDataAsString();

    // 4. Parse the ini file
    var bandSections = parseIniSections_(iniContent);

    // 5. Prepare output rows
    var output = [['Band Name', 'DC2DC Level Table', 'DC2DC Level Value']];
    for (var band in bandSections) {
        var section = bandSections[band];
        for (var i = 0; i < DC2DC_KEYS.length; i++) {
            var key = DC2DC_KEYS[i];
            var value = section[key] || '';
            output.push([band, key, value]);
        }
    }

    // 6. Write to 'DC2DC Level' sheet
    var outSheet = ss.getSheetByName('DC2DC Level');
    if (!outSheet) outSheet = ss.insertSheet('DC2DC Level');
    outSheet.clearContents();
    outSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}

// Helper: Extract file ID from Google Drive URL
function extractDriveFileId_(url) {
    var match = url.match(/[-\w]{25,}/);
    return match ? match[0] : null;
}

// Helper: Parse ini file into sections with key-value pairs
function parseIniSections_(content) {
    var lines = content.split(/\r?\n/);
    var sections = {};
    var currentSection = null;
    for (var i = 0; i < lines.length; i++) {
        var line = lines[i].trim();
        if (!line || line.startsWith(';')) continue;
        var sectionMatch = line.match(/^\[(.+)\]$/);
        if (sectionMatch) {
            currentSection = sectionMatch[1];
            sections[currentSection] = {};
        } else if (currentSection) {
            var kv = line.split('=');
            if (kv.length >= 2) {
                var key = kv[0].trim();
                var value = kv.slice(1).join('=').trim();
                sections[currentSection][key] = value;
            }
        }
    }
    return sections;
}
