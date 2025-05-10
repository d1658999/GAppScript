/**
 * Loads PNG images from a specified Google Drive folder and inserts them into a Google Slides presentation.
 * Reads the Drive folder URL from cell A1 and the Slides presentation URL from cell A2 of the active sheet.
 */
function driveImagesToSlides() {
    // Get the active spreadsheet and sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();

    // Get the URLs from cells A1 and A2
    const driveFolderUrl = sheet.getRange('A1').getValue();
    const slidesPresentationUrl = sheet.getRange('A2').getValue();

    if (!driveFolderUrl || !slidesPresentationUrl) {
        Logger.log('Please enter the Google Drive folder URL in cell A1 and the Google Slides presentation URL in cell A2.');
        SpreadsheetApp.getUi().alert('Please enter the Google Drive folder URL in cell A1 and the Google Slides presentation URL in cell A2.');
        return;
    }

    try {
        // Extract IDs from URLs
        const folderId = extractIdFromUrl_(driveFolderUrl);
        const presentationId = extractIdFromUrl_(slidesPresentationUrl);

        if (!folderId) {
            Logger.log('Invalid Google Drive folder URL in A1.');
            SpreadsheetApp.getUi().alert('Invalid Google Drive folder URL in A1.');
            return;
        }
        if (!presentationId) {
            Logger.log('Invalid Google Slides presentation URL in A2.');
            SpreadsheetApp.getUi().alert('Invalid Google Slides presentation URL in A2.');
            return;
        }

        // Get the Drive folder and Slides presentation
        const folder = DriveApp.getFolderById(folderId);
        const presentation = SlidesApp.openById(presentationId);
        const slides = presentation.getSlides();

        // Get PNG files from the folder and sort them by name
        const filesIterator = folder.getFilesByType(MimeType.PNG);
        const fileList = [];
        while (filesIterator.hasNext()) {
            fileList.push(filesIterator.next());
        }
        fileList.sort((a, b) => a.getName().localeCompare(b.getName()));

        let imageCount = 0;
        // Iterate over the sorted file list
        for (const file of fileList) {
            const imageBlob = file.getBlob();

            // Append a new blank slide
            const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

            // Insert the image onto the slide
            const image = slide.insertImage(imageBlob);

            // Get slide dimensions
            const pageWidth = presentation.getPageWidth();
            const pageHeight = presentation.getPageHeight();

            // Get image dimensions
            const imageWidth = image.getWidth();
            const imageHeight = image.getHeight();

            // Calculate center position
            const left = (pageWidth - imageWidth) / 2;
            const top = (pageHeight - imageHeight) / 2;

            // Set image position
            image.setLeft(left);
            image.setTop(top);

            imageCount++;
            Logger.log(`Inserted image: ${file.getName()}`);
        }

        if (imageCount > 0) {
            presentation.saveAndClose(); // Save changes
            Logger.log(`Successfully inserted ${imageCount} PNG images into the presentation.`);
            SpreadsheetApp.getUi().alert(`Successfully inserted ${imageCount} PNG images into the presentation.`);
        } else {
            Logger.log('No PNG images found in the specified folder.');
            SpreadsheetApp.getUi().alert('No PNG images found in the specified folder.');
        }

    } catch (e) {
        Logger.log(`Error: ${e.toString()}\nStack: ${e.stack}`);
        SpreadsheetApp.getUi().alert(`An error occurred: ${e.message}`);
    }
}

/**
 * Helper function to extract the ID from Google Drive/Slides URLs.
 * Handles common URL formats.
 * @param {string} url The URL to extract the ID from.
 * @return {string|null} The extracted ID or null if not found.
 * @private
 */
function extractIdFromUrl_(url) {
    let id = null;
    // Try matching common Drive/Slides URL patterns
    const patterns = [
        /\/d\/([a-zA-Z0-9_-]+)\//, // /d/ID/edit
        /\/file\/d\/([a-zA-Z0-9_-]+)\//, // /file/d/ID/
        /id=([a-zA-Z0-9_-]+)/, // ?id=ID
        /\/folders\/([a-zA-Z0-9_-]+)/, // /folders/ID
        /\/presentation\/d\/([a-zA-Z0-9_-]+)\// // /presentation/d/ID/
    ];

    for (let i = 0; i < patterns.length; i++) {
        const match = url.match(patterns[i]);
        if (match && match[1]) {
            id = match[1];
            break;
        }
    }

    // If no match from patterns, assume the last part of the path might be the ID (less reliable)
    if (!id && url.includes('/')) {
        const parts = url.split('/');
        const potentialId = parts[parts.length - 1].split('?')[0]; // Get last part before query params
        // Basic check for typical ID format (alphanumeric, -, _)
        if (/^[a-zA-Z0-9_-]+$/.test(potentialId) && potentialId.length > 20) { // Add length check for robustness
            id = potentialId;
        }
    }


    return id;
}