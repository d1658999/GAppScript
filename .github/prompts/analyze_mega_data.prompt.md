## Requirement:
### Google sheets link: https://docs.google.com/spreadsheets/d/1nnzLzALFvtwW0Lz8aVGCw7O0iaPXwN3_wRUF3BC2iak/edit?resourcekey=0-HPA3lUt0jt3Z_AP3tHwVxw&gid=1997635806#gid=1997635806
### Google slide link:https://docs.google.com/presentation/d/1_zV6grTwFms2a0jE2i9Tz0X6mjbnCpZk8VDKIyBLxi8/edit?slide=id.p&resourcekey=0-KGmlRoHf9j9Klh6gSZjkbg#slide=id.p
### To get the sheet tab : `[NR_TX_LMH]Summary&NR_Test_1` and get the pictures of chart in this and then go to Google slide link and then create new slides and put all the pictures in the new slide with vertical and a good layout and design without any distortion and remain the original proportions. One slide only put max three pictures in a slide
### the detailed steps:
1. Open the Google Sheets link and navigate to the tab `[NR_TX_LMH]Summary&NR_Test_1`.
2. Take screenshots of all the charts in this tab.
3. Open the Google Slides link and create a new slide.
4. Insert the screenshots into the new slide.
5. Arrange the screenshots in a visually appealing layout.
6. Add any necessary titles or descriptions to the slide.
7. chartAspectRatio is 5:1, so make sure the images are resized accordingly.
8. don't need to the title of the chart
9. don't use getBlob() method to export image because the images which is taken are different from original image
10. it might use below partial method to export to Google Slides:
```
function exportChartAsImageCorrectly() {
  // 取得工作表中的圖表
  const sheet = SpreadsheetApp.getActiveSheet();
  const charts = sheet.getCharts();
  
  if (charts.length > 0) {
    const chart = charts[0]; // 取得第一個圖表
    
    // 建立臨時的Google Slides
    const slides = SlidesApp.create("temp_slides");
    
    // 將圖表插入到Slides中並轉換為圖片
    const imageBlob = slides
      .getSlides()[0]
      .insertSheetsChartAsImage(chart)
      .getAs("image/png");
    
    // 刪除臨時的Slides檔案
    DriveApp.getFileById(slides.getId()).setTrashed(true);
    
    // 儲存圖片到Google Drive
    const fileName = "圖表_" + new Date().getTime() + ".png";
    DriveApp.createFile(imageBlob.setName(fileName));
    
    Logger.log("圖表已成功匯出為: " + fileName);
  }
}

```
11. I have my own style for the pictures positioning and layout in the slides. Please make sure to follow these guidelines:
- the height of size for every picture should be lock apsect ratio and 1.36 inches 
- the first picture in the slide should be at X: 0.21 inch, Y: 1.11 inch
- the second picture in the slide should be at X: 2.94 inches, Y: 2.55 inches
- the third picture in the slide should be at X: 0.21 inches, Y: 3.98 inches
12. the title of slide should be used for the content ofthe cell `X2`
