function formatCollection() {

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = spreadsheet.getSheetByName("カリキュラム");

/*
  //3行目(C行)・5列目(E列)から、165行・3列 取得
  var range = sheet.getRange(3, 5, 165, 3);
*/
  var backcolor; 
  
  for(var i = 3; i <= 165; i++) {
    for(var j = 5; j <= 7; j++) {
      var cell = sheet.getRange(i, j);
      var cellColumn = cell.getColumn();
      backcolor = cell.getBackground();

      if(backcolor != "#cccccc") {
        switch (cellColumn) {
        case 5.0:
          cell.setBorder(true, true, true, false, true, true);
          cell.setVerticalAlignment("bottom");
          break
        case 6.0:
          cell.setBorder(true, false, true, false, true, true);
          cell.setVerticalAlignment("top");
          break
        case 7.0:
          cell.setBorder(true, false, true, true, true, true);
          cell.setVerticalAlignment("bottom");
          break
        }
      }

    }
  }
}
