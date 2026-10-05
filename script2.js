function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Table1");
  if(!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  var data = sheet.getRange("A1:O33").getValues();
  return ContentService.createTextOutput(JSON.stringify(data))
         .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Table1");
  if(!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  var data = JSON.parse(e.postData.contents);
  var action = data.action; 
  var roomNo = String(data.roomNo).trim();
  
  var rows = sheet.getRange("A1:O33").getValues();
  var targetRow = -1;
  
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === roomNo) {
      targetRow = i + 1; 
      break;
    }
  }
  
  if (targetRow !== -1) {
    // Agar action 'checkin' ya 'overwrite' hai toh data update/save kar do bina error ke
    if (action === 'checkin' || action === 'overwrite') {
      var checkInTime = new Date();
      sheet.getRange(targetRow, 1).setValue(checkInTime); 
      sheet.getRange(targetRow, 3).setValue(data.date1);   
      sheet.getRange(targetRow, 4).setValue(data.trainNo || "");
      sheet.getRange(targetRow, 5).setValue(data.lpName || "");
      sheet.getRange(targetRow, 6).setValue(data.lpHq || "");
      sheet.getRange(targetRow, 7).setValue(data.alpName || "");
      sheet.getRange(targetRow, 8).setValue(data.alpHq || "");
      sheet.getRange(targetRow, 9).setValue(data.tmName || "");
      sheet.getRange(targetRow, 10).setValue(data.tmHq || "");
      sheet.getRange(targetRow, 11).setValue(data.signOffTime || "");
    } 
    else if (action === 'checkout') {
      var checkOutTime = new Date();
      sheet.getRange(targetRow, 12).setValue(data.date2 || "");     
      sheet.getRange(targetRow, 13).setValue(data.toTime || "");    
      sheet.getRange(targetRow, 14).setValue(data.signOnTime || "");
      sheet.getRange(targetRow, 15).setValue(checkOutTime);         
      
      var completedRowData = sheet.getRange(targetRow, 1, 1, 15).getValues();
      
      var existingRow41 = sheet.getRange(41, 1).getValue();
      if (existingRow41 !== "" && existingRow41 !== null) {
        var lastRow = sheet.getLastRow();
        if (lastRow >= 41) {
          var rangeToMove = sheet.getRange(41, 1, lastRow - 40 + 1, 15);
          rangeToMove.copyTo(sheet.getRange(42, 1));
        }
      }
      
      sheet.getRange(41, 1, 1, 15).setValues(completedRowData);
      
      sheet.getRange(targetRow, 1).setValue(""); 
      sheet.getRange(targetRow, 3, 1, 13).clearContent(); 
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({"status": "success"}))
         .setMimeType(ContentService.MimeType.JSON);
}