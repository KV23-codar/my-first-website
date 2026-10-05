function doGet(e) {
  var callback = e.parameter.callback;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("B");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  }
  var data = sheet.getDataRange().getValues();

  var jsonString = JSON.stringify(data);

  if (callback) {
    // अगर कॉलबैक है तो JSONP फॉर्मेट में भेजें (जो कभी ब्लॉक नहीं होता)
    return ContentService.createTextOutput(callback + "(" + jsonString + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } else {
    return ContentService.createTextOutput(jsonString)
      .setMimeType(ContentService.MimeType.JSON);
  }
}