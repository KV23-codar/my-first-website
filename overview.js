function doGet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("OVERVIEW");
  
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({error: "OVERVIEW sheet not found"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // 1. बाईं ओर की टेबल्स का डेटा
  var overviewData = sheet.getRange("A2:D9").getValues();
  var occupancyData = sheet.getRange("A12:D15").getValues();
  var monthlyAvrageMeals = sheet.getRange("A18:D20").getValues();
  var inspectionDuesDate = sheet.getRange("A23:D33").getValues();

  // 2. दाईं ओर की पहली टेबल का डेटा (A36:D59)
  var rightTableData = sheet.getRange("A36:D59").getValues();

  // 3. दाईं ओर की दूसरी टेबल का डेटा (A62:D64)
  var rightTable2Data = sheet.getRange("A62:D64").getValues();

  var combinedData = {
    overview: overviewData,
    occupancy: occupancyData,
    monthly: monthlyAvrageMeals,
    inspection: inspectionDuesDate,
    rightTable: rightTableData,
    rightTable2: rightTable2Data  // यह की (Key) HTML में डेटा भेजने के लिए जरूरी है
  };
  
  return ContentService.createTextOutput(JSON.stringify(combinedData))
    .setMimeType(ContentService.MimeType.JSON);
}
