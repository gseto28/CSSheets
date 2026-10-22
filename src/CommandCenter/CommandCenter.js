// Grab settings from Command Center Sheet 
function setCommandCenterSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Command Center");

  if (!sheet) return;

  const data = sheet.getRange("A2:B6").getValues();
  const properties = PropertiesService.getDocumentProperties();

  data.forEach(([key, value]) => {
    if (key !== "" && key != null) {
      properties.setProperty(String(key), String(value ?? ""));
    }
  });
}