// Logging levels
const LOG_TRACE = 10;
const LOG_DEBUG = 20;
const LOG_INFO = 30;
const LOG_WARN = 40;
const LOG_ERROR = 50;
const LOG_FATAL = 60;
const LOG_OFF = 99;

const logLevels = {
  LOG_TRACE: 10,
  LOG_DEBUG: 20,
  LOG_INFO: 30,
  LOG_WARN: 40,
  LOG_ERROR: 50,
  LOG_FATAL: 60,
  LOG_OFF: 99
};

function getLogLevelName_(level) {
  switch (level) {
    case LOG_TRACE: return "TRACE";
    case LOG_DEBUG: return "DEBUG";
    case LOG_INFO: return "INFO";
    case LOG_WARN: return "WARN";
    case LOG_ERROR: return "ERROR";
    case LOG_FATAL: return "FATAL";
    case LOG_OFF: return "OFF";
    default: return "UNKNOWN";
  }
}

function writeLog(action, level, status, details) {

  const setLevel = PropertiesService.getDocumentProperties().getProperty("Log Level");

  if (logLevels[setLevel] > level) {
    return;
  }

  const spreedSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreedSheet.getSheetByName("Log");
  const MAX_DETAILS_LENGTH = 50000;
  const text = details == null ? "" : String(details);

  // Split details into 50,000-character chunks
  const chunks = [];
  if (text.length === 0) {
    chunks.push("");
  } else {
    for (let i = 0; i < text.length; i += MAX_DETAILS_LENGTH) {
      chunks.push(text.substring(i, i + MAX_DETAILS_LENGTH));
    }
  }

  const startRow = sheet.getLastRow() + 1;

  // Build rows
  const rows = chunks.map((chunk, index) => [
    index === 0 ? new Date() : "",
    index === 0 ? getLogLevelName_(level) : "",
    index === 0 ? action : "",
    index === 0 ? status : "",
    chunk
  ]);

  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

  // // Color only the first log level cell
  // if (level === LOG_ERROR) {
  //   sheet.getRange(startRow, 2).setFontColor("red");
  // } else if (level === LOG_WARN) {
  //   sheet.getRange(startRow, 2).setFontColor("yellow");
  // }
}

function clearLog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log");
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
}