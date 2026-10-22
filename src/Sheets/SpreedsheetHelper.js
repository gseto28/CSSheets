function verifySheetExist_(sheetName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

    if (sheet) {
        return true;
    }
    else {
        return false;
    }
}

function createLogSheet_() {
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet()

    sheet = spreadSheet.insertSheet("Log");

    sheet.getRange(1, 1, 1, 5).setValues([[
        "Timestamp",
        "Level",
        "Action",
        "Status",
        "Details"
    ]]);

    sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    sheet.setFrozenRows(1);

    // Set column widths
    sheet.setColumnWidth(1, 150); // Timestamp
    sheet.setColumnWidth(2, 80);  // Level
    sheet.setColumnWidth(3, 150); // Action
    sheet.setColumnWidth(4, 100); // Status
    sheet.setColumnWidth(5, 3000); // Details

    sheet.setTabColor("#000000"); // Black

    // Format Timestamp as date + time
    sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd hh:mm:ss");

    deleteUnusedColumns_(sheet);

    // Hide the sheet
    sheet.hideSheet();

}

function createCommandCenterSheet_() {
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet()

    sheet = spreadSheet.insertSheet("Command Center");

    // Command Center Settings
    sheet.getRange("A1:B6").setValues(
        [
            ["Command Center Settings", ""],
            ["Command Center Sheet Name", "Command Center"],
            ["Command Center Sheet Range", "A1:B6"],
            ["Login User Info Range", "A8:B16"],
            ["Log Sheet Name", "Log"],
            ["Log Level", ""]
        ]
    )

    // Command Center Cell Level formatting
    //A1
    sheet.getRange("A1").setFontWeight("bold")

    //B5
    const logLevelCell = sheet.getRange("B6")
    const logLevelRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(["LOG_TRACE", "LOG_DEBUG", "LOG_INFO", "LOG_WARN", "LOG_ERROR", "LOG_FATAL", "LOG_OFF"], true)
        .setAllowInvalid(false)
        .build();
    logLevelCell.setDataValidation(logLevelRule);
    logLevelCell.setValue("LOG_OFF");

    // Login User Info
    sheet.getRange("A8:A16").setValues(
        [
            ["Login User Information"],
            ["Login Timestamp"],
            ["status_code"],
            ["first_name"],
            ["last_name"],
            ["role_id"],
            ["client_id"],
            ["employee_id"],
            ["phone number"]
        ]
    )

    //A8
    sheet.getRange("A8").setFontWeight("bold")

    // // Set column widths
    sheet.setColumnWidth(1, 250);   // Settings
    sheet.setColumnWidth(2, 150);   // Value

    deleteUnusedSpace_(sheet);

    sheet.setTabColor("#000000"); // Black

    // Hide the sheet
    sheet.hideSheet();
}


function deleteUnusedSpace_(sheet) {
    deleteUnusedColumns_(sheet);
    const lastRow = sheet.getLastRow();
    sheet.deleteRows(lastRow + 1, sheet.getMaxRows() - lastRow);
}

function deleteUnusedColumns_(sheet) {
    const lastColumn = sheet.getLastColumn();
    sheet.deleteColumns(lastColumn + 1, sheet.getMaxColumns() - lastColumn);
}