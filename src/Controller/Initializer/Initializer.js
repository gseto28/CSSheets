function initializeSpreadSheet() {

    // create log sheet
    if (!verifySheetExist_("log")) {
        createLogSheet_();
    }

    // create command center sheet
    if (!verifySheetExist_("Command Center")) {
        createCommandCenterSheet_();
    }

}

function preflightChecklist() {
    // check if log sheet and command center sheet exist
    if (!verifySheetExist_("log") || !verifySheetExist_("Command Center")) {
        SpreadsheetApp.getUi().alert(
            'Error',
            'Spreadsheet has not been initialized. Please run Initialize Spreadsheet under Utilities.',
            SpreadsheetApp.getUi().ButtonSet.OK
        );
        return false;
    }
    return true;
}