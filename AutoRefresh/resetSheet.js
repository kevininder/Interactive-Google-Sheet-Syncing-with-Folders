/* Reset the dropdown menu. */

function reset() {
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.getRange('A3').activate();
    spreadsheet.getCurrentCell().setValue('0');
    spreadsheet.getRange('A4').activate();
    spreadsheet.getRange('B3').activate();
    spreadsheet.getCurrentCell().setValue('0');
    spreadsheet.getRange('B4').activate();
    spreadsheet.getRange('C3').activate();
    spreadsheet.getCurrentCell().setValue('Choose');
    spreadsheet.getRange('D3').activate();
    spreadsheet.getCurrentCell().setValue('Choose');
    spreadsheet.getRange('E3').activate();
    spreadsheet.getCurrentCell().setValue('Choose');
    spreadsheet.getRange('F3').activate();
    spreadsheet.getCurrentCell().setValue('0');
    spreadsheet.getRange('F4').activate();
  };