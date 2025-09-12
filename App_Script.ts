/* Author: SV
First Authored: 9/8/2025
Last Updated: 9/12/2025

Summary: This Script will be used to auto-copy and generate the 4-column inventory slots for the next week to avoid hand-copying everything

How it works:
There is a menu option called, "New Entry", click "Add Next Week Block" to add the next week's 4-block entry.
This Script is programmed to copy the previous week's MBO entry (ideally it shouldnt change throughout the week), unchecks all the checkboxes, and automatically sets up summation and difference formulas

This works in two parts:
1. addNextWeekBlock() will generate the 

All data and formulas are relative to the columns and not absolute data! So instead of specific data to the cell block, the formula calculation is relative to the data in the repsective cells (like A7 vs A!7).

Limitations/Need to Improve on:
- When adding new rows for inventory, the new rows will not autopopulate for the future entries, or included in the summation of product. Will need to go in this Script in order to add it in. 
- The Date column does not automatically merge in with the 3 other rows (manually have to merge)
- Helper functions for these would be helpful 

In the Works:
- Helper Functions and adding menu entries for the above limitations

*/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('New Entry')
    .addItem('Add Next Week Block', 'addNextWeekBlock') //add option for new data entry
    .addToUi();
}

function addNextWeekBlock() { 
  const sheet = SpreadsheetApp.getActive().getSheetByName('MView');  //tab name. DO NOT CHANGE to another tab IF you are not MV. If want to use for your sheet, need to copy this and make a new file (and add your tab name)...though I highly doubt anyone is seeing this :p

  const headerRowDate   = 1;   // date row
  const headerRowLabels = 2;   // the names of each of the columns ("MBO COUNT", "ACTUAL COUNT", "+/-", "Updated?")
  const firstDataRow    = 3;   // spacer 
  const metaCols        = 4;   // Column A-D meta data (Item name, MOQ, OH)
  const blockWidth      = 4;   // the 4 column block needed for each week's entry

  const lastCol = sheet.getLastColumn(); //this pulls the data from the rightmost column 
  const lastRow = sheet.getLastRow();

  // Start column of last complete 4-column block
  // lastBlockStart = the start of last week
  // nextBlockStart = the start of next week
  const lastBlockStart = lastCol - ((lastCol - metaCols) % blockWidth) - (blockWidth - 1);
  const nextBlockStart = lastBlockStart + blockWidth; ///

  // Next date = last date + 7 (fallback: today)
  let lastDate = sheet.getRange(headerRowDate, lastBlockStart).getValue();
  if (!(lastDate instanceof Date)) lastDate = new Date();
  const nextDate = new Date(lastDate.getTime() + 7 * 24 * 60 * 60 * 1000); /// this gets us the next date needed...so 7 days in advance

  // INSERT columns
  sheet.insertColumnsAfter(lastCol, blockWidth);

  // Set and merge date header
  const dateCell = sheet.getRange(headerRowDate, nextBlockStart);
  dateCell.setValue(nextDate);
  
  // Merge the date across all 4 columns
  const dateRange = sheet.getRange(headerRowDate, nextBlockStart, 1, blockWidth);
  dateRange.merge();
  dateRange.setHorizontalAlignment('center');
  dateRange.setFontWeight('bold');
  
  // Copy formatting from previous week's date header
  const prevDateRange = sheet.getRange(headerRowDate, lastBlockStart, 1, blockWidth);
  prevDateRange.copyFormatToRange(sheet, nextBlockStart, nextBlockStart + blockWidth - 1, headerRowDate, headerRowDate);

  // Set column headers
  sheet.getRange(headerRowLabels, nextBlockStart, 1, blockWidth)
    .setValues([['MBO COUNT','ACTUAL COUNT','+/- (gains or losses)','Updated Actual Count in MBO?']]);

  // Copy header formatting from previous week
  const prevHeaderRange = sheet.getRange(headerRowLabels, lastBlockStart, 1, blockWidth);
  prevHeaderRange.copyFormatToRange(sheet, nextBlockStart, nextBlockStart + blockWidth - 1, headerRowLabels, headerRowLabels);

  const nRows = lastRow - (firstDataRow - 1);
  if (nRows <= 0) return;

  // Copy all formatting from previous week's data area
  const prevDataRange = sheet.getRange(firstDataRow, lastBlockStart, nRows, blockWidth);
  prevDataRange.copyFormatToRange(sheet, nextBlockStart, nextBlockStart + blockWidth - 1, firstDataRow, lastRow);

  // MBO = previous week's ACTUAL COUNT (if available), otherwise copy from previous MBO
  const mboRange = sheet.getRange(firstDataRow, nextBlockStart, nRows, 1);
  mboRange.setFormulaR1C1('=IF(RC1="","", IF(RC[-3]<>"", RC[-3], IF(RC[-4]<>"", RC[-4], "")))');

  // +/- = ACTUAL - MBO (only calculate if ACTUAL is not empty)
  sheet.getRange(firstDataRow, nextBlockStart + 2, nRows, 1)
    .setFormulaR1C1('=IF(OR(RC[-1]="", RC[-2]=""), "", RC[-1]-RC[-2])');

  // Add checkboxes only to TOTAL rows (those with yellow background or containing "TOTAL")
  addCheckboxesToTotalRows(sheet, nextBlockStart + 3, firstDataRow, lastRow);

  // Auto-resize columns
  for (let col = nextBlockStart; col < nextBlockStart + blockWidth; col++) {
    sheet.autoResizeColumn(col);
  }

  SpreadsheetApp.getUi().alert(`New week block added for ${Utilities.formatDate(nextDate, Session.getScriptTimeZone(), "M/d/yyyy")}!\n\nReady for Monday inventory count.`);
}

// Helper function to add checkboxes only to TOTAL rows and not to every row
function updateFormulasForNewRows() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('MView');
  const headerRowDate = 1;
  const firstDataRow = 3;
  const metaCols = 4;
  const blockWidth = 4;
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  
  // For each week block, update formulas for all rows
  for (let blockStart = metaCols + 1; blockStart <= lastCol; blockStart += blockWidth) {
    const nRows = lastRow - (firstDataRow - 1);
    
    // Update MBO formulas
    const mboRange = sheet.getRange(firstDataRow, blockStart, nRows, 1);
    mboRange.setFormulaR1C1('=IF(RC1="","", IF(RC[-3]<>"", RC[-3], IF(RC[-4]<>"", RC[-4], "")))'); //this is set to grab the data from the previous week
    
    // Update +/- formulas
    const gainLossRange = sheet.getRange(firstDataRow, blockStart + 2, nRows, 1);
    gainLossRange.setFormulaR1C1('=IF(OR(RC[-1]="", RC[-2]=""), "", RC[-1]-RC[-2])'); //this is set to do the simeple [B] - [A] = +/- number 
    
    // Update checkboxes for TOTAL rows
    addCheckboxesToTotalRows(sheet, blockStart + 3, firstDataRow, lastRow);
  }
}

