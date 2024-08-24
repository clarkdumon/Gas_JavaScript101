function onOpenTrig() {
  SpreadsheetApp.getUi().createMenu('WFM Menu')
    .addItem('Import PTO', 'fetchDataFromGsheets')
    .addToUi();
}
function getLastRow(gsheetID) {
  let ss = SpreadsheetApp.openById(gsheetID)
  let sh = ss.getSheetByName('Process')
  let range = sh.getRange("A3:A").getValues().filter(data => {
    return data != ""
  }).length + 2
  return range
}
function destinationLastRow() {
  let destinationSheet = SpreadsheetApp.openById(WB_DESTINATION).getSheetByName('Conso')
  let range = destinationSheet.getRange("A:A").getValues().filter(data => {
    return data != ""
  }).length + 1
  return range
}


function fetchDataFromGsheets() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Starting Export')
  let destinationSheet = SpreadsheetApp.openById(WB_DESTINATION).getSheetByName('Conso')
  let lockingWeek = SpreadsheetApp.openById(WB_DESTINATION).getSheetByName('LockingWeek').getRange("A1").getDisplayValue()
  deleteafterlockweek(lockingWeek)    // call FN_deleteafterlockingWeek
  for (gsheet = 0; gsheet < ptoGsheet.length; gsheet++) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Exporting PTOs')
    exportData(ptoGsheet[gsheet])
  }

  let rangeA1 = destinationSheet.getRange("A1").getDataRegion().offset(1, 0, destinationSheet.getLastRow() - 1).getA1Notation()
  destinationSheet.getRange(rangeA1).sort(5).sort(3).sort(4)
  SpreadsheetApp.getActiveSpreadsheet().toast('Export Complete')
}

function exportData(gsheetID) { //Extracts data from PTO Template using GsheetID
  let dataRange = "A4:U" + getLastRow(gsheetID)
  let ss = SpreadsheetApp.openById(gsheetID)
  let sh = ss.getSheetByName('Process')
  let exportRange = sh.getRange(dataRange).getValues()
  let destinationSheet = SpreadsheetApp.openById(WB_DESTINATION).getSheetByName('Conso')

  console.log(exportRange.length)
  destinationSheet.getRange("A" + destinationLastRow()).offset(0, 0, exportRange.length, exportRange[0].length).setValues(exportRange)      //Paste PTO in the Destination Sheet

}

// Delete Preadded data Greater than and equal to Locking Week
//  

function deleteafterlockweek(lockingWeek) {
  let count = 0 // Check what row is the first instance of the locking week found
  let destinationSheet = SpreadsheetApp.openById(WB_DESTINATION).getSheetByName('Conso')
  let rangeWeek = destinationSheet.getRange("D:D").getDisplayValues()

  if (destinationSheet.getMaxRows() != 1) {  // Checks if Max row in the Destination Sheet is not only the header
    for (count; count < rangeWeek.length; count++) {
      if (rangeWeek[count] == lockingWeek) {
        break // breaks loop once first instance of locking week is found
      }
    }
    console.log(count)
    if (count == destinationSheet.getMaxRows()) { //Checks if no data was found to remove all added data in the sheet except the header
      count = 1
    }
    destinationSheet.deleteRows(count + 1, destinationSheet.getMaxRows() - count)
  }

}
