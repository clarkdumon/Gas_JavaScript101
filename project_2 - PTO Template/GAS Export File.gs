function getLastRow(gsheetID) {
  let ss = SpreadsheetApp.openById(gsheetID)
  let sh = ss.getSheetByName('Process')
  let range = sh.getRange("A3:A").getValues().filter(data => {
    return data != ""
  }).length + 2
  return range
}

function getdatata() {
  let destinationSheet = SpreadsheetApp.openById(destinationWB).getSheetByName('Conso')
  let rangeA1 = destinationSheet.getRange("A1").getDataRegion().offset(1,0,destinationSheet.getLastRow()-1).getA1Notation()
  destinationSheet.getRange(rangeA1).sort(5).sort(4)
  // destinationSheet.sort(4, true)
}
function fetchDataFromGsheets() {
  let destinationSheet = SpreadsheetApp.openById(destinationWB).getSheetByName('Conso')
  for (gsheet = 0; gsheet < ptoGsheet.length; gsheet++) {
    exportData(ptoGsheet[gsheet])
  }

  let rangeA1 = destinationSheet.getRange("A1").getDataRegion().offset(1,0,destinationSheet.getLastRow()-1).getA1Notation()
  destinationSheet.getRange(rangeA1).sort(5).sort(4)

}
