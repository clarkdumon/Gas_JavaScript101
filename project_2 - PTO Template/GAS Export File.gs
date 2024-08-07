function getLastRow(gsheetID) {
  let ss = SpreadsheetApp.openById(gsheetID)
  let sh = ss.getSheetByName('Process')
  let range = sh.getRange("A3:A").getValues().filter(data => {
    return data != ""
  }).length + 2
  return range
}
