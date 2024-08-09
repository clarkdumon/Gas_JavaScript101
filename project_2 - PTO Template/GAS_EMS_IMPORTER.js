function importApproved() {
  let dataNeeded = ['Employee EID', 'Date Convert', 'Leave Week', 'Agent Name', 'Team Leader', 'Approval Status', 'Remarks']
  let indexLog = []
  let pasteData = []
  let ptoRange = IMPORTPTO.getRange('A:U').getDisplayValues()
  let ptoHeader = ptoRange.shift()
  for (let scanIndex = 0; scanIndex < ptoHeader.length; scanIndex++) {
    if (!dataNeeded.includes(ptoHeader[scanIndex])) {
      indexLog.push(scanIndex)
    }
  }
  indexLog = indexLog.reverse()
  for (let scanIndex = 0; scanIndex < indexLog.length; scanIndex++) {
    ptoRange.map(res => {
      res.splice(indexLog[scanIndex], 1)
    })
  }

  // console.log(LOCKINGWEEK[0][0])
  ptoRange.map(res => {
    if (res[2] == LOCKINGWEEK[0][0] && res[5] == 'Approved') {
      pasteData.push(res)
    }
  })

  let lastrow = PTOEMS.getSheetByName('EMS').getLastRow() + 1
  if(!pasteData.length <1){
    PTOEMS.getSheetByName('EMS').getRange("A" + lastrow).offset(0, 0, pasteData.length, pasteData[0].length).setValues(pasteData)
  }else{
    SpreadsheetApp.getUi().alert("No Data to be pasted")
  }
  }

function getEmail() {
  
  const EMSTextrange = 'A2:K' + PTOEMS.getSheetByName('EMS').getLastRow()
  const EMSrange = PTOEMS.getSheetByName('EMS').getRange(EMSTextrange).getDisplayValues()
  //EMSrange[0][7] Agent Email
  //EMSrange[0][8] Agent Email

  for(let i = 0; i<EMSrange.length;i++){
    if(EMSrange[i][7]==""){
      PTOEMS.getSheetByName('EMS').getRange(i+2,8).setFormula('xlookup(A'+(i+2)+', IMPORTRANGE("1pAh3oiqV_VJVUsE1hBVwY4i2zIlTtSc4Tv1v3vdmiNo","Active!$D:$D"),IMPORTRANGE("1pAh3oiqV_VJVUsE1hBVwY4i2zIlTtSc4Tv1v3vdmiNo","Active!$Q:$Q"),,0,-1)')
      PTOEMS.getSheetByName('EMS').getRange(i+2,9).setFormula('xlookup(E'+(i+2)+', IMPORTRANGE("1pAh3oiqV_VJVUsE1hBVwY4i2zIlTtSc4Tv1v3vdmiNo","Active!$H:$H"),IMPORTRANGE("1pAh3oiqV_VJVUsE1hBVwY4i2zIlTtSc4Tv1v3vdmiNo","Active!$Q:$Q"),,0,-1)')
    }
    
  }

  const getVals = PTOEMS.getSheetByName('EMS').getRange(EMSTextrange).getDisplayValues()
  PTOEMS.getSheetByName('EMS').getRange(EMSTextrange).setValues(getVals)

}
