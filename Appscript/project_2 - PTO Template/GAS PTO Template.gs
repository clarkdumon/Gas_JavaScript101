function onOpenTrig() {
  SpreadsheetApp.getUi().createMenu('WFM Menu')
    .addItem('Lock PTO Release', 'lockPTOProcessingWeek')
    .addToUi();


}

const ssPTOTemp = SpreadsheetApp.openById('1u5FVAYtJ8V--zcmDO2o6-EvdQfvhWaWrYz06KYsqGh4');
const shFormQuery = ssPTOTemp.getSheetByName('Form Query');
const dataRange = shFormQuery.getRange('A1:G')
const dataRangeValues = shFormQuery.getRange('A2:G').getDisplayValues()

function lockPTOProcessingWeek() {
SpreadsheetApp.getActiveSpreadsheet().toast('Automation Starting','Automation Status');
const shLockWeek = ssPTOTemp.getSheetByName('Locking Week').getRange('A1').getDisplayValues();

  for (let i = 0; i < dataRangeValues.length; i++) {

    if(dataRangeValues[i][0]==""){break;}
      if(dataRangeValues[i][4]!=shLockWeek)
      {
        shFormQuery.getRange(i+2,6).setFormula('=XLOOKUP(C'+(i+2)+',MF!C:C,MF!I:I,"-")')
        shFormQuery.getRange(i+2,7).setFormula('=XLOOKUP(B'+(i+2)+',MF!Q:Q,MF!G:G,"-")')
      }else
      {
        shFormQuery.getRange(i+2,6).setValue(dataRangeValues[i][5])
        shFormQuery.getRange(i+2,7).setValue(dataRangeValues[i][6])
      }
  }

SpreadsheetApp.getActiveSpreadsheet().toast('Automation Complete','Automation Status');

}

function onOpenFillBlanks(){
  
    for (let i =0;i<dataRangeValues.length;i++){
          if(dataRangeValues[i][0]==""){break;}
              if(dataRangeValues[i][5]==""){
                  shFormQuery.getRange(i+2,6).setFormula('=XLOOKUP(C'+(i+2)+',MF!C:C,MF!I:I,"-")')
                  shFormQuery.getRange(i+2,7).setFormula('=XLOOKUP(B'+(i+2)+',MF!Q:Q,MF!G:G,"-")')
                  shFormQuery.getRange(i+2,9).setFormula('=IF(A'+(i+2)+'="","",IF(G'+(i+2)+'<>"-","-",COUNTIF($H$2:H'+(i+2)+',H'+(i+2)+')))')
                  
              }
    }

}
