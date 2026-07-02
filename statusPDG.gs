function checkPDGStatus(nomNakl) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName("Параметры");
  let wr = ws.getRange(2,11,1000,1);
  let arr = wr.getDisplayValues();
  let strInRange =  arr.some(row => row[0] == nomNakl);
  
  if (strInRange) {
    return "сдан";
  } else return "";
}
