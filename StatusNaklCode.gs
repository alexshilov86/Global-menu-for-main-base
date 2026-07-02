function loadFormNakladnaya(){ // функция создает окно Sidebar для нахождения статуса
  var email = Session.getActiveUser().getUsername().toString();
  if (hasAccess(email)) {
    const htmlSideBar = HtmlService.createTemplateFromFile("StatusNaklStatus");
    const htmlOutput = htmlSideBar.evaluate();
    const ui = SpreadsheetApp.getUi()
    ui.showSidebar(htmlOutput);
  } else {
    const htmlSideBar = HtmlService.createTemplateFromFile("NoAccess");
    const htmlOutput = htmlSideBar.evaluate();
    const ui = SpreadsheetApp.getUi()
    ui.showSidebar(htmlOutput);
  }
  //return htmlSideBar;
}

function hasAccess(email) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName("Параметры");
  let k = 6;
  let AccessList = [];
  while (ws.getRange(k, 5).getValue() != "") {
    AccessList.push(ws.getRange(k, 5).getValue());
    k++;
  }
  return AccessList.includes(email);
}

function InitColumns() {
  ///функция возвращает объект, содержащий номера столбцов
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName("Параметры");
  let columns = {};
  for (let counter = 2; counter <= 39; counter = counter + 1) {
    let name = ws.getRange(counter, 3).getValue();
    let col = ws.getRange(counter, 2).getValue();
    columns[name] = col;
  }
  return columns;
}

function ShowRowNakl(data) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getActiveSheet();
  let row = ws.getActiveCell().getRow();
  let tmpc = InitColumns();
  return [row, tmpc];
}

function getDataFromTable (curRow, columns, direction) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getActiveSheet();
  let data = ws.getRange(curRow,1,1,35).getValues();
  if (direction == "next") {
    while (ws.isRowHiddenByFilter(curRow)) curRow = curRow + 1;
  } else {
      while (ws.isRowHiddenByFilter(curRow)) curRow = curRow - 1;
  }

  ws.getRange(curRow,17,1,5).activate();
  
  
  //return data[0][columns["nakl_col"]].toString();
  return [ws.getRange(curRow, columns["nakl_col"]).getValue(), curRow];
}

function getNextPrevRow(curRow, direction) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getActiveSheet();
  if (direction == "next") {
    let k = 1;
    while (ws.isRowHiddenByFilter(curRow + k)) k++;
  } else {
    let k = -1;
    while (ws.isRowHiddenByFilter(curRow + k)) k--;
  }
  curRow = curRow + k;
  return curRow;
}

function returnCurRowData() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getActiveSheet();
  let i = ws.getActiveRange().getRow();
  let tmpc = InitColumns();
  return [i, tmpc, ws];
}

function addToTableNeDostavlen(curRow, columns, data){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getActiveSheet();
  ws.getRange(curRow, columns["status_col"]).setValue(data);  
}

function getUsernameForAcc() {
  let email = Session.getActiveUser().getUsername().toString();
  Logger.log(email); 
  return email;
}

function addToTableDostavlen(curRow, columns, data){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getActiveSheet();
  let delivDateFact = new Date(data[1]).toLocaleDateString();
  ws.getRange(curRow, columns["status_col"]).setValue(data[0]);
  ws.getRange(curRow, columns["date_dost_col"]).setValue(delivDateFact);
}

function getNaklStatus(nomNakl, columns, curRow) {
  if (nomNakl == "") return ["",""];
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getActiveSheet();
  let TK = ws.getRange(curRow, columns["perev_col"]).getValue();
  if (TK.toLowerCase().includes("пэк")){
    return getNaklStatusPEC(nomNakl);
  } else if (TK.toLowerCase().includes("сдэк")){
    return getNaklStatusCDEK(nomNakl);
  } else return ["",""];
}

function getCDEKAccessToken() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName("Параметры");  
  let MyGrantType = "client_credentials";
  let Account = "Gba5kyWTQlh5PTyf8jlFDGZd6coikpbR";
  let MyPassword = "lMiHIsvMCWVInICvP2r7aQkRCSxfNJRY";
  let myurl = "https://api.cdek.ru/v2/oauth/token?parameters&";
  myurl = myurl + "grant_type=" + MyGrantType;
  myurl = myurl + "&client_id=" + Account;
  myurl = myurl + "&client_secret=" + MyPassword;
  let options = {
     "headers" : {
       "Content-Type": "application/json; charset=utf-8"
     },
    "method" : "POST"   
  };

  let Access_token_cell_time = ws.getRange(2, 7);
  var now = new Date();
  //ws.getRange(6,6).setValue(now.valueOf());
  if (now.valueOf() > Access_token_cell_time.getValue().valueOf() + 3400*1000) {
    const response = UrlFetchApp.fetch(myurl, options);
      let TokenJSON = JSON.parse(response.getContentText());
      let Token = TokenJSON["access_token"];
      ws.getRange(2,6).setValue(Token);
      ws.getRange(2,7).setValue(now);
  }
  return ws.getRange(2,6).getValue();
}

function getNaklStatusCDEK(nomNakl){
  let access_token = getCDEKAccessToken();
  let myurl = "https://api.cdek.ru/v2/orders?cdek_number=" + nomNakl;
  let options = {
     "headers" : {
       "Content-Type": "application/json; charset=utf-8",
       "Authorization": "Bearer " + access_token
     },
    "method" : "GET"   
  };
  const response = UrlFetchApp.fetch(myurl, options);
  let statusJSON = JSON.parse(response.getContentText());
 
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName("Параметры"); 
  let entity = statusJSON["entity"];
  let statuses = entity["statuses"];
  let status = statuses[0]["name"].toLowerCase();
  let delivDateFact = entity["planned_delivery_date"];
  Logger.log(delivDateFact);
  if (status == "вручен") {
    return ["доставлен", delivDateFact]
  } else {
    //return [status, delivDateFact];
    return ["в пути", delivDateFact];
  }
}

function getNaklStatusPEC(nomNakl) {
  nomNakl = nomNakl.replace(/\s/g, '');
  var url = "https://kabinet.pecom.ru/api/v1/cargos/status/";
  let login64 = "R2xvYmFsaW5zdG9yOjNDNDdBREJCRTlCOTEzQ0UwNTJCMUJGRjg0RUUxMjYzQjhGODU5RTY=";
  let tmps = '{ "cargoCodes": ["' + nomNakl +'"]}';
  var options = {
     "async": true,
     "crossDomain": true,
     "method" : "POST",
     'payload': tmps,
     "headers" : {
       "Authorization" : "Basic " + login64,
       "Content-Type": "application/json; charset=utf-8"
    }
  };
  const response = UrlFetchApp.fetch(url, options);
  var statusJSON = JSON.parse(response.getContentText());
  let cargostatus =  statusJSON["cargos"][0]["info"]["cargoStatus"].toLowerCase();
  let delivDatePlan = statusJSON["cargos"][0]["info"]["deliveryPlanDate"];//arrivalPlanDateTime;deliveryPlanDate
  let delivDateFact = statusJSON["cargos"][0]["info"]["receivedByClientDateTime"];
  //delivDatePlan = new Date(delivDatePlan).toLocaleDateString();
  
  if (cargostatus == "доставлен") {
    return [cargostatus, delivDateFact];
  } else {
    return [cargostatus, delivDatePlan];
  }   
}

function updateAllStatusses(first_row, last_row, columns) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getActiveSheet();
  let ws1 = ss.getSheetByName("Параметры");
  first_row = parseInt(first_row);
  last_row = parseInt(last_row);
  let lastRowInTable = ws.getLastRow();
  let k = 0;
  
  if (isNaN(last_row)) last_row = lastRowInTable; //если последняя строка не заполнена, то до конца таблицы
  for (let tmprow = first_row; tmprow <= last_row; tmprow++) {
    let TK = ws.getRange(tmprow, columns["perev_col"]).getValue();
    let curstatus = ws.getRange(tmprow, columns["status_col"]).getValue();
    let nomNakl = ws.getRange(tmprow, columns["nakl_col"]).getValue().toString();
    if ((TK.toLowerCase() == "сдэк" || TK.toLowerCase() == "пэк" ) && curstatus.toLowerCase() != "доставлен" && !ws.isRowHiddenByFilter(tmprow) && nomNakl != "") {
      k++;
      ws.getRange(tmprow, 17,1,5).activate();
      let data = getNaklStatus(nomNakl, columns, tmprow);
      if (data[0] == "доставлен") {
        addToTableDostavlen(tmprow, columns, data);
      } else {
        let delivDatePlan = new Date(data[1]).toLocaleDateString();
        //delivDatePlan = Utilities.formatDate(delivDatePlan, 'GMT', 'dd.MM.yy');
        if (delivDatePlan == "01.01.1970") delivDatePlan = "";
        addToTableNeDostavlen(tmprow, columns, data[0] + " " + delivDatePlan);
      }
    }
  }
  return k;
}


