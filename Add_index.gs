function add_index() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getActiveSheet();
  let ws_param = ss.getSheetByName("Параметры");

  if (!hasAccess(Session.getEffectiveUser().getUsername())){
    SpreadsheetApp.getUi().alert("Нет доступа!");
    return 0;
  }
  
  if ( ws.getName() !== "Готовится к отгрузке"){
    Logger.log(ws.getName())
    return 0
  }

  if (!hasAccess(Session.getEffectiveUser().getUsername())){
    SpreadsheetApp.getUi().alert("Нет доступа!");
    return 0;
  }

  let sel_range = ws.getActiveRange();
  r = sel_range.getColumn();
  if (r != 21) {return 0}
  let columns = InitColumns();
  let tmp_dict = {};
  for (let i = sel_range.getRow(); i<=sel_range.getLastRow();i++){
    if (ws.getRange(i,r).getValue() != ""){
      return 0;
    }
  }

  for (let i = sel_range.getRow(); i<=sel_range.getLastRow();i++){
    Logger.log (i);
    let tmps = ws.getRange(i, columns["gor_otpr_col"]).getValue() + ws.getRange(i, columns["gor_nazn_col"]).getValue();
    tmps = tmps + ws.getRange(i, columns["poluch_col"]).getValue() + ws.getRange(i, columns["perev_col"]).getValue();
    
    if (! (tmps in tmp_dict)) {
      // tmp_dict[tmps] = ws.getRange(i, columns["gor_otpr_col"]).getValue()[0]+ ws.getRange(i, columns["gor_nazn_col"]).getValue()[0] + getNextNakl() ;
      tmp_dict[tmps] = getNextNakl() ;
      ws.getRange(i, columns["nakl_col"]).setValue(tmp_dict[tmps]);
    } else {
      ws.getRange(i, columns["nakl_col"]).setValue(tmp_dict[tmps]);
    }
  } 
}

function getNextNakl() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws_param = ss.getSheetByName("Параметры");

  let i = 2;
  do {i= i+1}   while (ws_param.getRange(i,14).getValue() != "")
  let nextnumber = ws_param.getRange(i-1,14).getValue() +3;
  ws_param.getRange(i,14).setValue( nextnumber);
  // Logger.log(i);
  // Logger.log(nextnumber);
  return nextnumber;
}