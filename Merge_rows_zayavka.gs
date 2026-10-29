function merge_rows_zayavka() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getActiveSheet();
  let sel_range = ws.getActiveRange();
  let columns = InitColumns();
  let city_dict = {};
  let tmpcity = "";

  if (!hasAccess(Session.getEffectiveUser().getUsername())){
    SpreadsheetApp.getUi().alert("Нет доступа!");
    return 0;
  }
  
  if ( ws.getName() !== "Заявки" && ws.getName() !== "Готовится к отгрузке" && ws.getName() !== "возвраты" ){
    Logger.log(ws.getName())
    return 0
  }

  for (let i = sel_range.getRow(); i<=sel_range.getLastRow();i++){
    tmpcity = ws.getRange(i, columns["gor_nazn_col"]).getValue().trim();
    tmpcity = tmpcity + ws.getRange(i, columns["gor_otpr_col"]).getValue().trim();
    tmpcity = tmpcity + ws.getRange(i, columns["poluch_col"]).getValue().trim();

    if (!city_dict.hasOwnProperty(tmpcity)){
      city_dict[tmpcity]=[i]
    } else{
      city_dict[tmpcity].push(i);
      ws.getRange(i, columns["gor_nazn_col"]).setBackground("#EA3B52");
    }
  }

  for (let city in city_dict){
    let seti = [];
    let proj = [];
    let opis = [];
    let json_info = [];
    let vol = [];
    for (i of city_dict[city]){
      seti.push(ws.getRange(i, columns["set_col"]).getValue());
      proj.push(ws.getRange(i, columns["pr_col"]).getValue());
      opis.push(ws.getRange(i, columns["opis_gr_col"]).getValue());
      json_info.push(ws.getRange(i, columns["json_col"]).getValue());
      vol.push(ws.getRange(i, columns["vol_col"]).getValue());
    }
    ws.getRange(city_dict[city][0],columns["set_col"]).setValue([...new Set(seti)].join(" - "));
    ws.getRange(city_dict[city][0],columns["pr_col"]).setValue([...new Set(proj)].join(" - "));
    ws.getRange(city_dict[city][0],columns["opis_gr_col"]).setValue([...new Set(opis)].join(" - "));
    ws.getRange(city_dict[city][0],columns["json_col"]).setValue(json_info.join(", "));
    ws.getRange(city_dict[city][0],columns["vol_col"]).setValue([...new Set(vol)].join(" - "));

  }

  // Logger.log(city_dict)
}
