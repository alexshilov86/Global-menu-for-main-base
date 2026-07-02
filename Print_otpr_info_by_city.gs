function Print_onpravlenie_info_by_city() {
  //Функция выводит на сверку материалы, объединенные по город+адрес+получатель

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getActiveSheet();
  let ws_itog = ss.getSheetByName("Сверка_материала");
  let sel_range = ws.getActiveRange();
  let columns = InitColumns();
  let city_dict = {}; //словарь для выделенных городов назначения
  let city_vol_dict = {} //словарь для объемов
  let city_dict_join_art = {};
  let tmp_art_dict = {};
  let missedRows = [];
  let fragile = ["81735959"];
  let send_by_place = ["лабытнанги", "советский", "белоярский", "инта"];
  //очистка 
  ws_itog.getRange(1,2,3000,5).clearContent();
  //
  
  //формирование строки для выделенных строк
  let tmp_arr = [];
  //Вспомогательный массив с объединением по городу (city_vol_dict нужно для кол-ва паллет)
  for (let i = sel_range.getRow(); i<=sel_range.getLastRow();i++){
    if (ws.getRange(i, columns["json_col"]).getValue() !== "" ){
      if (!ws.isRowHiddenByFilter(i)){
        //добавляем в tmp_arr json, и в него добавляем еще reciver_adr и reciver
        //сначала переводим в json объект
        let tmp_json = JSON.parse("[" + ws.getRange(i, columns["json_col"]).getValue() + "]");
        //добавляем в него reciver_adr и reciver
        for (let otpr_pr of tmp_json){
          otpr_pr["reciver_adr"] = ws.getRange(i, columns["adr_nazn_col"]).getValue();
          otpr_pr["reciver"] = ws.getRange(i, columns["poluch_col"]).getValue();
        }
        tmp_arr.push (...tmp_json);
        let city_otpr_vol = ws.getRange(i, columns["gor_nazn_col"]).getValue() + ", " + ws.getRange(i, columns["adr_nazn_col"]).getValue() + ", " + ws.getRange(i, columns["poluch_col"]).getValue();
        if (!city_vol_dict.hasOwnProperty(city_otpr_vol)) {
          city_vol_dict[city_otpr_vol] = [ws.getRange(i, columns["kol_mest_col"]).getValue()];
        }
        else {
          city_vol_dict[city_otpr_vol].push(ws.getRange(i, columns["kol_mest_col"]).getValue());
        }
      }
    } else {
      missedRows.push(ws.getRange(i, columns["gor_nazn_col"]).getValue());
    }
  } 

  // Logger.log(city_vol_dict);

  tmp_str = "[" + tmp_arr.join(", ") + "]";
  let package_json = tmp_arr;
  for (let otpr of package_json ){
    let crit = otpr["reciver_city"] + ", " + otpr["reciver_adr"] + ", " + otpr["reciver"];
    if (!city_dict.hasOwnProperty(crit )) {
      city_dict[crit] = otpr["all_cargos"];  
      
    } else {
      city_dict[crit] = city_dict[crit].concat(otpr["all_cargos"]);
    }
  }

  //вспомогательный словарь {Город: {арт:кол-во}} 
  for (let city in city_dict){
    tmp_art_dict = {};
    tmp_info_dict = {};
    tmp_projects = [];
    for (let proj of city_dict[city]){
      tmp_projects.push(proj ["project"]);
      for (let art of proj["project_cargos"] ){
        tmp_art_dict[art["art"]] = (tmp_art_dict[art["art"]]||0) + art["count"];
      }
      tmp_info_dict ['project'] = tmp_projects.join(" - ");
      tmp_info_dict ['cargos'] =tmp_art_dict;
    }
    city_dict_join_art[city] = tmp_info_dict;
  }

  //  Logger.log(city_dict_join_art["Элиста"]['cargos']);
  //  return 0;

// Распечатка результата
  let cur_r = 4;
  let now = new Date();
  let users_info = "Пользователь: " + Session.getEffectiveUser().getUsername();
  users_info =  users_info + ". Вызов: " + Utilities.formatDate(now, Session.getTimeZone(), "dd.MM   HH:mm:ss")
  ws_itog.getRange(2, 3).setValue(users_info);

  if (missedRows.join(" ").length > 3){
    ws_itog.getRange(2,4).setValue("Пропущенные города: ");
    ws_itog.getRange(2,6).setValue(missedRows.join(";"))
  }
  
  for (let city in city_dict_join_art){
    ws_itog.getRange(cur_r, 3).setValue(city);
    
    //проверка городов с отправкой местами
    for (const place_citi of send_by_place){
        if (city.toLowerCase().includes(place_citi)){
          ws_itog.getRange(cur_r, 3).setValue(city + "          ---   Местами!!!");
        }
      }
    
  

    // ws_itog.getRange(cur_r, 4).setValue("кол-во");
    if (get_vol(city_vol_dict[city]) != 0){
      ws_itog.getRange(cur_r, 5).setValue(get_vol(city_vol_dict[city]) + " паллет")
    }
    
    ws_itog.getRange(cur_r, 2).setValue(city_dict_join_art[city]['project']);
    cur_r++;
    for (let art_info in city_dict_join_art[city]['cargos']){
      ws_itog.getRange(cur_r, 3).setValue(art_info);
      ws_itog.getRange(cur_r, 4).setValue(city_dict_join_art[city]['cargos'][art_info]);
      for (const fragile_art of fragile){
        if (art_info.includes(fragile_art)){
          ws_itog.getRange(cur_r, 2).setValue("Хрупкое");
        }
      }
      
      cur_r++;
    }
    cur_r++;
  }
  let lastrow = ws_itog.getLastRow();
  ws_itog.setRowHeights(4, lastrow, 50); //устанавливаем высоту строк
  const maxRows = ws_itog.getMaxRows();
  const allRowsRange = ws_itog.getRange(2, 1, maxRows-1, 1);
  ws_itog.unhideRow(allRowsRange); //раскрыть скрытые строки

  ws_itog.activate();

}

function test_uname(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getActiveSheet(); 
  let sel_range = ws.getActiveRange();
  for (let i = sel_range.getRow(); i<=sel_range.getLastRow();i++){
    if (!ws.isRowHiddenByFilter(i)){
      Logger.log(i);
    }
  }

}

function clear_itog_json(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws_itog = ss.getSheetByName("Сверка_материала");
  ws_itog.getRange(3,3,3000,5).clearContent();

}