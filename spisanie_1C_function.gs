function Print_spisanie_1C_function(source_list, ws) {
  //функция распечатывает списание для выделенных строк
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws_itog = ss.getSheetByName("Списание");
  let columns = InitColumns();
  let project_dict = {}; //словарь для выделенных городов назначения
  let project_dict_join_art = {};
  let project_city_dict = {};
  let missedRows = [];
  //очистка 
  // ws_itog.getRange(1,2,3000,5).clearContent();
  //
  
  //проверка одинаковой даты
  let now = new Date();
  let tmp_date = ws.getRange(source_list[0], columns["date_otgr_col"]).getValue();
  if (tmp_date == ""){
    tmp_date = Utilities.formatDate(now, Session.getTimeZone(), "dd.MM.yyyy")
  }else{
    tmp_date = Utilities.formatDate(tmp_date, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd.MM.yyyy");
  }
  
  for (let i of source_list){
    let s_date = ws.getRange(i, columns["date_otgr_col"]).getValue();
    if (s_date == ""){
      s_date = Utilities.formatDate(now, Session.getTimeZone(), "dd.MM.yyyy")
    } else{
      s_date = Utilities.formatDate(s_date, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd.MM.yyyy");
    }
    
    if (s_date !== tmp_date){
      SpreadsheetApp.getUi().alert("Выбраны разные даты отгрузки!");
      return 0;
    }
  }
  s_date = tmp_date;
  

  //формирование строки для выделенных строк
  let tmp_arr = [];
  
  for (let i of source_list){
    console.log (ws.getRange(i, columns["json_col"]).getValue());
    if (ws.getRange(i, columns["json_col"]).getValue() !== "" ){
      if (!ws.isRowHiddenByFilter(i)){tmp_arr.push (ws.getRange(i, columns["json_col"]).getValue())}
    } else {
      missedRows.push(ws.getRange(i, columns["gor_nazn_col"]).getValue());
    }
  } 
  tmp_str = "[" + tmp_arr.join(", ") + "]";

  let package_json = JSON.parse(tmp_str);
  
  //Вспомогательный массив с объединением по проектам
  for (let otpr of package_json ){
    
    for (let all_cargo of otpr["all_cargos"]){
      if (!project_dict.hasOwnProperty(all_cargo["project"])){
        project_dict[all_cargo["project"]] = all_cargo["project_cargos"];
        project_city_dict[all_cargo["project"]] =[otpr["reciver_city"]] ;
      } else {
        project_dict[all_cargo["project"]] = project_dict[all_cargo["project"]].concat(all_cargo["project_cargos"]);
        project_city_dict[all_cargo["project"]].push(otpr["reciver_city"]);
      }
    }
  }

  //Во вспомогательном массиве объединяем артикулы
  for (let proj in project_dict){
    tmp_art_dict = {}
    for (let art_info of project_dict[proj]){
      tmp_art_dict[art_info["art"]] = (tmp_art_dict[art_info["art"]] || 0)+art_info["count"]
    }
    project_dict_join_art[proj] =tmp_art_dict;
  }

  //Распечатка результата
  let cur_r = ws_itog.getLastRow()+1;
  let all_art_count = 0;
  for (let proj in project_dict_join_art){
    all_art_count = all_art_count + Object.keys(project_dict_join_art[proj]).length
  }
  now = new Date();
  //let users_info = "Пользователь: " + Session.getEffectiveUser().getUsername();
  //users_info =  users_info + ". Вызов: " + Utilities.formatDate(now, Session.getTimeZone(), "dd.MM   HH:mm:ss")
  ws_itog.getRange(cur_r, 1, all_art_count,1).setValue(Session.getEffectiveUser().getUsername());
  ws_itog.getRange(cur_r, 2, all_art_count,1).setValue(Utilities.formatDate(now, Session.getTimeZone(), "dd.MM   HH:mm:ss"));
  ws_itog.getRange(cur_r, 3, all_art_count,1).setValue(s_date);

  for (let proj in project_dict_join_art){
    let art_count_in_project = Object.keys(project_dict_join_art[proj]).length
    ws_itog.getRange(cur_r, 4).setValue(proj);
    let filteredproject_city_dict = project_city_dict[proj].filter(item => item !== undefined);
    ws_itog.getRange(cur_r, 8).setValue(filteredproject_city_dict.join(";"));
    for (let art_info in project_dict_join_art[proj]){
      ws_itog.getRange(cur_r, 6).setValue(art_info);
      ws_itog.getRange(cur_r, 7).setValue(project_dict_join_art[proj][art_info]);
      ws_itog.getRange(cur_r, 5).setValue(get_art(art_info));

      cur_r++;
    }
  }

  return 0;
}

function test_Print_spisanie_1C_function(){
  let source_list = [55,56,57,58,59];
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws_itog = ss.getSheetByName("Готовится к отгрузке");

  let res = Print_spisanie_1C_function(source_list, ws_itog);
  // for (let i of source_list){
  //   console.log (i)
  // }
}