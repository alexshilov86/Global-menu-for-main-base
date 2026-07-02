function Print_onpravlenie_info_by_projects() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getActiveSheet();
  let ws_itog = ss.getSheetByName("Сверка_материала");
  let sel_range = ws.getActiveRange();
  let columns = InitColumns();
  let project_dict = {}; //словарь для выделенных городов назначения
  let project_dict_join_art = {};
  let project_city_dict = {};
  let missedRows = [];
  let fragile = ["81735959"];
  //очистка 
  ws_itog.getRange(1,2,3000,5).clearContent();
  //
  
  
  //формирование строки для выделенных строк
  let tmp_arr = [];
  
  for (let i = sel_range.getRow(); i<=sel_range.getLastRow();i++){
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
  let cur_r = 4;
  const now = new Date();
  let users_info = "Пользователь: " + Session.getEffectiveUser().getUsername();
  users_info =  users_info + ". Вызов: " + Utilities.formatDate(now, Session.getTimeZone(), "dd.MM   HH:mm:ss")
  ws_itog.getRange(2, 3).setValue(users_info);

  if (missedRows.join(" ").length > 3){
    ws_itog.getRange(2,4).setValue("Пропущенные города: ");
    ws_itog.getRange(2,6).setValue(missedRows.join(";"))
  }
  
  for (let proj in project_dict_join_art){
    ws_itog.getRange(cur_r, 3).setValue(proj);
    ws_itog.getRange(cur_r, 4).setValue("кол-во");
    let filteredproject_city_dict = project_city_dict[proj].filter(item => item !== undefined);
    ws_itog.getRange(cur_r, 5).setValue(filteredproject_city_dict.join(";"));
    
    cur_r++;
    for (let art_info in project_dict_join_art[proj]){
      ws_itog.getRange(cur_r, 3).setValue(art_info);
      ws_itog.getRange(cur_r, 4).setValue(project_dict_join_art[proj][art_info]);
      for (const fragile_art of fragile){
        if (art_info.includes(fragile_art)){
          ws_itog.getRange(cur_r, 5).setValue("Хрупкое");
        }
      }      
      cur_r++;
    }
    cur_r++;
  }

  // Logger.log(JSON.stringify(project_dict_join_art))
  ws_itog.activate();
}