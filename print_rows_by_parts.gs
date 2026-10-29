function Print_onpravlenie_info_by_parts() {
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

  //Получение значения
  let part_count = 5;
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Введите кол-во городов:',
    '',
    ui.ButtonSet.OK_CANCEL);
  const button = result.getSelectedButton();
  const responseText = result.getResponseText();

  if (button == ui.Button.OK) {
    // User clicked "OK". Use the input value:
    try {
      part_count = parseInt(responseText,10)
    } catch (error){
      let part_count = 5;
    }
  } else {
    return 0;
  }


  
  //формирование строки для выделенных строк
  let tmp_arr = [];
  //Вспомогательный массив с объединением по городу
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
    if (!city_dict.hasOwnProperty(crit)) {
      city_dict[crit] = otpr["all_cargos"];  
      
    } else {
      city_dict[crit] = city_dict[crit].concat(otpr["all_cargos"]);
    }
  }

  //вспомогательный словарь {Город: {cargos:{арт:кол-во}, project:str}} 
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

  
  //  return 0;

// Распечатка результата
  
  let now = new Date();
  let users_info = "Пользователь: " + Session.getEffectiveUser().getUsername();
  users_info =  users_info + ". Вызов: " + Utilities.formatDate(now, Session.getTimeZone(), "dd.MM   HH:mm:ss")
  ws_itog.getRange(2, 3).setValue(users_info);

  if (missedRows.join(" ").length > 3){
    ws_itog.getRange(2,4).setValue("Пропущенные города: ");
    ws_itog.getRange(2,6).setValue(missedRows.join(";"))
  }
  
  let part_dict = {};
  cur_part = 0;
  citis_count_full = 0;
  let cur_r = 4;
  for (let city in city_dict_join_art){
    part_dict[city] = city_dict_join_art[city];
    cur_part ++;
    citis_count_full++;
    if ((cur_part == part_count) || (citis_count_full == Object.keys(city_dict_join_art).length)){
      let tmp_part = get_full_count(part_dict);
      //распечатываем общий результат
      Logger.log(cur_r);
      ws_itog.getRange(cur_r, 3).setValue("Общее кол-во на " + Object.keys(part_dict).length + " города(ов)");
      cur_r++;
      for (let art_part in tmp_part["cargos_full"]){
        ws_itog.getRange(cur_r, 3).setValue(art_part);
        ws_itog.getRange(cur_r, 4).setValue(tmp_part["cargos_full"][art_part]);
        cur_r++;
      }
      ws_itog.getRange(cur_r, 3).setValue("Города");
      cur_r++;
      for (let city_part of tmp_part["citis"]){
        ws_itog.getRange(cur_r, 3).setValue(city_part["city"] + " " + city_part["projects"]);
        cur_r++;
      }

      //распечатка фасовки
      
      for (let city in part_dict){
        cur_r++;
        ws_itog.getRange(cur_r, 3).setValue(city);
        //проверка городов с отправкой местами
        for (const place_citi of send_by_place){
            if (city.toLowerCase().includes(place_citi)){
              ws_itog.getRange(cur_r, 3).setValue(city + "          ---   Местами!!!");
            }
        }        
        ws_itog.getRange(cur_r, 4).setValue("кол-во");
        ws_itog.getRange(cur_r, 2).setValue(part_dict[city]['project']);
        cur_r++;
        for (let art_info in part_dict[city]['cargos']){
          ws_itog.getRange(cur_r, 3).setValue(art_info);
          ws_itog.getRange(cur_r, 4).setValue(part_dict[city]['cargos'][art_info]);
          for (const fragile_art of fragile){
            if (art_info.includes(fragile_art)){
              ws_itog.getRange(cur_r, 2).setValue("Хрупкое");
            }
          }
          
          cur_r++;
        }
      }
      cur_r++;
      cur_part = 0;
      part_dict = {};
    }
    
  }
  let lastrow = ws_itog.getLastRow();
  ws_itog.setRowHeights(4, lastrow, 50); //устанавливаем высоту строк
  const maxRows = ws_itog.getMaxRows();
  const allRowsRange = ws_itog.getRange(2, 1, maxRows-1, 1);
  ws_itog.unhideRow(allRowsRange); //раскрыть скрытые строки    

  
  ws_itog.activate();
}

// вспомогательный словарь {Город: {cargos:{арт:кол-во}, project:str}} 

function get_full_count(city_dict_join_art){
// функция возвращает объект {cargos_full:{art:count}, citis:[{city:str, projects:str}]}
  let cargos_full = {};
  let citis = [];
  for (let city in city_dict_join_art){
    for (let cargo in city_dict_join_art[city]["cargos"]){
      if (! cargos_full.hasOwnProperty(cargo)){
        cargos_full[cargo] = city_dict_join_art[city]["cargos"][cargo];
      }else{
        cargos_full[cargo] = cargos_full[cargo] + city_dict_join_art[city]["cargos"][cargo];
      }
    }
    citis.push({"city": city, "projects": city_dict_join_art[city]["project"]})
  }

  return ({"cargos_full": cargos_full, "citis": citis})
}


function test_get_full_count(){
  let art_dict = {"Красноярск":{"cargos":{"80842875 Выдeлитель не световой в зоне крючков Gillette Labs для МК базовый":1}, "project":"13318"},
  "Магнитогорск":{"cargos":{"80842875 Выдeлитель не световой в зоне крючков Gillette Labs для МК базовый":1}, "project":"13321"},
  "Иваново":{"cargos":{"80842875 Выдeлитель не световой в зоне крючков Gillette Labs для МК базовый":2}, "project":"13345"},
  "Уфа":{"cargos":{"80842876 Выдeлитель световой в зоне крючков Gillette Labs для МК базовый":1}, "project":"13318"}}

  Logger.log(get_full_count(art_dict));
}

