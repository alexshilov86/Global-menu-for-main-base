function html_1c_form_show() {
  const html1c_form = HtmlService.createTemplateFromFile("1c_diolog_form");
  const htmlOutput = html1c_form.evaluate();
  const ui = SpreadsheetApp.getUi()
  ui.showSidebar(htmlOutput)  
}

function returnLoadData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Готовится к отгрузке"); // Replace "Sheet1" with your sheet name
  const range = sheet.getDataRange();
  let values = range.getValues();
  const columns = InitColumns();
  s_city_list = [];
  for (let i = 1; i<values.length; i++){
    s_city = values[i][columns["gor_otpr_col"]-1] //город отправления
    if (!s_city_list.includes(s_city) && s_city !== ""){
      s_city_list.push(s_city)
    } 
  }

  s_date_list = []
  for (let i = 1; i<values.length; i++){
    let s_date = values[i][columns["date_otgr_col"]-1] //дата отправления
    try {
      s_date = Utilities.formatDate(s_date, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd.MM.yyyy");
    }catch (error) {s_date = ""}
    if (!s_date_list.includes(s_date) && s_date !== ""){
      s_date_list.push(s_date)
    } 
  }
  // console.log (s_date_list)
  return [s_city_list, s_date_list];
}

function test_returnLoadData(){
  // let data_s = returnLoadData();
  console.log ("DDDD");
  let arr = ["dd", "dd"];
  console.log (arr.length)
}
function AddSpisanie(s_city, s_date){
  const ws_data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Готовится к отгрузке");
  // const ws_data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Контрольный диапазон");
  
  const ws_itog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Списание");
  const values = ws_data.getDataRange().getValues();
  const columns = InitColumns();
  let project_dict = {}; //словарь для выделенных городов назначения
  let project_dict_join_art = {};
  let project_city_dict = {};
  
  
  //обход по всем строкам
  let tmp_arr = [];
  let missed_rows = [];
  let spb_list = [];
  for (let i = 1; i < values.length; i++) {
    let tmp_city = values[i][columns["gor_otpr_col"]-1].toLowerCase();
    let tmp_date = values[i][columns["date_otgr_col"]-1];
    let tmp_resiver_city = values[i][columns["gor_nazn_col"]-1];
    try {
      tmp_date = Utilities.formatDate(tmp_date, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd.MM.yyyy");
    }catch (error) {tmp_date = ""}
    if (tmp_city == s_city && tmp_date == s_date){
      if (tmp_resiver_city.toLowerCase().includes("петербург")){
        spb_list.push(i+1); //строки с спб собираем в отдельный массив
      } else{ 
        if (values[i][columns["json_col"]-1] !== ""){
          tmp_arr.push(values[i][columns["json_col"]-1])
        }else{
          missed_rows.push(i)
        }
      }
    }
  }

  //печать результата
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
  const now = new Date();
  //let users_info = "Пользователь: " + Session.getEffectiveUser().getUsername();
  //users_info =  users_info + ". Вызов: " + Utilities.formatDate(now, Session.getTimeZone(), "dd.MM   HH:mm:ss")
  ws_itog.getRange(cur_r, 1, all_art_count,1).setValue(Session.getEffectiveUser().getUsername());
  ws_itog.getRange(cur_r, 2, all_art_count,1).setValue(Utilities.formatDate(now, Session.getTimeZone(), "dd.MM   HH:mm:ss"));
  ws_itog.getRange(cur_r, 3, all_art_count,1).setValue(s_date);

  for (let proj in project_dict_join_art){
    let art_count_in_project = Object.keys(project_dict_join_art[proj]).length
    // ws_itog.getRange(cur_r, 4, art_count_in_project,1 ).setValue(proj);
    ws_itog.getRange(cur_r, 4).setValue(proj);
    let filteredproject_city_dict = project_city_dict[proj].filter(item => item !== undefined);
    ws_itog.getRange(cur_r, 8).setValue(filteredproject_city_dict.join(";"));
    for (let art_info in project_dict_join_art[proj]){
      ws_itog.getRange(cur_r, 6).setValue(art_info);
      ws_itog.getRange(cur_r, 7).setValue(project_dict_join_art[proj][art_info]);
      ws_itog.getRange(cur_r, 5).setValue(get_art(art_info));
      
      cur_r++;
    }
    //cur_r++;
  }
  //печать результата для СПб
  if (spb_list.length > 0){let res = Print_spisanie_1C_function(spb_list, ws_data);}
  
  return [tmp_arr, missed_rows];
}

function test_AddSpisanie(){
  let s_city = "сходня";
  let s_date = "10.03.2026";

  res = AddSpisanie(s_city, s_date);
  console.log (res[0].length)
}

function get_art(full_art){
  full_art = full_art.replace(",", "");

  let space_index = full_art.indexOf(" ");
  // console.log (space_index);
  let art = full_art.substring(0, space_index);
  return art;
}

function test_get_art(){
  let full_art = "81685696 Магнитный ключ Gillette";
  art = get_art(full_art);
  console.log (art);
}
