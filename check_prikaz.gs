function check_prikaz() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws_active = ss.getActiveSheet();
  Logger.log (ws_active.getName());
  
  if (!hasAccess(Session.getEffectiveUser().getUsername())){
    SpreadsheetApp.getUi().alert("Нет доступа!");
    return 0;
  }

  if (ws_active.getName() != "Приходы на склад") {
    Logger.log ("Неверная страница")
    return 0;
  }
  let ws_otgr = ss.getSheetByName("Отгружено");
  let ws_kontr = ss.getSheetByName("Контрольный диапазон");
  let ws_gotov = ss.getSheetByName("Готовится к отгрузке");
  let ws_zayvki = ss.getSheetByName("Заявки");
  let ws_prikaz = ss.getSheetByName("Приходы на склад");

  ws_prikaz.getRange(2,2,100,8).clearContent();

  let prikaz_list = [];
  let proj_dict = {};
  proj_dict_plan = {};
  ws_list = [ws_otgr, ws_kontr, ws_gotov, ws_zayvki];
  // ws_list = [ws_gotov, ws_zayvki];
  
  let date_now = new Date();
  let date_diff = ws_prikaz.getRange(1, 14).getValue();
  for (ws of ws_list){
    lastrow = ws.getLastRow();
    const dataRange = ws.getDataRange();
    const values = dataRange.getValues();
    const filteredValues = values.filter((row, index) => {
      date1 = row[24];
      date_ago = 1000;
      try {
        date_ago = calculateDateDifferenceInDays(date1, date_now);
      } catch (error) {
          // date_ago = 1000;
      }
      return date_ago <= date_diff;
    })
    //в отгруженных берем отфильтрованные по дате. В остальных все
    if (ws.getName() != "Заявки") {values_for_search = filteredValues} else {values_for_search = values}
    for (row of values_for_search) {
      tmprec = {}
      if (row[8] == "Сходня") {
        tmprec["date"] = row[24];
        tmprec["sender"] = row[4];
        tmprec["opisanie"] = get_nom_pr(row[17]);
        tmprec["project"] = row[2];
        prikaz_list.push(tmprec);
      }
      if (row[4] == "Сходня") {
        tmpproj = row[2];
        tmpproj = (tmpproj + " ").replace(" ", "");
        tmpproj = (tmpproj).replaceAll(String.fromCharCode(32), "");
        
        tmpproj_list = tmpproj.split("-");
        if (ws.getName() != "Заявки"){
          for (proj of tmpproj_list) {
            if (! (proj in proj_dict)) {
              proj_dict[proj] = 1;
            } else {
              proj_dict[proj] = proj_dict[proj] + 1;
            }
          }
        }else {
          for (proj of tmpproj_list) {
            if (! (proj in proj_dict_plan)) {
              proj_dict_plan[proj] = 1;
            } else {
              proj_dict_plan[proj] = proj_dict_plan[proj] + 1;
            }
          }

        }
      }
    }
  }

  Logger.log (proj_dict_plan);

  for (proj in proj_dict_plan){
    Logger.log (proj + " " + proj.charCodeAt(5));
    // if (proj.length == 6) {
    //   Logger.log (proj)
    // }
  }
  

  prihod_proj_dict = {};
  for (rec of prikaz_list) {
    proj = rec['project'];
    if (proj == '13141') {
      Logger.log(proj_dict_plan['13141 '])
    }
    if (! (proj in prihod_proj_dict)) {
      if (proj in proj_dict){ otgruzh = proj_dict[proj]} else {otgruzh = 0};
      if (proj in proj_dict_plan) {plan = proj_dict_plan[proj]} else {plan = 0}
      prihod_proj_dict[proj] = [proj, Utilities.formatDate(rec['date'], "GMT+3", "dd.MM.yy"), rec['sender'], rec['opisanie'], otgruzh, plan ];
    }
    else {
      prihod_proj_dict[proj][1] = prihod_proj_dict[proj][1] + " / " + Utilities.formatDate(rec['date'], "GMT+3", "dd.MM.yy"); 
      prihod_proj_dict[proj][3] = prihod_proj_dict[proj][3] + " / " + rec['opisanie'];

    }
  }
  let i = 2
  for (proj in prihod_proj_dict) {
    ws_prikaz.getRange(i, 2, 1, 6).setValues([prihod_proj_dict[proj]]);
    i = i + 1;
  }
}
  


function test_date(){
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName("возвраты");
  let activeCell = ws.getRange(4,25);
  let date1 = new Date();
  let date2 = activeCell.getValue(); 
  Logger.log (ws.getName());
  Logger.log(date2);
  Logger.log (calculateDateDifferenceInDays(date1, date2));
  Logger.log (Utilities.formatDate(date1, "GMT+3", "dd.MM.yy") + " " + date1)
}

function calculateDateDifferenceInDays(date1, date2) {
  // Get the time in milliseconds for each date
  let time1 = date1.getTime();
  let time2 = date2.getTime();

  // Calculate the difference in milliseconds
  let diffInMilliseconds = Math.abs(time2 - time1); // Use Math.abs for absolute difference
  // Convert milliseconds to days
  var millisecondsInADay = 1000 * 60 * 60 * 24;
  var diffInDays = Math.floor(diffInMilliseconds / millisecondsInADay);
  return diffInDays
}

function get_nom_pr(full_prikaz) {
  const regex = /[OoОо]{1}-[0-9]{6}/gi;
  let matches = [full_prikaz];
  if (regex.test(full_prikaz)) {
    matches = full_prikaz.match(regex);}
  return matches[0];
}

