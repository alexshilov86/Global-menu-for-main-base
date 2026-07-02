function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var ui2 = SpreadsheetApp.getUi();

  ui2.createMenu('Global logistic') //общие команды
    // .addItem('Статус накладной', 'loadFormNakladnaya')
    .addItem('Объединить строки', 'merge_rows_zayavka')
    .addItem('Добавить номер отправления', 'add_index')
    .addItem('Проверить приходы на Сходню', 'check_prikaz')
    .addToUi();

  ui.createMenu('Global sklad') //общие команды
    .addItem('Списание 1С общее', 'html_1c_form_show')
    .addItem('Списание 1С для выделеных строк', 'Print_spisanie_1C_for_rows')
    .addToUi();
  
  ui.createMenu('Global Instore') //общие команды
    // .addItem('Статус накладной', 'loadFormNakladnaya')
    .addItem('Список материалов по городам', 'Print_onpravlenie_info_by_city')
    .addItem('Список материалов по проектам', 'Print_onpravlenie_info_by_projects')
    .addItem('Распечатать по частям', 'Print_onpravlenie_info_by_parts')
    .addItem('Сохранить страницу в Эксель', 'openAsExcel')
    
    .addToUi();



} 