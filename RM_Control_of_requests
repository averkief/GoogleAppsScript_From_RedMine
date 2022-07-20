//-- На будущее
//  1. Если в поле description нет оборудования оборудования, то это ломает таблицу.
//    Номенклатура:  далее пусто нет {{collapse(  ...... ) }}
//    1.1. Нужно понять как обработать как обработать таблицу, чтоб сделать сепарацию по пробелу и по табуляции
//  2. Листы не допускаются изменений "Общее" и "Инженеры"
//    2.1. На листе "Общее" идет алгоритм учета последний строки со значением, если поставить любой символ в любую колонку, то он начнет отсчитывать строку последней ячейки с символом
//    2.2. На листе "Инженеры" приведен список инженеров который переходит в выпадающее меню на лист с "Текущая дата". При добавлении/убирании инженеров нужно пересчитывать ячейки и заменять getRange на новое значение. строка 361
//  3. В письмо можно вложить еще таблицу в формате pdf но прикладывается все листы с портретным разделением, что не по фен-шую. Да и особо не нужно.
//--

const keyAccess = SpreadsheetApp.openById("=====ссылка_на_таблицу_с_ключем_доступа=====").getSheetByName('Key').getRange("A1:A1").getCell(1, 1).getValue();
const globalTable = SpreadsheetApp.openById("=====ссылка_на_рабочую_таблицу====="); //Выбор таблицы по ID

//--
//  --Функуия создания меню скриптов для запуска--
//--
function onOpen () {
  let ui = SpreadsheetApp.getUi();
   ui.createMenu("Скрипт")
  .addItem("1. Получить отчет из RM", "ReadingTicketsXML")
  .addItem("2. Обработка листа текущей даты", "FilterData")
  .addItem("3. Отправка на почту", "EmailSend")
  .addToUi();  
}

//
//=================РАЗДЕЛ ФУНКЦИЙ ОТЧЕТА==1. Получить отчет из RM===============================================================================================================
//

//--
//  --Функция получения текущей даты
//--
function GetCurrentDay () {

  let currentDay = new Date (); 
  let dd = String(currentDay .getDate()).padStart(2, '0');
  let mm = String(currentDay .getMonth() + 1).padStart(2, '0'); //January is 0!
  let yyyy = currentDay .getFullYear();
  currentDay = dd + '.' + mm + '.' + yyyy;
  //Logger.log (currentDay);
  return currentDay;

}


//
//=================РАЗДЕЛ ФОРМИРОВАНИЯ ОТЧЕТА==1. Получить отчет из RM===============================================================================================================
//

//--
//  --Запрос с авторизацией, возвращет страницу--
//--
function APIRequest (reqUrl) {
   
  //let key = ""; //Ключ авторизации в redmine
  //let url = encodeURI(reqUrl);
  let payload = { 'Authorization' : 'Basic ' + keyAccess};
  
  let opt = {
    'method'  : 'GET',
    'headers' : payload,
    'muteHttpExceptions' : true
  };
  
  let response = UrlFetchApp.fetch(reqUrl, opt);
  return response.getContentText();

}


//--
//  --Функция разделения поля subject и возврачает номер обращения
//--
function GetSubjectNumber (subNum) {
  subSplitNum = subNum.split(" ")[0];
  subNumEnd = subSplitNum.replace('№', ''); //Убирает символ №
  //subSplitNum.substr(1);   //Возможно убирает первый символ
  return subNumEnd
}

//--
//  --Функция разделения поля subject и возврачает имя контрагента
//--
function GetSubjectNameKA (subKA) {
  //subsplitka = subka.split(" ");
  subSplitaKA = subKA.substring(11) //Костыль который убирает 11 символов в начале, и остаются только данные по контрагенту
  //Logger.log (subSplitaKA);
  return subSplitaKA
}

//--
//  --Функция редактирования даты
//--
function GetCurrentDataTime (dataTime) {
  dataTimeCurrent = dataTime.replace('Z', '').replace('T', ' / ');
  return dataTimeCurrent
}

//--
//  --Функция парсинга description и возвращения "Всего товаров"
//--
function GetDescTotalGoods (textDesc) {
  textDescSplit = textDesc.split(" "); // разделение слов на индексы через пробел
  findIndexCollapse = textDescSplit.indexOf('collapse'); // не помню зачем его искать
  findIndexText = textDescSplit.indexOf('товаров:');
  totalGoogs = textDescSplit[findIndexText + 1];

  return totalGoogs
  

}

//--
//  --Функция парсинга description и возвращения Первого артикула
//--
function GetDescFirstArticle (textDesc) {
  textDescSplit = textDesc.split("|");
  //Logger.log (textDescSplit);
  findIndexText = textDescSplit.indexOf('Артикул');
  //Logger.log (findIndexText + 1);
  firstArticle = textDescSplit[findIndexText + 1];
  //Logger.log (totalGoogs);

  return firstArticle
}

//--
//  --Функция парсинга description и возвращения Гарнтии первого артикула
//--
function GetDescWarrantyFirst (textDesc) {
  textDescSplit = textDesc.split("|");
  //Logger.log (textDescSplit);
  findIndexText = textDescSplit.indexOf('Гарантия');
  //Logger.log (findIndexText + 1);
  warrantyFirst = textDescSplit[findIndexText + 1];
  //Logger.log (totalGoogs);

  return warrantyFirst
}


//--
//  --Парсинг тикетов с XML
//--
function ReadingTicketsXML (){

  let dateNameNewSheetRM = globalTable.getSheetByName(GetCurrentDay());  // Создаем новый лист в текущей таблице
  if (dateNameNewSheetRM != null) {                             // провека условия на отсутсвие имени в названии листа
      globalTable.deleteSheet(dateNameNewSheetRM);
  }
  dateNameNewSheetRM = globalTable.insertSheet(); 
  dateNameNewSheetRM.setName(GetCurrentDay());    // вставить дату в имени станицы, не понятно как это взаимодейстует с 56 строкой

  let rangeRM = dateNameNewSheetRM.getRange("A1:ZZ1000"); //Определение используемых ячеек

  //Добавить назывние стобвцов
  rangeRM.getCell(1,1).setValue("Номер RM"); 
  rangeRM.getCell(1,2).setValue("Тип"); 
  rangeRM.getCell(1,3).setValue("Номер обращения"); 
  rangeRM.getCell(1,4).setValue("Контрагент"); 
  rangeRM.getCell(1,5).setValue("Текущий статус"); 
  rangeRM.getCell(1,6).setValue("Дата создания"); 
  rangeRM.getCell(1,7).setValue("Последнее обновление"); 
  rangeRM.getCell(1,8).setValue("Всего товаров"); 
  rangeRM.getCell(1,9).setValue("Оборудование"); 
  rangeRM.getCell(1,10).setValue("Гарантия"); 
  rangeRM.getCell(1,11).setValue("Ответственный"); 

  let textRespRM = APIRequest("https://support.redmine.ru/issues.xml?project_id=166&status_id=1&limit=100"); //Авторизовывает в системе и возвращает страницу, Test
  // заменить ссылку на портал
  
  //Данные по началу колокнок
  let columnBegin = 1;
  let lineBegin = 2;
  let cellAction = rangeRM.getCell(lineBegin,columnBegin);
  
  //Парсер XML redMine
  let documentRM = XmlService.parse(textRespRM);
  let rootRM = documentRM.getRootElement();
  
  let issueRMs = rootRM.getChildren('issue');
  issueRMs.forEach(issueRM => {
    let idRM = issueRM.getChild('id').getText();
    //Logger.log (idRM);
    idRichText = SpreadsheetApp.newRichTextValue()    // Какая-то неведомая хня, которая вставляет в ячейку текст со ссылкой
                .setText(idRM)
                .setLinkUrl('https://support.redmine.ru/issues/' + idRM) // заменить ссылку на портал
                .build();
    cellAction.setRichTextValue(idRichText);
    //cellAction.setValue(idRM);  //Запись информации в ячейку, Это просто вставляет данные без ссылки
    cellAction = rangeRM.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки

    let trackerRM = issueRM.getChild('tracker').getAttribute('name').getValue();
    cellAction.setValue(trackerRM);  //Запись информации в ячейку
    cellAction = rangeRM.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки

    //Разделяет поля subject и Записывает в ячейку номер обращения
    cellAction.setNumberFormat('@STRING@');
    let subjectnumRM = GetSubjectNumber(issueRM.getChild('subject').getText());
    cellAction.setValue(subjectnumRM);  //Запись информации в ячейку
    cellAction = rangeRM.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки
    
    //Разделяет поля subject и Записывает в ячейку имя контрагента 
    let subjectkaRM = GetSubjectNameKA(issueRM.getChild('subject').getText());
    cellAction.setValue(subjectkaRM);  //Запись информации в ячейку
    cellAction = rangeRM.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки

    /*// Записывает в ячейку номер обращения и контрагента, тот текст который есть в subject
    let subjectcheckRM = issueRM.getChild('subject').getText();
    cellAction.setValue(subjectcheckRM);  //Запись информации в ячейку
    cellAction = rangeRM.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки
    */

    let statusRM = issueRM.getChild('status').getAttribute('name').getValue();
    cellAction.setValue(statusRM);  //Запись информации в ячейку
    cellAction = rangeRM.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки

    let createdDataRM = GetCurrentDataTime(issueRM.getChild('created_on').getText());
    cellAction.setValue(createdDataRM);  //Запись информации в ячейку
    cellAction = rangeRM.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки

    let updatedDataRM = GetCurrentDataTime(issueRM.getChild('updated_on').getText());
    cellAction.setValue(updatedDataRM);  //Запись информации в ячейку
    cellAction = rangeRM.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки

    // Обработка текста из description и Вывод количества товара
    let descTotalGoodsRM = GetDescTotalGoods(issueRM.getChild('description').getText());
    cellAction.setValue(descTotalGoodsRM);  //Запись информации в ячейку
    cellAction.setHorizontalAlignment("center").setVerticalAlignment("middle");
    cellAction = rangeRM.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки
    
    // Обработка текста из description и вывод Первого артикула
    let descFirstArticleRM = GetDescFirstArticle(issueRM.getChild('description').getText());
    cellAction.setNumberFormat('@STRING@');
    descFirArtRichText = SpreadsheetApp.newRichTextValue()    // Какая-то неведомая хня, которая вставляет в ячейку текст со ссылкой
                .setText(descFirstArticleRM)
                .setLinkUrl('https://shop.nag.ru/search?search=' + descFirstArticleRM)
                .build();
    cellAction.setRichTextValue(descFirArtRichText);
    //cellAction.setValue(descFirstArticleRM);  //Запись информации в ячейку, Это просто вставляет данные без ссылки
    cellAction.trimWhitespace(); //Удалить все проделы в ячейке в начале и конце слова
    //https://shop.nag.ru/search?search=EHWIC-D-8ESG-P    вставить ссылку поиска на магазин NAG
    cellAction = rangeRM.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки

    // Обработка текста из description и вывод Гарантии по перву артикулу
    let descWarrantyFirstRM = GetDescWarrantyFirst(issueRM.getChild('description').getText());
    cellAction.setValue(descWarrantyFirstRM);  //Запись информации в ячейку
    cellAction = rangeRM.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки

    //опуститься на строчку ниже
    columnBegin = 1;
    cellAction = rangeRM.getCell(++lineBegin,columnBegin); //Перезапись координат со сдвигом строки
    
  });

  dateNameNewSheetRM.autoResizeColumn(1);
  dateNameNewSheetRM.autoResizeColumn(2);
  dateNameNewSheetRM.autoResizeColumn(3);
  dateNameNewSheetRM.autoResizeColumn(4);
  dateNameNewSheetRM.autoResizeColumn(5);
  dateNameNewSheetRM.autoResizeColumn(6);
  dateNameNewSheetRM.autoResizeColumn(7);
  dateNameNewSheetRM.autoResizeColumn(8);
  dateNameNewSheetRM.autoResizeColumn(9);
  dateNameNewSheetRM.autoResizeColumn(10);
  dateNameNewSheetRM.autoResizeColumn(11);

}

//
//=================РАЗДЕЛ ФУНКЦИЙ ОТЧЕТА==2. Обработка листа с текущей даты===============================================================================================================
//

//Функуия обработки даты из ячейки
function GetCellCurrentDay (cellData) {

  let dd = String(cellData .getDate()).padStart(2, '0');
  let mm = String(cellData .getMonth() + 1).padStart(2, '0'); //January is 0!
  let yyyy = cellData .getFullYear();
  cellReturnData = yyyy + mm + dd;
  //Logger.log (currentDay);
  return cellReturnData;

}

function GetValueData (dataold) {
  editData = dataold.split(" / ")[0];
  editData = editData.replace('-', '').replace('-', '');
  return editData
}

function FilterData (){
  let sheetGeneral = globalTable.getSheetByName('Общее'); // подключаем станицу Общее
  let sheetDataRerort = globalTable.getSheetByName(GetCurrentDay()); //подключаем станицу с Текущей датой

  let rangeGeneral = sheetGeneral.getRange("A1:ZZ1000"); //Определение используемых ячеек
  let rangeDataRerort = sheetDataRerort.getRange("A1:ZZ1000"); //Определение используемых ячеек

  let controlData = GetCellCurrentDay(rangeGeneral.getCell(1,8).getValue()); //Получаем дату из ячейка на странице Общее
  //Logger.log (contralData);

  let cellRowEndGeneral = sheetGeneral.getLastRow(); // Получение последней не пустой стороки с листа Общее. Важно ни чего не писать лишнего в таблицу
  //Logger.log (cellRowEndGeneral);
  let cellRowEndDataRerort = sheetDataRerort.getLastRow(); // Получение последней не пустой стороки с листа Текущей даты
  //Logger.log (cellRowEndDataRerort);

  
  rangeGeneral.getCell(cellRowEndGeneral + 1, 1).setValue(GetCurrentDay()); // Вводим дату в ячейку
  rangeGeneral.getCell(cellRowEndGeneral + 1, 2).setValue(cellRowEndDataRerort - 1); // Вводим кол-во обращений в ячеку

  // Обработки листа Текущей даты по контрольной дате, не превышеные красим зеленым фоном
  for ( let i = 2; i<=cellRowEndDataRerort; i++ ) {
    rmDataCreade = sheetDataRerort.getRange(i, 6).getValue();
    //Logger.log (rmDataCreade);
    rmEditDataCreade = GetValueData(rmDataCreade)
    //Logger.log (rmEditDataCreade);
    
    if (controlData <= rmEditDataCreade) {
      for (var row = 1; row <= 11; row++) {  // 11 потому что количество столбцов
        sheetDataRerort.getRange(i, row).setBackground('#76e34b');  
      }
    }
  }
  
  let quantityOverReport = 0 //Подсчет просроченных обращений

  for ( let j = 2; j<=cellRowEndDataRerort; j++ ) {
        
    if (rangeDataRerort.getCell(j, 1).getBackground() != '#76e34b') {
      ++quantityOverReport
    }
  }

  rangeGeneral.getCell(cellRowEndGeneral + 1, 3).setValue(quantityOverReport); // Вводим кол-во просроченных обращений в ячеку

  // Часть по добавлению ответсвенных по обращению
  let sheetEngineer = globalTable.getSheetByName('Инженеры'); // подключаем станицу Инженеры
  let rangeEngineer = sheetEngineer.getRange("A2:A14"); // Определение используемых ячеек где прописаны инженеры
  let ruleEngineer = SpreadsheetApp.newDataValidation().requireValueInRange(rangeEngineer).build();  // получаем выпадающий список инженеров
  //rangeGeneral.getCell(5, 5).setDataValidation(ruleEngineer);

  // Дополнение листа Текущей даты по установке выподающих списков инженеров
  for ( let ii = 2; ii<=cellRowEndDataRerort; ii++ ) {
    rangeDataRerort.getCell(ii, 11).setDataValidation(ruleEngineer);
  }

}


//
//=================РАЗДЕЛ ФУНКЦИЙ ОТЧЕТА==3. Отправка на почту===============================================================================================================
//

function EmailSend (){
  let sheetESDataRerort = globalTable.getSheetByName(GetCurrentDay()); //подключаем станицу с Текущей датой
  let rangeESDataRerort = sheetESDataRerort.getRange("A1:ZZ1000"); //Определение используемых ячеек
  
  let cellESRowEndDataRerort = sheetESDataRerort.getLastRow(); // Получение последней не пустой стороки с листа Текущей даты

  let sheetEngineerES = globalTable.getSheetByName('Инженеры'); // подключаем станицу Инженеры
  let rangeEngineerES = sheetEngineerES.getRange("A2:A14"); // Определение используемых ячеек где прописаны инженеры
  let listEngineerES = rangeEngineerES.getValues();
  //Logger.log (listEngineerES[2]);

  let messageToEMailES = "Коллеги, добрый день.<br>Прошу проверить и взять в работу обращения: <br>";
  

  for ( let iii = 0; iii<13; iii++) { //Получаем инженеров по индексу
    //Logger.log (listEngineerES[iii]); //Получаем инженеров по индексу
    for ( let jj = 2; jj<=cellESRowEndDataRerort; jj++ ) {  // Цикл обработки листа Текущей даты
      //Logger.log (rangeESDataRerort.getCell(jj, 11).getValue());
      
      //if (rangeDataRerort.getCell(jj, 1).getBackground() != '#76e34b' && rangeESDataRerort.getCell(jj, 11).getValue() == listEngineerES[iii]) {
      if (rangeESDataRerort.getCell(jj, 11).getValue() == listEngineerES[iii]) {
        messageToEMailES = messageToEMailES + '<br>' + '<b>' + listEngineerES[iii] + '</b>'; //Добавление к тексту письма Фамиилии инженера
        messageToEMailES = messageToEMailES + '<p style="margin-left: 50px;">' + rangeESDataRerort.getCell(1, 1).getValue() + ':  ' + rangeESDataRerort.getCell(jj, 1).getValue() + '</p>' + 'https://support.redmine.ru/issues/' + rangeESDataRerort.getCell(jj, 1).getValue();
        for ( let jjj = 2; jjj <= 10; jjj++) {
          messageToEMailES = messageToEMailES + '<p style="margin-left: 50px;">' + rangeESDataRerort.getCell(1, jjj).getValue() + ':  ' + rangeESDataRerort.getCell(jj, jjj).getValue() + '</p>';
        }
        //Logger.log ('сравнение работет');
      }
    }
  }
  
  let emailAddressES = 'name1@domen.ru' + ',' + 'name2@domen.ru';
  //let emailMyAddressES = 'copy mail';
  let subjectES = 'DIAG Отчет обращений без взятия в работу более 3 дней от ' + GetCurrentDay();

  MailApp.sendEmail({
    to: emailAddressES,
    //cc: emailMyAddressES,
    subject: subjectES,
    htmlBody: messageToEMailES,
  });

}
