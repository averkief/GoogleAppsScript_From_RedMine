//--
//  --Запрос с авторизацией в RedMine, возвращет страницу--
//--
function APIRequest (reqUrl) {
  //let key = 'Тут ключ'; //Ключ авторизации в redmine
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
//  --Функция получения текущей даты вывод будет в виде 30.12.2021
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

//--
//  --Функция получения текущей даты ЧП Екатеринбург вид 2021-12-13T11:15:16.000Z
//--
function GetNowDataTime () {

  let nowDataTime = new Date();  // Получаем текущую дату и время, часовой пояс Нью-Йорк -10 от Екб
  nowDataTime.setHours(nowDataTime.getHours() + 5);  // Преобразуем часы добавив 5 часов, для добавления времени Екб от Гринвича
  nowDataTime.setMilliseconds(0); // установить нули в миллисекундах
  let nowDataTimeISO = nowDataTime.toISOString(); // Преобразуем формат стандарта ISO и смены часового пояса на Лондон
  nowDataTimeISO = DataTimeToNumber(nowDataTimeISO);

  return nowDataTimeISO
}

//--
//  --Функция получения текущей даты ЧП Екатеринбург в виде числа 
//--
function DataTimeToNumber (time4Num) {

  numDataTime = Number(time4Num.replace('-', '').replace('-', '').replace('T', '').replace(':', '').replace(':', '').replace('.000', '').replace('Z', ''));
  
  return numDataTime;
}

function DataTimeToNumberXML (time4NumXML) {

  numDataTimeXML = Number(time4NumXML.replace('-', '').replace('-', '').replace(' / ', '').replace(':', '').replace(':', ''));
  
  return numDataTimeXML;
}

//--
//  --Функция редактирования даты 2021-10-04T05:08:20Z в 2021-10-04 / 05:08:20
//--
function GetCurrentDataTime (dataTime) {
  dataTimeCurrent = dataTime.replace('Z', '').replace('T', ' / ');
  return dataTimeCurrent;
}

// Начало ввода глобальных переменных и подключение таблиц

var keyAccess = SpreadsheetApp.openById("=====ссылка_на_таблицу_с_ключем_доступа=====").getSheetByName('Key').getRange("A1:A1").getCell(1, 1).getValue();

var globaRMparserTable = SpreadsheetApp.openById("=====ссылка_на_рабочую_таблицу====="); //Выбор таблицы по парсингу, получаем доступ к таблице по ID
var globaRMparserSheetService = globaRMparserTable.getSheetByName('Service'); // подключаем станицу "Service"
var globaRMparserRangeService = globaRMparserSheetService.getRange("A1:ZZ1000"); //Определение используемых ячеек в таблице PARSER

//var globaRMreportTable = SpreadsheetApp.openById("----------НЕТ ТАБЛИЦЫ-------"); //Выбор таблицы по выводу отчета парсинга, получаем доступ к таблице по ID
//var globaRMreportSheet = globaRMreportTable.getSheetByName('Отчет по парсеру'); // подключаем станицу "Отчет по парсеру"
//var globaRMreportRange = globaRMreportSheet.getRange("A1:ZZ1000"); //Определение используемых ячеек в таблице ОТЧЕТЫ

var numberParserScript = globaRMparserRangeService.getCell(1, 2).getValue(); // Начала счетчика количества запуска отчета

var globalUrl = 'https://support.redmine.ru/issues.xml?project_id=1';  
// заменить ссылку на портал

// Конец глобальных переменных


//--
//  --Функция парсинга всех запроса файла XML из RedMine проект NAG TAG
//--
function ParserNAGTAGissues (){
  numberParserScript++;  // добавление счетчика запросов
  globaRMparserRangeService.getCell(1, 2).setValue(numberParserScript);  // 
  
  //создание нового листа с номером и датой парсинга
  let globaRMparserSheet = globaRMparserTable.getSheetByName(numberParserScript + ' - Все ' + GetCurrentDay());  // Создаем новый лист в текущей таблице
  if (globaRMparserSheet != null) {                             // провека есть ли такой листа в таблице, если есть выводит Sheet, если нет null
    globaRMparserTable.deleteSheet(globaRMparserSheet);                  // удаление листа если такой есть
  }

  globaRMparserSheet = globaRMparserTable.insertSheet();           //Создаем новый лист
  globaRMparserSheet.setName(numberParserScript + ' - Все' + GetCurrentDay());                    // установим имя этому листу

  let globaRMparserRange = globaRMparserSheet.getRange("A1:ZZ1000"); //Определение используемых ячеек

  // установка наименований в столбец листа Парсинга
  let nameColumsParser = ["Номер RM", "Статус", "Приоритет", "Назначена", "Дата начала", "Контрагент", "Последний от комментарий", "Дата комментария"] 
  let nameColumsParserItem = nameColumsParser.length // подсчет кол-ва в массиве
  for ( let i = 0; i<=nameColumsParserItem; i++) {
    globaRMparserRange.getCell(1,i+1).setValue(nameColumsParser[i]);   // Записать данные в ячейку
  }


  // ссылка получения первичных данных, выводит только первые 25 обращений
  let xmlParserItem = APIRequest(globalUrl); //Авторизовывает в системе и возвращает страницу, Test

  //Начальная ячейка для ввода данных по парсингу всех обращений
  let lineBegin = 2;
  let columnBegin = 1;
  let cellActionParser = globaRMparserRange.getCell(lineBegin,columnBegin);

  //Парсер XML redMine
  let xmlDocumentRMItem = XmlService.parse(xmlParserItem);
  let xmlRootRMItem = xmlDocumentRMItem.getRootElement();
  let issuesRMItemTotal = xmlRootRMItem.getAttribute('total_count').getValue();
        Logger.log(issuesRMItemTotal) // эти данные нужно переместить в Таблицу Отчета

  let issuesRMItemPage = Math.ceil(issuesRMItemTotal/25) //расчет количества страниц
        Logger.log(issuesRMItemPage)

  for ( let i = 1; i<=issuesRMItemPage; i++) {
    let xmlParserPage = APIRequest(globalUrl + '&page=' + i); //
    let xmlDocumentRMPage = XmlService.parse(xmlParserPage);
    let xmlRootRMPage = xmlDocumentRMPage.getRootElement();
    Logger.log (globalUrl + '&page=' + i);

    let xmlIssuesRMPage = xmlRootRMPage.getChildren('issue');
    xmlIssuesRMPage.forEach(xmlIssueRMPage => {
      //Получение ID обращения 
      let xmlIdRMPages = xmlIssueRMPage.getChild('id').getText(); 
      //Logger.log (xmlIdRMPages);
        issueIdRichText = SpreadsheetApp.newRichTextValue()    // Какая-то неведомая хня, которая вставляет в ячейку текст со ссылкой
                  .setText(xmlIdRMPages)
                  .setLinkUrl('https://support.redmine.ru/issues/' + xmlIdRMPages)
                  .build();
      cellActionParser.setRichTextValue(issueIdRichText);
      //cellActionParser.setValue(xmlIdRMPages);  //Запись информации в ячейку, Это просто вставляет данные без ссылки
      cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки

      //Получение Статуса обращения
      let xmlStatusRMPages = xmlIssueRMPage.getChild('status').getAttribute('name').getValue();
      cellActionParser.setValue(xmlStatusRMPages);  //Запись информации в ячейку
      cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки
 
      //Получение Приоритета обращения
      let xmlPriorityRMPages = xmlIssueRMPage.getChild('priority').getAttribute('name').getValue();
      cellActionParser.setValue(xmlPriorityRMPages);  //Запись информации в ячейку
      cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки
      

      //Logger.log (xmlIssueRMPage.getChild('assigned_to'));
      if ( xmlIssueRMPage.getChild('assigned_to') != null) {
        let xmlAssignedRMPages = xmlIssueRMPage.getChild('assigned_to').getAttribute('name').getValue();
        cellActionParser.setValue(xmlAssignedRMPages);  //Запись информации в ячейку
        cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки

      /* Получить Id пользователя на которого назначена
        let xmlAssignedRMPagesId = xmlIssueRMPage.getChild('assigned_to').getAttribute('id').getValue();
        cellActionParser.setValue(xmlAssignedRMPagesId);  //Запись информации в ячейку
        cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки
        */
      }
      else {
        let xmlAssignedRMPages = ">>Не назначена<<";
        cellActionParser.setValue(xmlAssignedRMPages);  //Запись информации в ячейку
        cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки
      }
      /*
      //Получение Приоритета обращения
      let xmlAssignedRMPages = xmlIssueRMPage.getChild('assigned_to').getAttribute('name').getValue();
      cellActionParser.setValue(xmlAssignedRMPages);  //Запись информации в ячейку
      cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки
      */
      /*Получение темы обращения
      let xmlSubjectRMPages = xmlIssueRMPage.getChild('subject').getText(); 
      cellActionParser.setValue(xmlSubjectRMPages);  //Запись информации в ячейку, Это просто вставляет данные без ссылки
      cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки
      */

      let xmlCreatedRMPages = GetCurrentDataTime(xmlIssueRMPage.getChild('created_on').getText());
      cellActionParser.setValue(xmlCreatedRMPages);  //Запись информации в ячейку
      cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки
      

      columnBegin = 1; //Вернуть значение начальной колонки в первую
      cellActionParser = globaRMparserRange.getCell(++lineBegin,columnBegin); //Перезапись координат со сдвигом на строчку ниже

    }); // Конец цикла обработки одной страницы Page   
  } // Конец цикла обработки всех страниц Page


  //Записываем значения столбца в массив переменную
  let cellRowEndParserSheet = globaRMparserSheet.getLastRow();
  //Logger.log(cellRowEndParserSheet);
  let allNumbersRequest  = globaRMparserSheet.getRange(2, 1, cellRowEndParserSheet-1).getValues();
  //Logger.log(allNumbersRequest);
  let allNumbersRequestItem = allNumbersRequest.length; // подсчет кол-ва в массиве, общее кол-во обращений
  
  //Начальная ячейка для ввода данных по парсингу каждого обращения
  let lineBeginTwo = 2;
  let columnBeginTwo = 6;
  let cellActionIssue = globaRMparserRange.getCell(lineBeginTwo,columnBeginTwo);

    //Парсер каждого обращения
  for ( let i = 0; i<allNumbersRequestItem; i++) {
    Logger.log ("https://support.redmine.ru/issues/" + allNumbersRequest[i] + ".xml?include=journals");
    let xmlParserJournals = APIRequest("https://support.redmine.ru/issues/" + allNumbersRequest[i] + ".xml?include=journals"); //
    
    let xmlDocumentRMJournals = XmlService.parse(xmlParserJournals);
    let xmlRootRMJournals = xmlDocumentRMJournals.getRootElement();


    let xmlFieldsChild = xmlRootRMJournals.getChild('custom_fields');
    let xmlFieldChildrens = xmlFieldsChild.getChildren('custom_field');
    let xmlKaChildrens = xmlFieldChildrens[10].getValue();
    cellActionIssue.setValue(xmlKaChildrens);  //Запись информации в ячейку
    cellActionIssue = globaRMparserRange.getCell(lineBeginTwo,++columnBeginTwo); //Перезапись координат со сдвигом колонки

    let xmlJournalsChild = xmlRootRMJournals.getChild('journals');
    let xmlJournalChildrens = xmlJournalsChild.getChildren('journal');
    
    let xmlJournalChildrensLast = xmlJournalChildrens[xmlJournalChildrens.length - 1] //получить последний элемент из списка
    Logger.log (xmlJournalChildrensLast);
    
    if (xmlJournalChildrensLast != null) { 
      let journalNameLast = xmlJournalChildrensLast.getChild('user').getAttribute('name').getValue();
      cellActionIssue.setValue(journalNameLast);  //Запись информации в ячейку
      let journalNotesLast = xmlJournalChildrensLast.getChild('notes').getText();
      cellActionIssue.setNote(journalNotesLast);
      //тут команда для установки комментария
      cellActionIssue = globaRMparserRange.getCell(lineBeginTwo,++columnBeginTwo); //Перезапись координат со сдвигом колонки

      let journalCreatedLast = GetCurrentDataTime(xmlJournalChildrensLast.getChild('created_on').getText());
      cellActionIssue.setValue(journalCreatedLast);  //Запись информации в ячейку
      if ( 50000 <= (GetNowDataTime() - DataTimeToNumberXML(journalCreatedLast)) < 1000000 ) {    // Заливка Желтый вычитание дат для получения времени просроченного ответа
        cellActionIssue.setBackground('#f7f56a');
      } //Конец условия цвета
      if ( 1000000 <= (GetNowDataTime() - DataTimeToNumberXML(journalCreatedLast)) < 2000000 ) {    // Заливка Розовый вычитание дат для получения времени просроченного ответа
        cellActionIssue.setBackground('#f78b6a');
      } //Конец условия цвета
      if ( 2000000 < (GetNowDataTime() - DataTimeToNumberXML(journalCreatedLast))) {    // Заливка Красный вычитание дат для получения времени просроченного ответа
        cellActionIssue.setBackground('#f53838');
      } //Конец условия цвета
      cellActionIssue = globaRMparserRange.getCell(lineBeginTwo,++columnBeginTwo); //Перезапись координат со сдвигом колонки    
    }
    else {
      cellActionIssue.setValue(">>Не обработана роботом<<")
    }
    columnBeginTwo = 6; //Вернуть значение начальной колонки в первую
    cellActionIssue = globaRMparserRange.getCell(++lineBeginTwo,columnBeginTwo); //Перезапись координат со сдвигом на строчку ниже

  }

  //Выравнивание колонок по ширине
  let cellComEndParserSheet = globaRMparserSheet.getLastColumn(); // Получение последней не пустой колонки с листа ПарсераДата.

  for (let i = 1; i<=cellComEndParserSheet; i++) {
    globaRMparserSheet.autoResizeColumn(i);
  }  // Конец цикла выравнивания колонок

  Browser.msgBox("Отчет составлен")
} // КОНЕЦ ФУНКЦИИ ParserNAGTAGissues


//--
//  --Функция парсинга обращений не назначеных файла XML из RedMine проект NAG TAG
//--
function NullNAGTAGissues (){
  numberParserScript++  // добавление счетчика запросов
  globaRMparserRangeService.getCell(1, 2).setValue(numberParserScript);  // 
  
  //создание нового листа с номером и датой парсинга
  let globaRMparserSheet = globaRMparserTable.getSheetByName(numberParserScript + ' - не назначена ' + GetCurrentDay());  // Создаем новый лист в текущей таблице
  if (globaRMparserSheet != null) {                             // провека есть ли такой листа в таблице, если есть выводит Sheet, если нет null
    globaRMparserTable.deleteSheet(globaRMparserSheet);                  // удаление листа если такой есть
  }

  globaRMparserSheet = globaRMparserTable.insertSheet();           //Создаем новый лист
  globaRMparserSheet.setName(numberParserScript + ' - не назначена ' + GetCurrentDay());                    // установим имя этому листу

  let globaRMparserRange = globaRMparserSheet.getRange("A1:ZZ1000"); //Определение используемых ячеек

  // установка наименований в столбец листа Парсинга
  let nameColumsParser = ["Номер RM", "Статус", "Приоритет", "Назначена", "Дата начала", "Контрагент", "Последний от комментарий", "Дата комментария"] 
  let nameColumsParserItem = nameColumsParser.length // подсчет кол-ва в массиве
  for ( let i = 0; i<=nameColumsParserItem; i++) {
    globaRMparserRange.getCell(1,i+1).setValue(nameColumsParser[i]);   // Записать данные в ячейку
  }


  // ссылка получения первичных данных, выводит только первые 25 обращений
  let xmlParserItem = APIRequest("https://support.redmine.ru/issues.xml?project_id=1"); //Авторизовывает в системе и возвращает страницу, Test
  // заменить ссылку на портал

  //Начальная ячейка для ввода данных по парсингу всех обращений
  let lineBegin = 2;
  let columnBegin = 1;
  let cellActionParser = globaRMparserRange.getCell(lineBegin,columnBegin);

  //Парсер XML redMine
  let xmlDocumentRMItem = XmlService.parse(xmlParserItem);
  let xmlRootRMItem = xmlDocumentRMItem.getRootElement();
  let issuesRMItemTotal = xmlRootRMItem.getAttribute('total_count').getValue();
        Logger.log(issuesRMItemTotal) // эти данные нужно переместить в Таблицу Отчета

  let issuesRMItemPage = Math.ceil(issuesRMItemTotal/25) //расчет количества страниц
        Logger.log(issuesRMItemPage)

  for ( let i = 1; i<=issuesRMItemPage; i++) {
    let xmlParserPage = APIRequest("https://support.redmine.ru/issues.xml?project_id=1&page=" + i); // заменить ссылку на портал
    let xmlDocumentRMPage = XmlService.parse(xmlParserPage);
    let xmlRootRMPage = xmlDocumentRMPage.getRootElement();
    Logger.log ("https://support.redmine.ru/issues.xml?project_id=1&page=" + i); // заменить ссылку на портал

    let xmlIssuesRMPage = xmlRootRMPage.getChildren('issue');
    xmlIssuesRMPage.forEach(xmlIssueRMPage => {
      
      if ( xmlIssueRMPage.getChild('assigned_to') == null) {
      
      //Получение ID обращения
      let xmlIdRMPages = xmlIssueRMPage.getChild('id').getText(); 
      //Logger.log (xmlIdRMPages);
        issueIdRichText = SpreadsheetApp.newRichTextValue()    // Какая-то неведомая хня, которая вставляет в ячейку текст со ссылкой
                  .setText(xmlIdRMPages)
                  .setLinkUrl('https://support.nag.ru/issues/' + xmlIdRMPages)
                  .build();
      cellActionParser.setRichTextValue(issueIdRichText);
      //cellActionParser.setValue(xmlIdRMPages);  //Запись информации в ячейку, Это просто вставляет данные без ссылки
      cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки

      //Получение Статуса обращения
      let xmlStatusRMPages = xmlIssueRMPage.getChild('status').getAttribute('name').getValue();
      cellActionParser.setValue(xmlStatusRMPages);  //Запись информации в ячейку
      cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки
 
      //Получение Приоритета обращения
      let xmlPriorityRMPages = xmlIssueRMPage.getChild('priority').getAttribute('name').getValue();
      cellActionParser.setValue(xmlPriorityRMPages);  //Запись информации в ячейку
      cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки
      

      //Logger.log (xmlIssueRMPage.getChild('assigned_to'));
      if ( xmlIssueRMPage.getChild('assigned_to') != null) {
        let xmlAssignedRMPages = xmlIssueRMPage.getChild('assigned_to').getAttribute('name').getValue();
        cellActionParser.setValue(xmlAssignedRMPages);  //Запись информации в ячейку
        cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки

      }
      else {
        let xmlAssignedRMPages = ">>Не назначена<<";
        cellActionParser.setValue(xmlAssignedRMPages);  //Запись информации в ячейку
        cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки
      }
      
      let xmlCreatedRMPages = GetCurrentDataTime(xmlIssueRMPage.getChild('created_on').getText());
      cellActionParser.setValue(xmlCreatedRMPages);  //Запись информации в ячейку
      cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки
      

      columnBegin = 1; //Вернуть значение начальной колонки в первую
      cellActionParser = globaRMparserRange.getCell(++lineBegin,columnBegin); //Перезапись координат со сдвигом на строчку ниже

      } // Конец условия assigned_to') = nul

    }); // Конец цикла обработки одной страницы Page   
  } // Конец цикла обработки всех страниц Page


  //Записываем значения столбца в массив переменную
  let cellRowEndParserSheet = globaRMparserSheet.getLastRow();
  //Logger.log(cellRowEndParserSheet);
  let allNumbersRequest  = globaRMparserSheet.getRange(2, 1, cellRowEndParserSheet-1).getValues();
  //Logger.log(allNumbersRequest);
  let allNumbersRequestItem = allNumbersRequest.length; // подсчет кол-ва в массиве, общее кол-во обращений
  
  //Начальная ячейка для ввода данных по парсингу каждого обращения
  lineBeginTwo = 2;
  columnBeginTwo = 6;
  let cellActionIssue = globaRMparserRange.getCell(lineBeginTwo,columnBeginTwo);

    //Парсер каждого обращения
  for ( let i = 0; i<allNumbersRequestItem; i++) {
    let xmlParserJournals = APIRequest("https://support.redmine.ru/issues/" + allNumbersRequest[i] + ".xml?include=journals"); //
    let xmlDocumentRMJournals = XmlService.parse(xmlParserJournals);
    let xmlRootRMJournals = xmlDocumentRMJournals.getRootElement();
    Logger.log ("https://support.redmine.ru/issues/" + allNumbersRequest[i] + ".xml?include=journals");
    
    let xmlFieldsChild = xmlRootRMJournals.getChild('custom_fields');
    let xmlFieldChildrens = xmlFieldsChild.getChildren('custom_field');
    let xmlKaChildrens = xmlFieldChildrens[10].getValue();
    cellActionIssue.setValue(xmlKaChildrens);  //Запись информации в ячейку
    cellActionIssue = globaRMparserRange.getCell(lineBeginTwo,++columnBeginTwo); //Перезапись координат со сдвигом колонки

    let xmlJournalsChild = xmlRootRMJournals.getChild('journals');
    let xmlJournalChildrens = xmlJournalsChild.getChildren('journal');
    
    let xmlJournalChildrensLast = xmlJournalChildrens[xmlJournalChildrens.length - 1] //получить последний элемент из списка
    Logger.log (xmlJournalChildrensLast);
    
    if (xmlJournalChildrensLast != null) { 
      let journalNameLast = xmlJournalChildrensLast.getChild('user').getAttribute('name').getValue();
      cellActionIssue.setValue(journalNameLast);  //Запись информации в ячейку
      let journalNotesLast = xmlJournalChildrensLast.getChild('notes').getText();
      cellActionIssue.setNote(journalNotesLast);
      //тут команда для установки комментария
      cellActionIssue = globaRMparserRange.getCell(lineBeginTwo,++columnBeginTwo); //Перезапись координат со сдвигом колонки

      let journalCreatedLast = GetCurrentDataTime(xmlJournalChildrensLast.getChild('created_on').getText());
      cellActionIssue.setValue(journalCreatedLast);  //Запись информации в ячейку
      if ( 50000 <= (GetNowDataTime() - DataTimeToNumberXML(journalCreatedLast)) < 1000000 ) {    // Заливка Желтый вычитание дат для получения времени просроченного ответа
        cellActionIssue.setBackground('#f7f56a');
      } //Конец условия цвета
      if ( 1000000 <= (GetNowDataTime() - DataTimeToNumberXML(journalCreatedLast)) < 2000000 ) {    // Заливка Розовый вычитание дат для получения времени просроченного ответа
        cellActionIssue.setBackground('#f78b6a');
      } //Конец условия цвета
      if ( 2000000 < (GetNowDataTime() - DataTimeToNumberXML(journalCreatedLast))) {    // Заливка Красный вычитание дат для получения времени просроченного ответа
        cellActionIssue.setBackground('#f53838');
      } //Конец условия цвета
      cellActionIssue = globaRMparserRange.getCell(lineBeginTwo,++columnBeginTwo); //Перезапись координат со сдвигом колонки    
    }
    else {
      cellActionIssue.setValue(">>Не обработана роботом<<")
    }
    columnBeginTwo = 6; //Вернуть значение начальной колонки в первую
    cellActionIssue = globaRMparserRange.getCell(++lineBeginTwo,columnBeginTwo); //Перезапись координат со сдвигом на строчку ниже

  }

  //Выравнивание колонок по ширине
  let cellComEndParserSheet = globaRMparserSheet.getLastColumn(); // Получение последней не пустой колонки с листа ПарсераДата.

  for (let i = 1; i<=cellComEndParserSheet; i++) {
    globaRMparserSheet.autoResizeColumn(i);
  }  // Конец цикла выравнивания колонок

  Browser.msgBox("Отчет составлен")
} // КОНЕЦ ФУНКЦИИ NullNAGTAGissues



//--
//  --Функция парсинга проекта NAG TAG по отдельному ответсвенному
//--
function NameNAGTAGissues (nameReport, idReport){
  
  numberParserScript++  // добавление счетчика запросов
  globaRMparserRangeService.getCell(1, 2).setValue(numberParserScript);  // Вставить новое значение в счетчик

  //создание нового листа с номером и датой парсинга
  let globaRMparserSheet = globaRMparserTable.getSheetByName(numberParserScript + ' ' + GetCurrentDay());  // Создаем новый лист в текущей таблице
  if (globaRMparserSheet != null) {                             // провека есть ли такой листа в таблице, если есть выводит Sheet, если нет null
    globaRMparserTable.deleteSheet(globaRMparserSheet);                  // удаление листа если такой есть
  }

  globaRMparserSheet = globaRMparserTable.insertSheet();           //Создаем новый лист
  globaRMparserSheet.setName(numberParserScript + ' - ' + nameReport + ' ' + GetCurrentDay());                    // установим имя этому листу

  let globaRMparserRange = globaRMparserSheet.getRange("A1:ZZ1000"); //Определение используемых ячеек

  // установка наименований в столбец листа Парсинга
  let nameColumsParser = ["Номер RM", "Статус", "Приоритет", "Назначена", "Дата начала", "Контрагент", "Последний от комментарий", "Дата комментария"] 
  let nameColumsParserItem = nameColumsParser.length // подсчет кол-ва в массиве
  for ( let i = 0; i<=nameColumsParserItem; i++) {
    globaRMparserRange.getCell(1,i+1).setValue(nameColumsParser[i]);   // Записать данные в ячейку
  }

  // ссылка получения первичных данных, выводит только первые 25 обращений
  let xmlParserItem = APIRequest(globalUrl + '&assigned_to_id=' + idReport); //Авторизовывает в системе и возвращает страницу, Test
  
  //Начальная ячейка для ввода данных по парсингу всех обращений
  let lineBegin = 2;
  let columnBegin = 1;
  let cellActionParser = globaRMparserRange.getCell(lineBegin,columnBegin);
  Logger.log(cellActionParser)

  //Парсер XML redMine
  let xmlDocumentRMItem = XmlService.parse(xmlParserItem);
  let xmlRootRMItem = xmlDocumentRMItem.getRootElement();
  let issuesRMItemTotal = xmlRootRMItem.getAttribute('total_count').getValue();
        //Logger.log('Обращений:' + issuesRMItemTotal) // эти данные нужно переместить в Таблицу Отчета

  let issuesRMItemPage = Math.ceil(issuesRMItemTotal/25) //расчет количества страниц
        //Logger.log('Страниц:' + issuesRMItemPage)

  for ( let i = 1; i<=issuesRMItemPage; i++) {
    let xmlParserPage = APIRequest(globalUrl + '&assigned_to_id=' + idReport + '&page=' + i); //
    let xmlDocumentRMPage = XmlService.parse(xmlParserPage);
    let xmlRootRMPage = xmlDocumentRMPage.getRootElement();
    Logger.log (globalUrl + '&assigned_to_id=' + idReport + '&page=' + i);

    let xmlIssuesRMPage = xmlRootRMPage.getChildren('issue');
    xmlIssuesRMPage.forEach(xmlIssueRMPage => {
      //Получение ID обращения
      let xmlIdRMPages = xmlIssueRMPage.getChild('id').getText(); 
      //Logger.log (xmlIdRMPages);
        issueIdRichText = SpreadsheetApp.newRichTextValue()    // Какая-то неведомая хня, которая вставляет в ячейку текст со ссылкой
                  .setText(xmlIdRMPages)
                  .setLinkUrl('https://support.redmine.ru/issues/' + xmlIdRMPages)
                  .build();
      cellActionParser.setRichTextValue(issueIdRichText);
      //cellActionParser.setValue(xmlIdRMPages);  //Запись информации в ячейку, Это просто вставляет данные без ссылки
      cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки

      //Получение Статуса обращения
      let xmlStatusRMPages = xmlIssueRMPage.getChild('status').getAttribute('name').getValue();
      cellActionParser.setValue(xmlStatusRMPages);  //Запись информации в ячейку
      cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки
 
      //Получение Приоритета обращения
      let xmlPriorityRMPages = xmlIssueRMPage.getChild('priority').getAttribute('name').getValue();
      cellActionParser.setValue(xmlPriorityRMPages);  //Запись информации в ячейку
      cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки
      

      //Logger.log (xmlIssueRMPage.getChild('assigned_to'));
      let xmlAssignedRMPages = xmlIssueRMPage.getChild('assigned_to').getAttribute('name').getValue();
      cellActionParser.setValue(xmlAssignedRMPages);  //Запись информации в ячейку
      cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки
      
      let xmlCreatedRMPages = GetCurrentDataTime(xmlIssueRMPage.getChild('created_on').getText());
      cellActionParser.setValue(xmlCreatedRMPages);  //Запись информации в ячейку
      cellActionParser = globaRMparserRange.getCell(lineBegin,++columnBegin); //Перезапись координат со сдвигом колонки
      
      columnBegin = 1; //Вернуть значение начальной колонки в первую
      cellActionParser = globaRMparserRange.getCell(++lineBegin,columnBegin); //Перезапись координат со сдвигом на строчку ниже

    }); // Конец цикла обработки одной страницы Page   
  } // Конец цикла обработки всех страниц Page 

  //Записываем значения столбца в массив переменную
  let cellRowEndParserSheet = globaRMparserSheet.getLastRow();
  //Logger.log(cellRowEndParserSheet);
  let allNumbersRequest  = globaRMparserSheet.getRange(2, 1, cellRowEndParserSheet-1).getValues();
  //Logger.log(allNumbersRequest);
  let allNumbersRequestItem = allNumbersRequest.length; // подсчет кол-ва в массиве, общее кол-во обращений
  

  //Начальная ячейка для ввода данных по парсингу каждого обращения
  let lineBeginTwo = 2;
  let columnBeginTwo = 6;
  let cellActionIssue = globaRMparserRange.getCell(lineBeginTwo,columnBeginTwo);

    //Парсер каждого обращения
  for ( let i = 0; i<allNumbersRequestItem; i++) {
    let xmlParserJournals = APIRequest("https://support.redmine.ru/issues/" + allNumbersRequest[i] + ".xml?include=journals"); //
    let xmlDocumentRMJournals = XmlService.parse(xmlParserJournals);
    let xmlRootRMJournals = xmlDocumentRMJournals.getRootElement();
    Logger.log ("https://support.redmine.ru/issues/" + allNumbersRequest[i] + ".xml?include=journals");
    
    let xmlFieldsChild = xmlRootRMJournals.getChild('custom_fields');
    let xmlFieldChildrens = xmlFieldsChild.getChildren('custom_field');
    let xmlKaChildrens = xmlFieldChildrens[10].getValue();
    cellActionIssue.setValue(xmlKaChildrens);  //Запись информации в ячейку
    cellActionIssue = globaRMparserRange.getCell(lineBeginTwo,++columnBeginTwo); //Перезапись координат со сдвигом колонки

    let xmlJournalsChild = xmlRootRMJournals.getChild('journals');
    let xmlJournalChildrens = xmlJournalsChild.getChildren('journal');
    
    let xmlJournalChildrensLast = xmlJournalChildrens[xmlJournalChildrens.length - 1] //получить последний элемент из списка
    Logger.log (xmlJournalChildrensLast);
    
    if (xmlJournalChildrensLast != null) { 
      let journalNameLast = xmlJournalChildrensLast.getChild('user').getAttribute('name').getValue();
      cellActionIssue.setValue(journalNameLast);  //Запись информации в ячейку
      let journalNotesLast = xmlJournalChildrensLast.getChild('notes').getText();
      cellActionIssue.setNote(journalNotesLast);
      //тут команда для установки комментария
      cellActionIssue = globaRMparserRange.getCell(lineBeginTwo,++columnBeginTwo); //Перезапись координат со сдвигом колонки

      let journalCreatedLast = GetCurrentDataTime(xmlJournalChildrensLast.getChild('created_on').getText());
      cellActionIssue.setValue(journalCreatedLast);  //Запись информации в ячейку
      if ( 50000 <= (GetNowDataTime() - DataTimeToNumberXML(journalCreatedLast)) < 1000000 ) {    // Заливка Желтый вычитание дат для получения времени просроченного ответа
        cellActionIssue.setBackground('#f7f56a');
      } //Конец условия цвета
      if ( 1000000 <= (GetNowDataTime() - DataTimeToNumberXML(journalCreatedLast)) < 2000000 ) {    // Заливка Розовый вычитание дат для получения времени просроченного ответа
        cellActionIssue.setBackground('#f78b6a');
      } //Конец условия цвета
      if ( 2000000 < (GetNowDataTime() - DataTimeToNumberXML(journalCreatedLast))) {    // Заливка Красный вычитание дат для получения времени просроченного ответа
        cellActionIssue.setBackground('#f53838');
      } //Конец условия цвета
      cellActionIssue = globaRMparserRange.getCell(lineBeginTwo,++columnBeginTwo); //Перезапись координат со сдвигом колонки    
    }
    else {
      cellActionIssue.setValue(">>Не обработана роботом<<")
    }
    columnBeginTwo = 6; //Вернуть значение начальной колонки в первую
    cellActionIssue = globaRMparserRange.getCell(++lineBeginTwo,columnBeginTwo); //Перезапись координат со сдвигом на строчку ниже

  }

  //Выравнивание колонок по ширине
  let cellComEndParserSheet = globaRMparserSheet.getLastColumn(); // Получение последней не пустой колонки с листа ПарсераДата.

  for (let i = 1; i<=cellComEndParserSheet; i++) {
    globaRMparserSheet.autoResizeColumn(i);
  }  // Конец цикла выравнивания колонок

  Browser.msgBox("Отчет составлен")
} // Конец функции BeginNAGTAGissues






//--
//  --Функция запуска триггера по выбранному критерию
//--
function TriggerButton () {

  if (globaRMparserRangeService.getCell(4, 2).isChecked()===true) {  // По - Всем обращениям
    ParserNAGTAGissues();
  
    globaRMparserRangeService.getCell(4, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 4, 2

  if (globaRMparserRangeService.getCell(5, 2).isChecked()===true) {  // По - Не назначена
    NullNAGTAGissues();
  
    globaRMparserRangeService.getCell(5, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 5, 2

  if (globaRMparserRangeService.getCell(7, 2).isChecked()===true) {  // По - Группа СЦ -
    let nameAssigned = globaRMparserRangeService.getCell(7, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(7, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(7, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 7, 2

  if (globaRMparserRangeService.getCell(8, 2).isChecked()===true) {  // По - ERD
    let nameAssigned = globaRMparserRangeService.getCell(8, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(8, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(8, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 8, 2

  if (globaRMparserRangeService.getCell(9, 2).isChecked()===true) {  // По - SNR Беспроводное оборудование
    let nameAssigned = globaRMparserRangeService.getCell(9, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(9, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(9, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 9, 2

  if (globaRMparserRangeService.getCell(10, 2).isChecked()===true) {  // По - VoIP
    let nameAssigned = globaRMparserRangeService.getCell(10, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(10, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(10, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 10, 2

  if (globaRMparserRangeService.getCell(11, 2).isChecked()===true) {  // По - ДРП media -
    let nameAssigned = globaRMparserRangeService.getCell(11, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(11, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(11, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 11, 2

  if (globaRMparserRangeService.getCell(12, 2).isChecked()===true) {  // По - Видеонаблюдение
    let nameAssigned = globaRMparserRangeService.getCell(12, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(12, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(12, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 12, 2

  if (globaRMparserRangeService.getCell(13, 2).isChecked()===true) {  // По - Беспроводное оборудование
    let nameAssigned = globaRMparserRangeService.getCell(13, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(13, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(13, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 13, 2

  if (globaRMparserRangeService.getCell(14, 2).isChecked()===true) {  // По - Andrey Rudykh
    let nameAssigned = globaRMparserRangeService.getCell(14, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(14, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(14, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 14, 2

  if (globaRMparserRangeService.getCell(15, 2).isChecked()===true) {  // По - Ramil Asharapov
    let nameAssigned = globaRMparserRangeService.getCell(15, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(15, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(15, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 15, 2

  if (globaRMparserRangeService.getCell(16, 2).isChecked()===true) {  // По - Александр Айплатов
    let nameAssigned = globaRMparserRangeService.getCell(16, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(16, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(16, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 16, 2

  if (globaRMparserRangeService.getCell(17, 2).isChecked()===true) {  // По - Александр Кощеев
    let nameAssigned = globaRMparserRangeService.getCell(17, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(17, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(17, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 17, 2
  
  if (globaRMparserRangeService.getCell(18, 2).isChecked()===true) {  // По - Александр Лалыко
    let nameAssigned = globaRMparserRangeService.getCell(18, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(18, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(18, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 18, 2
  
  if (globaRMparserRangeService.getCell(19, 2).isChecked()===true) {  // По - Александр Орлов
    let nameAssigned = globaRMparserRangeService.getCell(19, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(19, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(19, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 19, 2
  
  if (globaRMparserRangeService.getCell(20, 2).isChecked()===true) {  // По - Алексей Сонькин
    let nameAssigned = globaRMparserRangeService.getCell(20, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(20, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(20, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 20, 2
  
  if (globaRMparserRangeService.getCell(21, 2).isChecked()===true) {  // По - Анатолий Лаврентьев
    let nameAssigned = globaRMparserRangeService.getCell(21, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(21, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(21, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 21, 2
  
  if (globaRMparserRangeService.getCell(22, 2).isChecked()===true) {  // По - Андрей Аверкиев
    let nameAssigned = globaRMparserRangeService.getCell(22, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(22, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(22, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 22, 2
  
  if (globaRMparserRangeService.getCell(23, 2).isChecked()===true) {  // По - Андрей Шепелев
    let nameAssigned = globaRMparserRangeService.getCell(23, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(23, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(23, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 23, 2
  
  if (globaRMparserRangeService.getCell(24, 2).isChecked()===true) {  // По - Армен Ероносян
    let nameAssigned = globaRMparserRangeService.getCell(24, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(24, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(24, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 24, 2
  
  if (globaRMparserRangeService.getCell(25, 2).isChecked()===true) {  // По - Афанасий Белюшин
    let nameAssigned = globaRMparserRangeService.getCell(25, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(25, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(25, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 25, 2
  
  if (globaRMparserRangeService.getCell(26, 2).isChecked()===true) {  // По - Виктор Ткаченко
    let nameAssigned = globaRMparserRangeService.getCell(26, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(26, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(26, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 26, 2
  
  if (globaRMparserRangeService.getCell(27, 2).isChecked()===true) {  // По - Владимир Козубский
    let nameAssigned = globaRMparserRangeService.getCell(27, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(27, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(27, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 27, 2
  
  if (globaRMparserRangeService.getCell(28, 2).isChecked()===true) {  // По - Владислав Гуркин
    let nameAssigned = globaRMparserRangeService.getCell(28, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(28, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
      
    globaRMparserRangeService.getCell(28, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 28, 2
  
  if (globaRMparserRangeService.getCell(29, 2).isChecked()===true) {  // По - Дмитрий Брусенцев
    let nameAssigned = globaRMparserRangeService.getCell(29, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(29, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(29, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 29, 2
  
  if (globaRMparserRangeService.getCell(30, 2).isChecked()===true) {  // По - Дмитрий Лизунов
    let nameAssigned = globaRMparserRangeService.getCell(30, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(30, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(30, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 30, 2
  
  if (globaRMparserRangeService.getCell(31, 2).isChecked()===true) {  // По - Дулин Павел Евгеньевич
    let nameAssigned = globaRMparserRangeService.getCell(31, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(31, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(31, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 31, 2
  
  if (globaRMparserRangeService.getCell(32, 2).isChecked()===true) {  // По - Евгений Мирхасанов
    let nameAssigned = globaRMparserRangeService.getCell(32, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(32, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(32, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 32, 2
  
  if (globaRMparserRangeService.getCell(33, 2).isChecked()===true) {  // По - Илья Леушин
    let nameAssigned = globaRMparserRangeService.getCell(33, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(33, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(33, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 33, 2
  
  if (globaRMparserRangeService.getCell(34, 2).isChecked()===true) {  // По - Ирина Загвоздина
    let nameAssigned = globaRMparserRangeService.getCell(34, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(34, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(34, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 34, 2
  
  if (globaRMparserRangeService.getCell(35, 2).isChecked()===true) {  // По - Карина Ероховец
    let nameAssigned = globaRMparserRangeService.getCell(35, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(35, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(35, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 35, 2
  
  if (globaRMparserRangeService.getCell(36, 2).isChecked()===true) {  // По - Кискин Игорь Александрович
    let nameAssigned = globaRMparserRangeService.getCell(36, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(36, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(36, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 36, 2
  
  if (globaRMparserRangeService.getCell(37, 2).isChecked()===true) {  // По - Константин Дадайкин
    let nameAssigned = globaRMparserRangeService.getCell(37, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(37, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(37, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 37, 2
  
  if (globaRMparserRangeService.getCell(38, 2).isChecked()===true) {  // По - Константин Немцев
    let nameAssigned = globaRMparserRangeService.getCell(38, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(38, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(38, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 38, 2
  
  if (globaRMparserRangeService.getCell(39, 2).isChecked()===true) {  // По - Максим Бабинцев
    let nameAssigned = globaRMparserRangeService.getCell(39, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(39, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(39, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 39, 2
  
  if (globaRMparserRangeService.getCell(40, 2).isChecked()===true) {  // По - Максим Ефремов
    let nameAssigned = globaRMparserRangeService.getCell(40, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(40, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(40, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 40, 2
  
  if (globaRMparserRangeService.getCell(41, 2).isChecked()===true) {  // По - Максим Левинзон
    let nameAssigned = globaRMparserRangeService.getCell(41, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(41, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(41, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 41, 2
  
  if (globaRMparserRangeService.getCell(42, 2).isChecked()===true) {  // По - Марат Багандов
    let nameAssigned = globaRMparserRangeService.getCell(42, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(42, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(42, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 42, 2
  
  if (globaRMparserRangeService.getCell(43, 2).isChecked()===true) {  // По - Марат Шайхов
    let nameAssigned = globaRMparserRangeService.getCell(43, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(43, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(43, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 43, 2
  
  if (globaRMparserRangeService.getCell(44, 2).isChecked()===true) {  // По - Микушин Григорий
    let nameAssigned = globaRMparserRangeService.getCell(44, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(44, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(44, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 44, 2
  
  if (globaRMparserRangeService.getCell(45, 2).isChecked()===true) {  // По - Наиль Тимканов
    let nameAssigned = globaRMparserRangeService.getCell(45, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(45, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(45, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 45, 2
  
  if (globaRMparserRangeService.getCell(46, 2).isChecked()===true) {  // По - Никита Девятьяров
    let nameAssigned = globaRMparserRangeService.getCell(46, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(46, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(46, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 46, 2
  
  if (globaRMparserRangeService.getCell(47, 2).isChecked()===true) {  // По - Олег Кузнецов
    let nameAssigned = globaRMparserRangeService.getCell(47, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(47, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(47, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 47, 2
  
  if (globaRMparserRangeService.getCell(48, 2).isChecked()===true) {  // По - Павел Шпилько
    let nameAssigned = globaRMparserRangeService.getCell(48, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(48, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(48, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 48, 2
  
  if (globaRMparserRangeService.getCell(49, 2).isChecked()===true) {  // По - Родион Бендер
    let nameAssigned = globaRMparserRangeService.getCell(49, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(49, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(49, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 49, 2
  
  if (globaRMparserRangeService.getCell(50, 2).isChecked()===true) {  // По - Роман Кудрявцев
    let nameAssigned = globaRMparserRangeService.getCell(50, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(50, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(50, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 50, 2
  
  if (globaRMparserRangeService.getCell(51, 2).isChecked()===true) {  // По - Шайхмуллин Ильяс Гулусович
    let nameAssigned = globaRMparserRangeService.getCell(51, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(51, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(51, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 51, 2
  
  if (globaRMparserRangeService.getCell(52, 2).isChecked()===true) {  // По - Юшков Роман Витальевич
    let nameAssigned = globaRMparserRangeService.getCell(52, 1).getValue();
    let idAssigned = globaRMparserRangeService.getCell(52, 3).getValue();
    NameNAGTAGissues(nameAssigned, idAssigned);
  
    globaRMparserRangeService.getCell(52, 2).uncheck(); // Возвращаем Флаг в исходное положение
  } // Конец условия if cheked 52, 2
  

} // Конец функции TriggerButton


