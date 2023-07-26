// Ввод глобальных переменных и подключение таблиц

const keyAccess = SpreadsheetApp.openById('=====ссылка_на_таблицу_с_ключем_доступа=====').getSheetByName('Key').getRange("A1:A1").getCell(1, 1).getValue(); // получение ключа доступа в Support на базе RedMine
const authHead = { 'Authorization' : 'Basic ' + keyAccess}; // Авторизация в Support

const globalTable = SpreadsheetApp.openById('=====ссылка_на_рабочую_таблицу====='); // Подключаем рабочую таблицу
const globalSheetReport = globalTable.getSheetByName('Report'); // Подключаем станицу "Report"
//const globalRangeReport = globalSheetReport.getRange("A1:ZZ1000"); //Подключаем используемые ячейки на странице "Report"


//--
//  --GET запрос, получение данных--
//--
function GetAPIRequest (reqUrl) {
  
  let optGet = {
    'method'  : 'GET',
    'headers' : authHead,
    'muteHttpExceptions' : true
  };
  
  let response = UrlFetchApp.fetch(reqUrl, optGet);
  return response.getContentText();

} // конец функции GetAPIRequest


//
//=================РАЗДЕЛ ВСПОМОГАТЕЛЬНЫХ ФУНКЦИЙ
//

function ProcessingXML(oneIssue) {

  let issueId = null;
  let issueTracker = null;
  let issueAssigned = null;
  let issueStart = null;
  let issueDue = null;


  issueId = oneIssue.getChild('id').getText();
  issueTracker = oneIssue.getChild('tracker').getAttribute('name').getValue();
  //let issueAssigned = oneIssue.getChild('assigned_to').getAttribute('name').getValue() ? oneIssue.getChild('assigned_to') != null : oneIssue.getChild('author').getAttribute('name').getValue();
  if ( oneIssue.getChild('assigned_to') != null) {
    issueAssigned = oneIssue.getChild('assigned_to').getAttribute('name').getValue()
  }
  else {
    issueAssigned = oneIssue.getChild('author').getAttribute('name').getValue();
  }
  issueStart = oneIssue.getChild('start_date').getText();
  issueDue = oneIssue.getChild('due_date').getText();

  return [issueId, issueTracker, issueAssigned, issueStart, issueDue]
  // Индексы - номер обращения, Тип обращения, Назначена на, Дата создания, Дата завершения
} // завершение функции ProcessingXML


//
//=================РАЗДЕЛ ОСНОВНОЙ ФУНКЦИЙ
//
function MainProcess() {

  // Создание массива и наименований столбцов
  const arreyIssues = [['Номер', 'Тип', 'Ответственный', 'Дата создания', 'Дата завершения']]; 
  
  // Начало получение количетво обращений и расчет страниц
  let beginXMLSuppurt = GetAPIRequest('https://support.redmine.ru/issues.xml?project_id=166&status_id=*&limit=100&created_on=%3E%3D2021-01-01'); //Авторизовывает в системе и возвращает страницу
  // заменить ссылку на портал
  let xmlFirstPageParse = XmlService.parse(beginXMLSuppurt);
  let xmlFirstPageRoot = xmlFirstPageParse.getRootElement();
  let xmlFirstPageTotal = xmlFirstPageRoot.getAttribute('total_count').getValue();  // Получение общего числа обращений в проекте "Сервисный центр"
  Logger.log(xmlFirstPageTotal)
  let valueTotalPage = Math.ceil(xmlFirstPageTotal/100) //расчет количества страниц
  Logger.log(valueTotalPage)

  for ( let i = 1; i<=valueTotalPage; i++) { // потом открыть
    Logger.log('https://support.redmine.ru/issues.xml?project_id=166&status_id=*&limit=100&created_on=%3E%3D2021-01-01&page=' + i)
    let xmlRaportIssues = GetAPIRequest('https://support.redmine.ru/issues.xml?project_id=166&status_id=*&limit=100&created_on=%3E%3D2021-01-01&page=' + i); //Авторизовывает в системе и возвращает страницу
    // заменить ссылку на портал
  
    //Парсер XML redMine
    let xmlPageParse = XmlService.parse(xmlRaportIssues);
    let xmlPageRoot = xmlPageParse.getRootElement();
    let xmlPageIssues = xmlPageRoot.getChildren('issue');
    
    xmlPageIssues.forEach(xmlPageIssue => {
      arreyIssues.push(ProcessingXML(xmlPageIssue));
    });

  }

  //Logger.log(arreyIssues);

  // Вставить данные массива в таблицу
  globalSheetReport.getRange(1, 1, arreyIssues.length, arreyIssues[0].length).setValues(arreyIssues);
  // Выравнивание по стобцам
  for (let i = 1; i <= arreyIssues[0].length; i++) {
    globalSheetReport.autoResizeColumn(i);
  }

} // завершение функции MainProcess

//
//=================ПЕСОЧНИЦА
//

// Если в ячейку вставить дату "07.06.2022" то вывод в скрипте будет "Tue Jun 07 00:00:00 GMT+05:00 2022"

function TestingGo () {
  const nameColumsParserss = [['Номер', 'Тип', 'Ответственный', 'Дата создания', 'Дата завершения']]; 
  //const nameColumsParserItem = nameColumsParser.length // подсчет кол-ва в массиве
  //let www = globalSheetReport.getRange(1, 1, 1, 5).getValues();
  //Logger.log(nameColumsParserss[0].length);
  globalSheetReport.getRange(1, 1, 1, nameColumsParserss[0].length).setValues(nameColumsParserss);
  
  
  //for ( let i = 0; i<=nameColumsParserItem; i++) {
  //  globalRangeReport.getCell(1,i+1).setValue(nameColumsParser[i]);   // Записать данные в ячейку
  //}
  

} // конец функции TestingGo






