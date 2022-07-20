// Ввод глобальных переменных и подключение таблиц

const keyAccess = SpreadsheetApp.openById("=====ссылка_на_таблицу_с_ключем_доступа=====").getSheetByName('Key').getRange("A1:A1").getCell(1, 1).getValue(); // получение ключа доступа в Support
const authHead = { 'Authorization' : 'Basic ' + keyAccess}; // Авторизация в Support

const globalTable = SpreadsheetApp.openById("=====ссылка_на_рабочую_таблицу====="); // Подключаем рабочую таблицу
const globalSheetService = globalTable.getSheetByName('Service'); // Подключаем станицу "Service"
const globalRangeService = globalSheetService.getRange("A1:ZZ1000"); //Подключаем используемые ячейки на странице "Service"

const globalSheetTemplates = globalTable.getSheetByName('Templates'); // Подключаем станицу "Templates"
const globalRangeTemplates = globalSheetService.getRange("A1:ZZ1000"); //Подключаем используемые ячейки на странице "Service"

const globalSheetActive = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Подключаем станицу на которой курсор
const globalRangeActive = globalSheetActive.getRange("A1:ZZ1000"); // Подключаем используемые ячейки страницы на которой курсор

const globalArrayEngineers = globalSheetService.getRange(2, 5, (globalSheetService.getLastRow() - 1), 1).getValues(); // Создание массива с id сотрудников СЦ
let globalNumberScript = globalRangeService.getCell(1,2).getValue(); // Счетчик запуска отчета

const fuckedDays = globalRangeService.getCell(2,2).getValue(); // Получение дней просрочки

const now = new Date();
const nowGrinvich = new Date(now.getTime() + 1000 * 60 * 60 * 4); // добавляем часовой пояс +4 по Екб для GMT 0000

//--
//  --Функуия создания меню скриптов для запуска--
//--
function onOpen () {
  let ui = SpreadsheetApp.getUi();
   ui.createMenu("Скрипт")
  .addItem("1. Получить отчет из RM", "ParserSupportSC")
  .addItem("2. ???Обработка данных из 1С", "DataFrom1C")
  .addItem("3. Отправка информации", "EmailSend")
  .addToUi();  
}

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


//--
//  --PUT запрос xml, обновление данных в саппорт--
//--
function PutAPIRequest (urlIssues, payloadToIssue) {

  let optPut = {
    'method'  : 'put',
    'headers' : authHead,
    'contentType' : 'application/xml; charset=utf-8',
    'payload' : payloadToIssue,
    'muteHttpExceptions' : true    
  };

  UrlFetchApp.fetch(urlIssues, optPut);  
  
} // конец функции PutAPIRequest

//
//=================РАЗДЕЛ ВСПОМОГАТЕЛЬНЫХ ФУНКЦИЙ
//

//  --Перевод миллисекунд в дни
function MsecToDay(millisecons) {
  return Math.ceil(millisecons / 1000 / 60 / 60 / 24)
}

//  --перевод в одномерный массив
function Flatten(arrayOfArrays) {
  return [].concat.apply([], arrayOfArrays);
}

function ProcessingXMLDataNameNoteProject(elementRoot) {
  let xmlParserIssueCreatedOn = elementRoot.getChild('created_on').getText();
  let xmlParserIssueProject = elementRoot.getChild('project').getAttribute('id').getValue();
  let xmlParserIssueName = '--none--';
  let xmlParserIssueNote = 'Комментарии отсутсвуют';

  let xmlParserIssueJournals = elementRoot.getChild('journals');
  let xmlParserIssueJournal = xmlParserIssueJournals.getChildren('journal')

  xmlParserIssueJournal.forEach(xmlParserIssueJournalEach => {
    let issueJournalUserId = parseInt(xmlParserIssueJournalEach.getChild('user').getAttribute('id').getValue());
    if (xmlParserIssueJournalEach.getChild('notes').getText() != '' && Flatten(globalArrayEngineers).includes(issueJournalUserId) == true) {
      xmlParserIssueCreatedOn = xmlParserIssueJournalEach.getChild('created_on').getText();
      xmlParserIssueName = xmlParserIssueJournalEach.getChild('user').getAttribute('name').getValue();
      xmlParserIssueNote = xmlParserIssueJournalEach.getChild('notes').getText();
    }
  });
  return [xmlParserIssueCreatedOn, xmlParserIssueName, xmlParserIssueNote, xmlParserIssueProject]
  // Индексы - Дата последнего комментария ISO, ФИО Пользователя, Текст комментария, Название проекта RM
}


//  --Создание XML для передачи в RM
function MakeXMLIssueAndSendToSupport(idIssue, textIssue) {

  let urlSupportIssue = 'https://support.redmine.ru/issues/' + idIssue + '.xml'; // ссыдка на обращение по API XML
  // заменить ссылку на портал

  // Создание XML файла
  let rootXML = XmlService.createElement('issue');
  let childXMLnotes = XmlService.createElement('notes').setText(textIssue); // вставляет комментарий в обращение
  rootXML.addContent(childXMLnotes);
  let documentXML = XmlService.createDocument(rootXML);
  let contentXML = XmlService.getPrettyFormat().format(documentXML);

  PutAPIRequest(urlSupportIssue, contentXML) // Запускать осторожно, добавляет текст в обращения на проде!!!!
}


//  --Функция получения текущей даты вывод будет в виде 30.12.2021
function GetCurrentDay () {
  let currentDay = new Date (); 
  let dd = String(currentDay .getDate()).padStart(2, '0');
  let mm = String(currentDay .getMonth() + 1).padStart(2, '0'); //January is 0!
  let yyyy = currentDay .getFullYear();
  currentDay = dd + '.' + mm + '.' + yyyy;
  //Logger.log (currentDay);
  return currentDay;

}

//  --Функция разделения поля subject и возврачает номер обращения и контрагента
function GetSubjectNumberAndKA (subjectFull) {
  let subSplitNum = subjectFull.split(' ')[0];
  let subNumEnd = subSplitNum.replace('№', ''); //Убирает символ №
  let subSplitaKA =  subjectFull.substring(11); //Костыль который убирает 11 символов в начале, и остаются только данные по контрагенту

  return [subNumEnd, subSplitaKA]
  // Индекс - Номер обращения, Контрагент
}

//  --Функция редактирования даты
function GetCurrentDataTime (dataTime) {
  dataTimeCurrent = dataTime.replace('Z', '').replace('T', ' / ');
  return dataTimeCurrent
}

//  --Функция парсинга description и возвращения "Всего товаров"
function GetDescTotalGoods (textDesc) {
  textDescSplit = textDesc.split(" ");
  findIndexCollapse = textDescSplit.indexOf('collapse');
  findIndexText = textDescSplit.indexOf('товаров:');
  totalGoogs = textDescSplit[findIndexText + 1];
  return totalGoogs
}

//  --Функция возвращает массив используя сепаратор ; + пробел  
function SplitNagTag (splitValue) {
  let arrSplitValue = splitValue.split('; ');
  return arrSplitValue
}


//
//=================РАЗДЕЛ ФУНКЦИЙ ОТЧЕТА 1. Получить отчет из RM
//

function ParserSupportSC() {

  let allIssues = []; // массив всех обращений, начинается с пустого
  let actualIssues = []; // массив выборки обращений обращений, начинается с пустого

  // Начало получение массива обращений
  let beginXMLSuppurt = GetAPIRequest('https://support.redmine.ru/issues.xml?project_id=166'); //Авторизовывает в системе и возвращает страницу, Test
  // заменить ссылку на портал
  
  let xmlFirstPageXML = XmlService.parse(beginXMLSuppurt);
  let xmlFirstPageRoot = xmlFirstPageXML.getRootElement();
  let xmlFirstPageTotal = xmlFirstPageRoot.getAttribute('total_count').getValue();  // Получение общего числа обращений в проекте "Сервисный центр"

  let valueTotalPage = Math.ceil(xmlFirstPageTotal/25) //расчет количества страниц

  //Logger.log(valueTotalPage)

  for ( let i = 1; i<=valueTotalPage; i++) {
    let xmlParserPage = GetAPIRequest('https://support.redmine.ru/issues.xml?project_id=166&page=' + i); // заменить ссылку на портал
    let xmlParserPageXML = XmlService.parse(xmlParserPage);
    let xmlParserPageRoot = xmlParserPageXML.getRootElement();
    let xmlParserPageIssues = xmlParserPageRoot.getChildren('issue');

    xmlParserPageIssues.forEach(xmlParserPageIssue => {
      let idIssueHead = xmlParserPageIssue.getChild('id').getText(); 
      allIssues.push(idIssueHead);
    });
  
  }
  Logger.log(allIssues)
  Logger.log(allIssues.length)
  // конец получение массива обращений

  //ПОД РАЗДЕЛ ФОРМИРОВАНИЕ ОТЧЕТА
  globalNumberScript++; // Добавление счечика отчета
  globalRangeService.getCell(1, 2).setValue(globalNumberScript);

  //создание нового листа с номером и датой парсинга
  let newReportSheet = globalTable.getSheetByName(globalNumberScript + ' / ' + GetCurrentDay()); // Создаем переменную с названием листа
  if (newReportSheet != null) {                                // провека есть ли такой листа в таблице, если есть выводит Sheet, если нет null
    globalTable.deleteSheet(newReportSheet);                   // удаление листа если такой есть
  }
  newReportSheet = globalTable.insertSheet();                  // Создаем новый лист
  newReportSheet.setName(globalNumberScript + ' / ' + GetCurrentDay()); // установим имя этому листу

  let newReportRange = newReportSheet.getRange("A1:ZZ1000"); //Определение используемых ячеек

  // Создание наименований столбцов
  let nameColumsParser = ["RM СЦ", "1С СЦ", "Контрагент", "Дата создания", "Дата п.коммента", "Кто вносил", "Ед. товара", "RM NAGTAG", "Дата п.коммита", "Автор коммита", "Ответ в СЦ", "Ответ клиенту"] 
  let nameColumsParserItem = nameColumsParser.length // подсчет кол-ва в массиве
  for ( let i = 0; i<=nameColumsParserItem; i++) {
    newReportRange.getCell(1,i+1).setValue(nameColumsParser[i]);   // Записать данные в ячейку
  }

  //Начальная ячейка для ввода данных по парсингу всех обращений
  let lineBegin = 2;
  let columnBegin = 1;
  let cellActionSet = newReportRange.getCell(lineBegin,columnBegin);

  //Парсер XML redMine
  //for ( let i = 0; i<actualIssues.length; i++) {
  for ( let i = 0; i<allIssues.length; i++) {
    Logger.log('Парсинг обращения: ' + allIssues[i])
    let xmlRaportIssue = GetAPIRequest('https://support.redmine.ru/issues/' + allIssues[i] + '.xml?include=journals,relations'); // заменить ссылку на портал
    let xmlRaportIssueXML = XmlService.parse(xmlRaportIssue);
    let xmlParserIssueRoot = xmlRaportIssueXML.getRootElement();

    let differenceTime = new Date(now.getTime()) - new Date(ProcessingXMLDataNameNoteProject(xmlParserIssueRoot)[0]); // Вычитание дат, возвращает значение в миллдисекунлдах
    if ( MsecToDay(differenceTime) >= fuckedDays) {
      actualIssues.push(allIssues[i]);
      Logger.log('Обращение вышло за пределы интервала: ' + allIssues[i])
      //Получение ID обращения
      let xmlRaportIssueId = xmlParserIssueRoot.getChild('id').getText();  // полученние номер обращения RM
        xmlRaportIssueIdRichText = SpreadsheetApp.newRichTextValue()    // Какая-то неведомая хня, которая вставляет в ячейку текст со ссылкой
                    .setText(xmlRaportIssueId)
                    .setLinkUrl('https://support.redmine.ru/issues/' + xmlRaportIssueId)
                    .build();
      cellActionSet.setRichTextValue(xmlRaportIssueIdRichText);
      cellActionSet = newReportRange.getCell(lineBegin, ++columnBegin);

      //Разделяет поля subject и Записывает в ячейку номер обращения
      cellActionSet.setNumberFormat('@STRING@');
      let xmlRaportIssueSubjectNum = GetSubjectNumberAndKA(xmlParserIssueRoot.getChild('subject').getText())[0];
      cellActionSet.setValue(xmlRaportIssueSubjectNum); 
      cellActionSet = newReportRange.getCell(lineBegin, ++columnBegin);

      //Разделяет поля subject и Записывает в ячейку название контрагента
      let xmlRaportIssueSubjectKA = GetSubjectNumberAndKA(xmlParserIssueRoot.getChild('subject').getText())[1];
      cellActionSet.setValue(xmlRaportIssueSubjectKA); 
      cellActionSet = newReportRange.getCell(lineBegin, ++columnBegin);

      //Получение даты создания обращения
      let xmlRaportIssueCreated = GetCurrentDataTime(xmlParserIssueRoot.getChild('created_on').getText());
      cellActionSet.setValue(xmlRaportIssueCreated); 
      cellActionSet = newReportRange.getCell(lineBegin, ++columnBegin);  

      //Получение даты последнего комментария сотрудника СЦ
      let xmlRaportIssueUpdateData = GetCurrentDataTime(ProcessingXMLDataNameNoteProject(xmlParserIssueRoot)[0]);
      cellActionSet.setValue(xmlRaportIssueUpdateData); 
      cellActionSet = newReportRange.getCell(lineBegin, ++columnBegin);  

      //Получение даты последнего комментария и текст сотрудника СЦ
      let xmlRaportIssueUpdateName = ProcessingXMLDataNameNoteProject(xmlParserIssueRoot)[1];
      let xmlRaportIssueUpdateNote = ProcessingXMLDataNameNoteProject(xmlParserIssueRoot)[2];
      cellActionSet.setValue(xmlRaportIssueUpdateName);
      cellActionSet.setNote(xmlRaportIssueUpdateNote);   
      cellActionSet = newReportRange.getCell(lineBegin, ++columnBegin);

      // Вывод количества товара с обработка текста из description 
      let xmlRaportIssueQuantity = GetDescTotalGoods(xmlParserIssueRoot.getChild('description').getText());
      cellActionSet.setValue(xmlRaportIssueQuantity);
      cellActionSet.setHorizontalAlignment("center").setVerticalAlignment("middle");
      cellActionSet = newReportRange.getCell(lineBegin, ++columnBegin);

      //Вывод связных обращений из NAGTAG
      let arrRaportRelations = []; // пустой массив для добавления в него связных обращений
      let xmlRaportIssueRelations = xmlParserIssueRoot.getChild('relations');
      if (xmlRaportIssueRelations != null) {
        let xmlRaportIssueRelation = xmlRaportIssueRelations.getChildren('relation');
        xmlRaportIssueRelation.forEach(xmlRaportIssueRelationEach => {
          arrRaportRelations.push(xmlRaportIssueRelationEach.getAttribute('issue_to_id').getValue());
        } 
        )
        cellActionSet.setNumberFormat('@STRING@');
        cellActionSet.setValue(arrRaportRelations.join('; ')); 
        cellActionSet = newReportRange.getCell(lineBegin, ++columnBegin);

        //-- получение данных из обращения arrRaportRelations[0]
        let xmlRaportIssueNagtag = GetAPIRequest('https://support.redmine.ru/issues/' + arrRaportRelations[0] + '.xml?include=journals,relations');
        let xmlRaportIssueXMLNagtag = XmlService.parse(xmlRaportIssueNagtag);
        let xmlParserIssueRootNagtag = xmlRaportIssueXMLNagtag.getRootElement();

        //Вывод Даты последнего комментария из NAG TAG
        let xmlRaportIssueNagtagCreated = GetCurrentDataTime(ProcessingXMLDataNameNoteProject(xmlParserIssueRootNagtag)[0]);
        cellActionSet.setValue(xmlRaportIssueNagtagCreated); 
        cellActionSet = newReportRange.getCell(lineBegin, ++columnBegin);

        //Вывод Автора последнего комментария из NAG TAG
        let xmlRaportIssueNagtagName = ProcessingXMLDataNameNoteProject(xmlParserIssueRootNagtag)[1];
        let xmlRaportIssueNagtagNote = ProcessingXMLDataNameNoteProject(xmlParserIssueRootNagtag)[2];
        cellActionSet.setValue(xmlRaportIssueNagtagName);
        cellActionSet.setNote(xmlRaportIssueNagtagNote);   
        cellActionSet = newReportRange.getCell(lineBegin, ++columnBegin); 
      }

      columnBegin = 1; //Вернуть значение начальной колонки в первую
      cellActionSet = newReportRange.getCell(++lineBegin,columnBegin); //Перезапись координат со сдвигом на строчку ниже
    }
  }

  Logger.log(actualIssues)
  Logger.log(actualIssues.length)



  for (let j = 1; j <= 10; j++) {
    newReportSheet.autoResizeColumn(j);
  }

  Browser.msgBox("Отчет составлен")
}



//Logger.log(xxx);


//
//=================РАЗДЕЛ ФУНКЦИЙ ОТЧЕТА 2. Обработка данных из 1С
//

function DataFrom1C(){ 



}

//
//=================РАЗДЕЛ ФУНКЦИЙ ОТЧЕТА 3. Отправка сообщений в RM и в почту
//

function EmailSend(){ 

  let lastCellActive = globalSheetActive.getLastRow(); // поределяем последнию строчку
  Logger.log(lastCellActive);

  let messageToEMailES = 'Коллеги, добрый день.<br>Прошу оповестить клиентов о длительной обработке оборудования в СЦ.<br>';
  let emailAddressES = 'name1@domen.ru' + ',' + 'name2@domen.ru' + ',' + 'name3@domen.ru'; // адресаты почты
  //let emailAddressES = 'name4@domen.ru'; // адресаты почты
  //let emailCopyAddressES = 'name5@domen.ru'; // адресаты в копии
  let subjectES = 'DIAG Отчет обращений без движения более ' + fuckedDays + ' дней от ' + GetCurrentDay(); // Тема письма

  for ( let i = 2; i<=lastCellActive; i++) {
    let replyIssuesIdSC = globalRangeActive.getCell(i, 1).getValue(); // получаем номера проекта СЦ
    let replyRow1CTag = globalRangeActive.getCell(i, 2).getValue(); // получаем номер 1С
    let replyRowKATag = globalRangeActive.getCell(i, 3).getValue(); // получаем Контрагента
    let replyIssuesIdTagArray = globalRangeActive.getCell(i, 8).getValue(); // получаем номера NAGTAG
    let replyIssuesTextSC = globalRangeActive.getCell(i, 11).getValue(); // получаем ответ для проекта СЦ
    let replyIssuesTextNagtag = globalRangeActive.getCell(i, 12).getValue(); // получаем ответ для КЛИЕНТА
    MakeXMLIssueAndSendToSupport(replyIssuesIdSC, replyIssuesTextSC); // !!!! ОТКРЫЛ, Добавляет информацию в ПРОД !!!!

    
    
    if (replyIssuesIdTagArray != '') {  // Если есть данные в ячейке NAGTAG то мы добавляем текст в обращение
      let splitReplyIssuesIdTagArray = SplitNagTag(replyIssuesIdTagArray); // разделение сторки на индексы
      let splitReplyIssuesTextTagArray = SplitNagTag(replyIssuesTextNagtag); // разделение сторки на индексы
      for ( let j = 0; j<splitReplyIssuesIdTagArray.length; j++) {
        //Logger.log(splitReplyIssuesIdTagArray[j] + ' / ' + splitReplyIssuesTextTagArray[j]);
        if (splitReplyIssuesTextTagArray[j] != '' && splitReplyIssuesTextTagArray[j] != null) {  // Если в индексе отсутсвует описание то не добавлять в SUPPORT
          MakeXMLIssueAndSendToSupport(splitReplyIssuesIdTagArray[j], splitReplyIssuesTextTagArray[j]); // !!!! ОТКРЫЛ, Добавляет информацию в ПРОД !!!!
        }
      }
    // It's WORK!!! Нужно проверить отладку на проде.
    
    }
    else {  // Если пустая ячейка NAGTAG то формируем письмо на email
      messageToEMailES += '<br>' + 'Тикет СЦ: ' + '<b>' + replyIssuesIdSC + '</b>' + ' / Обращение в 1C: ' + '<b>' + replyRow1CTag + '</b>' + ' / Контрагент: ' + '<b>' + replyRowKATag + '</b>' + '<br>';
      if (replyIssuesTextSC != '') { messageToEMailES += '<p style="margin-left: 50px;">' + 'Информация для внутреннего использования: ' + '<br>' + '<b><i>' + replyIssuesTextSC + '</i></b>'; }
      else { messageToEMailES += '<p style="margin-left: 50px;">' + 'Информация для внутреннего использования:' + '<br>' + '<b><i>' + 'Информацию не внесли в таблицу' + '</i></b>'}
      if (replyIssuesTextNagtag != '') { messageToEMailES += '<p style="margin-left: 50px;">' + 'Информация для передачи КЛИЕНТУ: ' + '<br>' + '<b><i>' + replyIssuesTextNagtag  + '</i></b>' + '<br>'; }
      else { messageToEMailES += '<p style="margin-left: 50px;">' + 'Информация для передачи КЛИЕНТУ:' + '<br>' + '<b><i>' + 'Информацию не внесли в таблицу' + '</i></b>'}
      messageToEMailES += '<p style="margin-left: 0px;">' + '<br>';
     }

  }
  
  MailApp.sendEmail({
  to: emailAddressES,
  //cc: emailCopyAddressES,
  subject: subjectES,
  htmlBody: messageToEMailES
  });


}



//
//=================ПЕСОЧНИЦА
//

function Testing(){ 


}







