const filesCreator = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
var filesCount = filesCreator.getLastRow();

const dataBase = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
var dataBaseLength = dataBase.getLastRow();


const templateFolder = DriveApp.getFoldersByName("templates").next();
var templates = templateFolder.getFiles();

var templateNames = [];
while (templates.hasNext()) {
  var templateName = templates.next();
  templateNames.push(templateName)
}

const result = DriveApp.getFoldersByName("result").next();


function collectData() {
   for (var i = 2; i <= filesCount; i++) {
      var fileName = filesCreator.getRange(i, 1).getValue();
      var thing = filesCreator.getRange(i, 2).getValue();
      var summ = filesCreator.getRange(i, 3).getValue();
      var link = filesCreator.getRange(i, 4).getValue();
      if (link == '') {
        for (var j = 2; j <= dataBaseLength; j++) {
          var dbFileName = filesCreator.getRange(j, 1).getValue();
          if (dbFileName == fileName) {
            var nomination = dataBase.getRange(j, 1).getValue();
            var lastName = dataBase.getRange(j, 2).getValue();
            var name = dataBase.getRange(j, 3).getValue();
            var midName = dataBase.getRange(j, 4).getValue();
            var job = dataBase.getRange(j, 5).getValue();
            var base = dataBase.getRange(j, 6).getValue();
            var address = dataBase.getRange(j, 7).getValue();
            var phoneNumbeer = dataBase.getRange(i, 8).getValue();
            var fax = dataBase.getRange(j, 9).getValue();
            var email = dataBase.getRange(j, 10).getValue();
            var bank = dataBase.getRange(j, 11).getValue();
            var rs = dataBase.getRange(j, 12).getValue();
            var bill = dataBase.getRange(j, 13).getValue();
            var bik = dataBase.getRange(j, 14).getValue();
            var inn = dataBase.getRange(j, 15).getValue();
            var kpp = dataBase.getRange(j, 16).getValue();
            var result = [[nomination, lastName, name, midName, job, base, address, phoneNumbeer, fax, email, bank, rs, bill, bik, inn, kpp], fileName, thing, summ];
            return result;
          }
        }
      }
   }
}

templateNames.forEach(function (templateName) {


   this[templateName+"PDF"] = function () {
      var data = collectData();
      let PDFfolderID = result.getFoldersByName('PDF').next().getId()
      let PDFfolder = DriveApp.getFolderById(PDFfolderID);
      let templateFileID = templateFolder.getFilesByName(templateName).next().getId();
      let templateFile = DriveApp.getFileById(templateFileID);

      const tempFile = templateFile.makeCopy(PDFfolder);
      const tempDocFile = DocumentApp.openById(tempFile.getId());
      const body = tempDocFile.getBody();
      body.replaceText("{Наименование}", data[0][0]);
      body.replaceText("{Фамилия}", data[0][1]);
      body.replaceText("{Имя}", data[0][2]);
      body.replaceText("{Отчество}", data[0][3]);
      body.replaceText("{Должность}", data[0][4]);
      body.replaceText("{Основание}", data[0][5]);
      body.replaceText("{Юр. Адрес}", data[0][6]);
      body.replaceText("{Телефон}", data[0][7]);
      body.replaceText("{Факс}", data[0][8]);
      body.replaceText("{Эл. Почта}", data[0][9]);
      body.replaceText("{Банк}", data[0][10]);
      body.replaceText("{Рс}", data[0][11]);
      body.replaceText("{Кор. Счет}", data[0][12]);
      body.replaceText("{БИК}", data[0][13]);
      body.replaceText("{ИНН}", data[0][14]);
      body.replaceText("{КПП}", data[0][15]);
      body.replaceText("{Предмет}", data[2]);
      body.replaceText("{Сумма}", data[3]);
      
    tempDocFile.saveAndClose();
    var date = Utilities.formatDate(new Date(), "GMT+3", "dd/MM/yyyy_HH:mm")
    const pdfFile = tempDocFile.getAs(MimeType.PDF);
    let newName = PDFfolder.createFile(pdfFile).setName(data[1]+"_"+date);
    PDFfolder.removeFile(tempFile);
    let pdfFileLink = PDFfolder.getFilesByName(newName).next().getUrl();
    for (var i = 2; i <= filesCount; i++) {
      if(filesCreator.getRange(i, 1).getValue() == data[1]) {
        filesCreator.getRange(i, 4).setValue(pdfFileLink);
      }
    }
  }


  this[templateName+"GDOC"] = function () {
    
      var data = collectData();
      let GDOCfolderID = result.getFoldersByName('GDOC').next().getId()
      let GDOCfolder = DriveApp.getFolderById(GDOCfolderID);
      let templateFileID = templateFolder.getFilesByName(templateName).next().getId();
      let templateFile = DriveApp.getFileById(templateFileID);

      const tempFile = templateFile.makeCopy(GDOCfolder);
      const tempDocFile = DocumentApp.openById(tempFile.getId());
      const body = tempDocFile.getBody();
      body.replaceText("{Наименование}", data[0][0]);
      body.replaceText("{Фамилия}", data[0][1]);
      body.replaceText("{Имя}", data[0][2]);
      body.replaceText("{Отчество}", data[0][3]);
      body.replaceText("{Должность}", data[0][4]);
      body.replaceText("{Основание}", data[0][5]);
      body.replaceText("{Юр. Адрес}", data[0][6]);
      body.replaceText("{Телефон}", data[0][7]);
      body.replaceText("{Факс}", data[0][8]);
      body.replaceText("{Эл. Почта}", data[0][9]);
      body.replaceText("{Банк}", data[0][10]);
      body.replaceText("{Рс}", data[0][11]);
      body.replaceText("{Кор. Счет}", data[0][12]);
      body.replaceText("{БИК}", data[0][13]);
      body.replaceText("{ИНН}", data[0][14]);
      body.replaceText("{КПП}", data[0][15]);
      body.replaceText("{Предмет}", data[2]);
      body.replaceText("{Сумма}", data[3]);
      
    tempDocFile.saveAndClose();
    var date = Utilities.formatDate(new Date(), "GMT+3", "dd/MM/yyyy_HH:mm")
    let newName = tempFile.setName(data[1]+"_"+date);
    let pdfFileLink = GDOCfolder.getFilesByName(newName).next().getUrl();
    for (var i = 2; i <= filesCount; i++) {
      if(filesCreator.getRange(i, 1).getValue() == data[1]) {
        filesCreator.getRange(i, 4).setValue(pdfFileLink);
      }
    }
  }


  this[templateName+"WORD"] = function () {
    
      var data = collectData();
      let WORDfolderID = result.getFoldersByName('WORD').next().getId()
      let WORDfolder = DriveApp.getFolderById(WORDfolderID);
      let templateFileID = templateFolder.getFilesByName(templateName).next().getId();
      let templateFile = DriveApp.getFileById(templateFileID);

      const tempFile = templateFile.makeCopy(WORDfolder);
      const tempDocFile = DocumentApp.openById(tempFile.getId());
      const body = tempDocFile.getBody();
      body.replaceText("{Наименование}", data[0][0]);
      body.replaceText("{Фамилия}", data[0][1]);
      body.replaceText("{Имя}", data[0][2]);
      body.replaceText("{Отчество}", data[0][3]);
      body.replaceText("{Должность}", data[0][4]);
      body.replaceText("{Основание}", data[0][5]);
      body.replaceText("{Юр. Адрес}", data[0][6]);
      body.replaceText("{Телефон}", data[0][7]);
      body.replaceText("{Факс}", data[0][8]);
      body.replaceText("{Эл. Почта}", data[0][9]);
      body.replaceText("{Банк}", data[0][10]);
      body.replaceText("{Рс}", data[0][11]);
      body.replaceText("{Кор. Счет}", data[0][12]);
      body.replaceText("{БИК}", data[0][13]);
      body.replaceText("{ИНН}", data[0][14]);
      body.replaceText("{КПП}", data[0][15]);
      body.replaceText("{Предмет}", data[2]);
      body.replaceText("{Сумма}", data[3]);
      
    tempDocFile.saveAndClose();
    var date = Utilities.formatDate(new Date(), "GMT+3", "dd/MM/yyyy_HH:mm")
    let newName = tempFile.setName(data[1]+"_"+date);
    let pdfFileLink = WORDfolder.getFilesByName(newName).next().getUrl();
    for (var i = 2; i <= filesCount; i++) {
      if(filesCreator.getRange(i, 1).getValue() == data[1]) {
        filesCreator.getRange(i, 4).setValue(pdfFileLink);
      }
    }
  }


});


function createMenu() {
  var ui =  SpreadsheetApp.getUi();
  var UI = ui.createMenu('Меню')

  templateNames.forEach(function (templateName) {
    UI.addSubMenu(
      SpreadsheetApp.getUi().createMenu(templateName.toString())
      .addItem('PDF',templateName+'PDF')
      .addItem('GDOC',templateName+'GDOC')
      .addItem('WORD',templateName+'WORD')
    ).addSeparator();
    UI.addToUi();
  });
  UI.addItem('🗘 Обновить меню', 'createMenu').addToUi();
}

function onOpen() {
  createMenu();
} 

