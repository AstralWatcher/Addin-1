
//Verzija 5

var darkBlue = "#2d5898";
var lightBlue = "#add8e6";
var fillYellow = "#fff200";
var fillGreen = "#247348";
var fillRed = "#cc0000";
var lightTextColor = "white";
var textAlignmentH = "center";
var darkTextColor = "black";
var moreRunes = false;

function doGet(){
  var html = HtmlService.createHtmlOutputFromFile('home').setTitle('Prozor generator rasporeda');
  return html;
}

function onInstall(e) { //SpreadsheetApp.flush(); za promene da se insta odraze
  onOpen(e);
}



function scrollTo(num){

  var first = -1;
  var second = -1;
  var spread = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spread.getActiveSheet();
  var sheetName = sheet.getName();
  if(sheetName == "Settings") 
    throw "Nije izabran list sa rasporedom";
  var range = sheet.getRange(3, 2, 1,sheet.getDataRange().getLastColumn()-1);
  var values = range.getValues()[0];
  var rac = 0;
  for(var i = 0; i< values.length; ++i){
   
    
    if(values[i] != "" && values[i])
    {
            
      if(first == -1)
      {
        first = i;
      }
      else if(first !=-1 && second == -1)
      {
        second = i;
        //break;
      }
   
      ++rac;
      Logger.log("values =" + values[i] + " NUM=" + num)
      var date = new Date(values[i+second-first]);
      var today = new Date(values[i]);
      if(today.getDate() >= parseInt(num,10) || date.getDate() >= parseInt(num,10))
      {
    
        break;
      }

    }   
  }
  
  Logger.log("First" + first);
  Logger.log("Second" + second);
  var dodatno=0;
  var dani = second-first;
  Logger.log("DANI=" + dani);
  var active = sheet.getActiveRange();
  if(parseInt(num,10) == 1)
    dani = 2;
  if(parseInt(num,10) > 2)
    ++rac;

  sheet.setActiveRange(sheet.getRange(3, sheet.getDataRange().getLastColumn(), 1, 1));
  SpreadsheetApp.flush();
  sheet.setActiveRange(sheet.getRange(3, dani*rac, 1, 1));
 
}
                               


function getAvailableTags() {
 var availableTags = getTableData("Settings","Radnici",0,true);
 return( availableTags );
}

function findUserRow(aName,aLastName,aNickname){
  var i;
  for(i=0; i< aName.map(function(value,index) { return value[0]; }).length-1;++i){
    if(aName[i] == "" && aLastName[i]=="" && aNickname[i] == "")
    {
      return i;
    }
  }
  return i+1;
}
function checkIfExists(nickname,aNickname){
  Logger.log("Nickname =" + nickname);
  if(aNickname =="")
    return false;
  
  for(i=0; i< aNickname.map(function(value,index) { return value[0]; }).length-1;++i){
    if(aNickname[i]!="")
      if(aNickname[i] == nickname)
      {
        return true;
      }

  }
  return false;
}

function formatToName(ime){
  if(ime.length>1)
   return ime.charAt(0).toUpperCase() + ime.substring(1).toLowerCase().trim();
  else 
    return ime.charAt(0).toUpperCase();
}

function addUser(form){
  Logger.log(form);
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  if(form.name != undefined || form.lastName != undefined || form.nickname != undefined)
  {
    var user = new Object();
    var timestamp = "";
    var date = new Date();
    if(!form.name) user.name = ""; else user.name = form.name.charAt(0).toUpperCase() + form.name.substring(1).toLowerCase().trim();
    if(!form.lastName) user.lastName = ""; else user.lastName = form.lastName.charAt(0).toUpperCase() + form.lastName.substring(1).toLowerCase().trim();
    if(!form.nickname) user.nickname = ""; else user.nickname = form.nickname.charAt(0).toUpperCase() + form.nickname.substring(1).toLowerCase().trim();
    if(!form.beleske) user.beleske = ""; else {user.beleske = form.beleske + " ; ";   timestamp= date.getFullYear() +"." + Number(date.getMonth() +1) + "." + date.getDate() + ":"; }
    if(!form.contact) user.contact = ""; else user.contact = form.contact;
    if(!form.birthday) user.birthday = ""; else user.birthday = form.birthday;   
    var userArray = [[user.name,user.lastName,user.nickname,user.contact,timestamp +  user.beleske,user.birthday]];
        
    var imena = getTableData("Settings","Mušterije",0,false);
    var prezimena = getTableData("Settings","Mušterije",1,false);
    var nadimak = getTableData("Settings","Mušterije",2,false);

    Logger.log(prezimena);
    Logger.log(nadimak);
    
    if(checkIfExists(user.nickname, nadimak)){
     throw "Vec postoji sa tim nadimkom mušterija";
    }
    
    var location = findUserRow(imena,prezimena,nadimak);
    
    var textFinder = settingsSheet.createTextFinder("Mušterije");
    var firstOccurrence = textFinder.findNext();
    var slobodan = settingsSheet.getRange(firstOccurrence.getRowIndex() + location +2, firstOccurrence.getRow(), 1, 6);
    Logger.log(slobodan.getA1Notation());
    slobodan.setValues(userArray);  
 
  } else
    throw "Mora se popuniti: ime, prezime ili nadimak novom mušteriji";  
}

function hideDays(){
 togglePastDays(true); 
}
function showDays(){
 togglePastDays(false); 
}

function scroll(){
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Unesite dan koji zelite [0-31]:');
  if (response.getSelectedButton() == ui.Button.OK) {
    scrollTo(response.getResponseText());
  }   
}


function makeNotes(form){
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var user = findUser(form,false);
  var notes = settingsSheet.getRange(user.getRowIndex(), user.getColumn()+2, 1, 1);
  var date = new Date();
  var timestamp = date.getFullYear() +"." + (Number(date.getMonth())+ 1) + "." + date.getDate() + ":"
  notes.setValue(timestamp +  form.beleske + " ; " + notes.getValue());
}
//Vraca cell @nadimk korisnika u tabeli Musterije
// strict - false ako je dodavanje, true ako je pretraga
function findUser(form,strict){
  Logger.log("Primio u findUser");
  Logger.log(form);
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  if(form.nickname != undefined && form.nickname){
    Logger.log("Usao u nickname");
    var textFinder = settingsSheet.createTextFinder(formatToName(form.nickname));
    textFinder.matchCase(false);
    var firstOccurrence = textFinder.findNext();
    var rangeTest = settingsSheet.getRange(firstOccurrence.getRowIndex(), firstOccurrence.getColumn()+2, 1, 1);
    if(strict && rangeTest.getValue() == "")
      throw "Nema beleške da prikaze";
     Logger.log("vraca nickname");
    return firstOccurrence;
  }else if(form.name != undefined && form.lastName != undefined && form.name && form.lastName){
     Logger.log("Usao name&lname");
    var textFinder = settingsSheet.createTextFinder(formatToName(form.lastName));
     textFinder.matchCase(false);
    var firstOccurrence = textFinder.findNext();
    var areMore = true;
    while(areMore){
      if(firstOccurrence == undefined){
        areMore = false;
        throw "Nema korisnika pod tim parametrima";
      }
      if(formatToName(form.name) == formatToName(settingsSheet.getRange(firstOccurrence.getRowIndex(), firstOccurrence.getColumn()-1, 1, 1).getValue()))
      {
         Logger.log("Vraca n&lname->nn");
         return settingsSheet.getRange(firstOccurrence.getRowIndex(), firstOccurrence.getColumn()+1, 1, 1);
      }
       
      firstOccurrence = textFinder.findNext();         
    }
  }else 
    throw "Popunite ime i prezime ILI nadimak za osobu";
  return undefined;
}

function showHistory(form){
  Logger.log(form);
  
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var firstOccurrence = findUser(form,true);
  if(firstOccurrence == undefined) throw "Došlo do greške prilikom traženja istorije";
  
  if(form.nickname != undefined){
    var rangeTest = settingsSheet.getRange(firstOccurrence.getRowIndex(), firstOccurrence.getColumn()+2, 1, 1);
    var res = rangeTest.getValue().replace(/;/gi, "<br>");
    var htmlOutput = HtmlService
    .createHtmlOutput('<p>' + res + '</p>')
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Istorija');
  }else if(form.name != undefined && form.lastName != undefined){
    var rangeTest = settingsSheet.getRange(firstOccurrence.getRowIndex(), firstOccurrence.getColumn()+2, 1, 1);
    var res = rangeTest.getValue().replace(/;/gi, "<br>");
    var htmlOutput = HtmlService
    .createHtmlOutput('<p>' + res + '</p>')
    .setWidth(500)
    .setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Istorija');
  
  } 
}

function onOpen(e) {
  SpreadsheetApp.getUi()
  .createMenu("Raspored")
  .addSubMenu(SpreadsheetApp.getUi().createMenu('Prikaz')
              .addItem("Sakri", 'hideDays')
              .addItem("Prikaži", 'showDays'))
   .addSubMenu(SpreadsheetApp.getUi().createMenu('Navigacija')
              .addItem("Na dan", 'scroll'))
    .addSeparator()
  .addItem('Generator', 'showSidebar')
  .addToUi();
   SpreadsheetApp.getUi().createMenu("Mušterije").addItem("Istorija", 'showUserSidebar').addItem("Dodaj", 'showUserAddSidebar').addToUi();
}

function showUserSidebar(){
var html = HtmlService.createHtmlOutputFromFile('users').setTitle('Prozor za mušterije');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showUserAddSidebar(){
var html = HtmlService.createHtmlOutputFromFile('usersAdd').setTitle('Prozor Dodavanja mušterija');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('home').setTitle('Prozor generator rasporeda');
  SpreadsheetApp.getUi().showSidebar(html);
}

function testButton(form){
 Logger.log(form);
  var info = getValuesFromForm(form);
  Logger.log(info.vreme);
  
}

function makeHeaderForTable(range,text){
  range.setBackground(darkBlue)
  range.setHorizontalAlignment(textAlignmentH);
  range.setFontColor(lightTextColor);
  range.merge();
  range.setValue(text);
  //range.addDeveloperMetadata(text);
}
// getTableData("Settings","Radnici",0,true) Prva kolona
function getTableData(sheetName,tableTextHeader,column,filter){
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var textFinder = settingsSheet.createTextFinder(tableTextHeader);
  var firstOccurrence = textFinder.findNext();
  var rangeTest = settingsSheet.getRange(firstOccurrence.getRowIndex()+2, firstOccurrence.getColumn()+column, settingsSheet.getDataRange().getLastRow(), 1);
  var valuesRange = rangeTest.getValues();
  var result = valuesRange;
  if(filter == true)
     result = [].concat.apply([], valuesRange).filter(String); 
  return result;
}

function test(form){
  Logger.log("POCEO");
  var info = getValuesFromForm(form);
  Logger.log(info.radnici);
  Logger.log(info);
}



// date - info.date
// days = 0 za taj dan, 1 za sutraDan itd.
// praznici iz TableNeradniDani
function dateShouldBeSkipped(date,days,praznici){
  var check = new Date(date.getTime() + days * 60 * 24 * 60000);
  var danUNedelji = check.getDay();
  if (danUNedelji == 0) { //SKIP NEDELJA
    return true;
  }
  for (var i = 0; i < praznici.length; ++i) { //SKIP PRAZNIK
    if (check.getFullYear() == praznici[i].getFullYear() && check.getMonth() == praznici[i].getMonth() && check.getDate() == praznici[i].getDate()) {
      return true;
    }
  }
  return false;
}

function translate(date) {
  switch (date.getMonth() + 1) {
    case 1:
      return "JAN";
    case 2:
      return "FEB";
    case 3:
      return "MAR";
    case 4:
      return "APR";
    case 5:
      return "MAJ";
    case 6:
      return "JUN";
    case 7:
      return "JUL";
    case 8:
      return "AVG";
    case 9:
      return "SEP";
    case 10:
      return "OKT";
    case 11:
      return "NOV";
    case 12:
      return "DEC";
    default:
      return "ERROR";
  }
}

function getValuesFromForm(form){
  Logger.log("Form in getValuesFromForm dole");
  Logger.log(form);
  var info = new Object();
  if(form == undefined) 
  {
    form = new Object();
    form.start = "";
    form.date = "";
    form.name = "";
    form.end = "";
    form.int = "";
  } 
  Logger.log(form);
  
  if(!Array.isArray(form.radnik))
  {
   Logger.log("DODATO []");
   info.radnici = [form.radnik];
  }
  else
  {
    info.radnici = form.radnik; 
  }
  
  
  info.start = form.start; 
  info.date  = form.date;
  info.name = form.name;
  info.end = form.end;
  info.int = form.int;  

  if(form.vreme == undefined) info.vreme = false
  else info.vreme = true;
  if (info.start == "")  info.start  = "8";
  if (info.date  == "") 
  {
    info.date  = "2019-1-1";
    info.sdate = "2019-1-1";
  } else info.sdate = info.date;
  if (info.end == "" )  info.end = "20";
  if (info.int == "") info.int = "15";
  Logger.log("info.start=" + info.start);
  Logger.log("info.date="+ info.date);
  Logger.log("info.end="+ info.end);
  Logger.log("info.int="+ info.int);
  var splitDate = info.date.split("-");
  Logger.log(splitDate);
  var datumZahtevani;
  try {
    datumZahtevani = new Date(splitDate[0], splitDate[1] - 1, splitDate[2]);
  }    catch (error) {
    Logger.log("error u trazenom datumu");
  }
  Logger.log("Prosao datum");
  Logger.log(datumZahtevani);
  info.date = datumZahtevani;
  if (info.name == "") {
    var nazivDatuma = translate(datumZahtevani);
    info.name = splitDate[0] + nazivDatuma;
  } 
 
   Logger.log("Kraj ucitavanja imena forme="+ info.name);
  Logger.log("Kraj ucitavanja imena forme="+ info.date);
  return info;
}

function applyFormating(range,colorBackground,colorText){
  range.setBackground(colorBackground)
  range.setHorizontalAlignment(textAlignmentH);
  range.setFontColor(colorText);
}

function addDailyConditionalFormating(range){
   var rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(GT(INDIRECT("RC", FALSE), NOW() ) ,LTE(INDIRECT("RC", FALSE), NOW()+4 ))')
    .setBackground(fillYellow)
    .setFontColor("black")
    .setRanges([range])
    .build();

   var rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenDateEqualTo(SpreadsheetApp.RelativeDate.TODAY)
    .setBackground(fillGreen)
    .setFontColor("white")
    .setRanges([range])
    .build();
  
   var rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenDateBefore(SpreadsheetApp.RelativeDate.TODAY)
    .setBackground(fillRed)
    .setFontColor("white")
    .setRanges([range])
    .build();
  
  var rules = range.getSheet().getConditionalFormatRules();
  
  rules.push(rule1);
  rules.push(rule2);
  rules.push(rule3);
  range.getSheet().setConditionalFormatRules(rules);  
}


function addMontlyConditionalFormating(range){
   var rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(GTE(INDIRECT("RC", FALSE), EDATE(NOW(),1)-DAY(EDATE(NOW(),1))+1 ) ,LTE(INDIRECT("RC", FALSE), DATE(YEAR(EDATE(NOW(),1)),MONTH(EDATE(NOW(),1))+1,1)-1))')
    .setBackground(fillYellow)
    .setFontColor("black")
    .setRanges([range])
    .build();
  
   var rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(GTE(INDIRECT("RC", FALSE), NOW()-DAY(NOW()) ), LTE(INDIRECT("RC", FALSE), DATE(YEAR(NOW()),MONTH(NOW())+1,1)-1))')
    .setBackground(fillGreen)
    .setFontColor("white")
    .setRanges([range])
    .build();
  
   var rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=LT( INDIRECT("RC",false) ,NOW()-DAY(NOW()))')
    .setBackground(fillRed)
    .setFontColor("white")
    .setRanges([range])
    .build();
  
  var rules = range.getSheet().getConditionalFormatRules();
  rules.push(rule1);
  rules.push(rule2);
  rules.push(rule3);
  range.getSheet().setConditionalFormatRules(rules);  
}

function makeTimeTable(currentWorksheet,info,rangeTimeTablePrvi,rangeTimeTable){
  
    
  rangeTimeTablePrvi.setValue("=$D$1");
  rangeTimeTablePrvi.setNumberFormat("hh:mm");
  applyFormating(rangeTimeTablePrvi,darkBlue,lightTextColor);
  
  rangeTimeTable.setFormula('=IF(A6="","",IF(A6+TIME(0,$G$1,0)<=$I$1,A6+TIME(0,$G$1,0),""))');

  rangeTimeTable.setNumberFormat("hh:mm");
  applyFormating(rangeTimeTable,darkBlue,lightTextColor);
  rangeTimeTable.setBorder(true, false, true, false, null, false, darkBlue, SpreadsheetApp.BorderStyle.SOLID);
  
  
  addMontlyConditionalFormating(currentWorksheet.getRange("B1"));
  addMontlyConditionalFormating(currentWorksheet.getRange("A2"));
}

function calculateSlots(info){
  var od = new Date();
  od.setHours(info.start);
  od.setMinutes(0);
  od.setSeconds(1);
  var dok = new Date();
  dok.setHours(info.end);
  dok.setMinutes(0);
  dok.setSeconds(0);
    
  var it = 0;
  for (var i = od; i < dok; ) {
    it++;
    i = new Date(i.getTime() + info.int * 60000);
  }
  return it;  
}

function makeHours(currentWorksheet,info){
  var range = currentWorksheet.getRange("A1:I1");
 
  var rangeValues = [["DATUM:","=A2" , "POCETAK:", "=TIME(" + info.start + ",0,0)", "INTERVAL:", null, info.int , "KRAJ:", "=TIME(" + info.end + ",0,0)"]]
  
  var rangeLeftHead = currentWorksheet.getRange(2, 1, 4, 1);
  Logger.log(">>>>>>>>>>>>>>>>>>>>>" + info.date);
  Logger.log(">>>>>>>>>>>>>>>>>>>>> getFullYear=" + info.date.getFullYear() +" getMonth()=" + info.date.getMonth()+1 + " , getDate() = " + info.date.getDate());
  var dateToWrite = info.sdate.split("-");
  var rangeLeftHeadValues = [["=DATE(" + dateToWrite[0] + "," + dateToWrite[1] + "," + dateToWrite[2] + ")"],[null],[null],["Vreme"]];
  
  
  range.setValues(rangeValues);
  range.setNumberFormats([["", "dd.mm.yyyy", "", "hh:mm", "", "", "00","", "hh:mm"]]);
  rangeLeftHead.setNumberFormats([["dd.mm.yyyy"],[""],[""],[""]]);
  applyFormating(range,darkBlue,lightTextColor);
  
  
  rangeLeftHead.setValues(rangeLeftHeadValues);
  applyFormating(rangeLeftHead,darkBlue,lightTextColor);
  
  var it = calculateSlots(info);

  Logger.log("IT CALCULATED " + it);
  
  var rangeTimeTablePrvi = currentWorksheet.getRange("A6");
  var rangeTimeTable = currentWorksheet.getRange(rangeTimeTablePrvi.getRow() + 1, 1, it-1, 1);

  makeTimeTable(currentWorksheet,info,rangeTimeTablePrvi,rangeTimeTable);
  
  return it;
}




function daysInMonth(date) {
  var month = date.getMonth() + 1;
  var year = date.getFullYear();
  return new Date(year, month, 0).getDate();
}

function makePeople(currentWorksheet,info,hoursCount){
  if(info.radnici == undefined || info.radnici == "")
    throw "Nema radnika izabrano";
  if(info.radnici.length == 0)
    throw "Nema radnika izabrano";
  //info.radnici = ["Bojan","Sofija"];
  //info.vreme = true;
  
  
  var columnDataRadnici =  getTableData("Settings","Radnici",0,true);
  var columnDataBojaRadnika = getTableData("Settings","Radnici",1,true);
  var neradniDani =  getTableData("Settings","Neradni Dani",0,true);
  
  var edays = daysInMonth(info.date);
  
  var radniciArray = new Array(columnDataRadnici.length);
  var bojaRadnikaArray = new Array(columnDataBojaRadnika.length);
  
  var brojac = 0;
  columnDataRadnici.forEach(function (item) {
    radniciArray[brojac] = item;
    ++brojac;
  });
  brojac = 0;
  columnDataBojaRadnika.forEach(function (item) {
    bojaRadnikaArray[brojac] = item;
    ++brojac;
  });
  
  if(columnDataBojaRadnika.length != columnDataRadnici.length)
    throw "Nema za svakog radnika boja u tabeli";
  
  var imaDana = 0;
  
  var valuesPerPerson = [["Usluga", "Cena", "Klijent", "Beleške", "Pomocnik","Aktivnost"]];
  
  var kolicinaKolonaPoOsobi = 6; //TODO promenuti ako ima vise kolona  
  var danZauzima = info.radnici.length * kolicinaKolonaPoOsobi;
  
  if(info.vreme) {
    danZauzima = info.radnici.length + danZauzima;
  }
  var dodatnoColona = 0;
  if(info.vreme)
    dodatnoColona = edays * info.radnici.length;
  

 
  
  
  for (var days = 0; days < edays; ++days) { //prebroj dane
    if(dateShouldBeSkipped(info.date,days,neradniDani))
      continue;
    ++imaDana;
  }
  
  if(currentWorksheet.getMaxColumns() < edays * danZauzima + dodatnoColona)
    currentWorksheet.insertColumns(20, imaDana * danZauzima + dodatnoColona);


  var danaRange = currentWorksheet.getRange(1,10,1,4);
  danaRange.setValues([["DANA:",imaDana,"TERMINA:",hoursCount]]);
  applyFormating(danaRange,darkBlue,lightTextColor);
  
  var rangePrethodni = currentWorksheet.getRange(4, 2, hoursCount*2, kolicinaKolonaPoOsobi)  

  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  var textFinder = settingsSheet.createTextFinder("Cenovnik");
  var firstOccurrence = textFinder.findNext();
  var rangeUsluga = settingsSheet.getRange(firstOccurrence.getRowIndex()+2, firstOccurrence.getColumn(), settingsSheet.getLastRow()-firstOccurrence.getRowIndex()-2, 1);
  

  var textFinder1 = settingsSheet.createTextFinder("Radnici");
  var firstOccurrence1 = textFinder1.findNext();
  var rangeRadnik = settingsSheet.getRange(firstOccurrence1.getRowIndex()+2, firstOccurrence1.getColumn(), settingsSheet.getLastRow() - firstOccurrence.getRowIndex()-2, 1);


  var rule = SpreadsheetApp.newDataValidation().requireValueInRange(rangeUsluga).build();
  var ruleRadnik = SpreadsheetApp.newDataValidation().requireValueInRange(rangeRadnik).build();
    
  var rangeDays = currentWorksheet.getRange(2, 2, 1, danZauzima);
  var rangeDates = currentWorksheet.getRange(3, 2, 1, danZauzima);
  var rangePerson = currentWorksheet.getRange(4, 2, 1, kolicinaKolonaPoOsobi);
  var rangeHeaders = currentWorksheet.getRange(5, 2, 1, kolicinaKolonaPoOsobi);

  var rangeCenaTable = currentWorksheet.getRange(6,3, hoursCount , 1);
  var rangeUslugaTable = currentWorksheet.getRange(6,2, hoursCount , 1);
  var rangeVremeTable = currentWorksheet.getRange(6,6, hoursCount , 1);
  var rangePomocnaUslugaTable = currentWorksheet.getRange(6,7, hoursCount , 1);
  var rangeRasporedPerson = currentWorksheet.getRange(6,2,hoursCount,kolicinaKolonaPoOsobi);
  
  var rangeTotalSum = currentWorksheet.getRange(6+hoursCount, 2, 1, kolicinaKolonaPoOsobi);
  var rangeTotalRow = currentWorksheet.getRange(6+hoursCount,1,1,1);
  
  var rangePomocnaVreme= currentWorksheet.getRange(6,kolicinaKolonaPoOsobi + 2, hoursCount , 1);
  
  
  currentWorksheet.deleteRows(rangeTotalSum.getRow()+ info.radnici.length +1, currentWorksheet.getMaxRows()-rangeTotalSum.getRow() - info.radnici.length -1);
   
  
  rangeTotalRow.setValue("INFO:");
  applyFormating(rangeTotalRow,darkBlue,lightTextColor);
  
  var range = SpreadsheetApp.getActive().getSheetByName("Settings").getRange('A5:A10');
  SpreadsheetApp.getActive().setNamedRange("R"+currentWorksheet.getName()+"TC", rangeTotalSum);
    Logger.clear();
  var sumsArray = [] //new Array(info.radnici.length);
  var workArray = [] //new Array(info.radnici.length);
  Logger.log(sumsArray);
  for(var fill = 0; fill< info.radnici.length; ++fill){
    sumsArray[fill] = "=SUM(";
    workArray[fill] = "=SUM(";
  }
  Logger.log("Init");
  Logger.log(sumsArray);
  Logger.log(workArray);
  var first = true;
  var additionalDist = 0;
  var additionalDistTotal = 0;
  if(info.vreme)
  {
     additionalDist = 1;
     additionalDistTotal = 1;
  }
   
  var rangeTimeTablePrvi;
  var rangeTimeTable;
  var currentMonth = new Date(info.date.getTime());
  var nextMonth;
  var stvarno = 0;
  for (var days = 0; days < edays; ++days) {
    if(dateShouldBeSkipped(info.date,days,neradniDani))
     continue;
    nextMonth = new Date(info.date.getTime() + days * 60 * 24 * 60000);
    if(currentMonth.getMonth() != nextMonth.getMonth()){
      Logger.log("CURRENT=" + currentMonth);
      Logger.log("NEXT DATE=" + nextMonth);
      Logger.log("Stvarno dana=" + stvarno);
      break;
    }
    ++stvarno;
    var formulaDays = '=SWITCH(UPPER(TEXT($A$2+' + days +', "dddd")), "MONDAY", "PONEDELJAK", "TUESDAY", "UTORAK", "WEDNESDAY", "SREDA", "THURSDAY", "CETVRTAK", "FRIDAY", "PETAK", "SATURDAY", "SUBOTA", "SUNDAY", "NEDELJA",  UPPER(TEXT($A$2+' +days +',"dddd")) )';
    rangeDays.merge();
    rangeDays.setFormula(formulaDays);
    rangeDays.setNumberFormat("dddd");
    applyFormating(rangeDays,darkBlue,lightTextColor);
    rangeDays.setBorder(false, true, false, true, false, false, darkTextColor, SpreadsheetApp.BorderStyle.SOLID);
 
  
    
    var formulaDates = "=$A$2+" + days;
    rangeDates.setFormula(formulaDates);
    rangeDates.merge();
    rangeDates.setNumberFormat('dd/mm/yyyy');
    applyFormating(rangeDates,darkBlue,lightTextColor);
    rangeDates.setBorder(false, true, false, true, false, false, darkTextColor, SpreadsheetApp.BorderStyle.SOLID);
    addDailyConditionalFormating(rangeDates);
    
    for (var i = 0; i < info.radnici.length; ++i) {
 

      var pomocni = currentWorksheet.getRange(rangeTotalSum.getRow()-hoursCount,rangeTotalSum.getColumn()+1); //Skroz gore //TODO PROVERI GORE.COL==DOLE.COL
      var pomocni2 = currentWorksheet.getRange(rangeTotalSum.getRow()-1,rangeTotalSum.getColumn()+1); //Skroz dole za sumu
      var strSum = '=SUM(' + pomocni.getA1Notation() + ':' + pomocni2.getA1Notation() +  ')';
      
      //Logger.log("=POM" + pomocni +":"+ pomocni2 + "rangeUslugaTable=" + rangeUslugaTable.getA1Notation() + " i RangeTotalSum=" + rangeTotalSum.getA1Notation() );
      
      var pomocniSum = currentWorksheet.getRange(rangeTotalSum.getRow(), rangeTotalSum.getColumn()+1, 1, 1);
      var pomocniWork = currentWorksheet.getRange(rangeTotalSum.getRow(), rangeTotalSum.getColumn()+3, 1, 1);
      
      sumsArray[i] = sumsArray[i] + "," + pomocniSum.getA1Notation();
      workArray[i] = workArray[i] + "," + pomocniWork.getA1Notation();
      Logger.log('sumsArray[' + i + ']=' + sumsArray[i]);
      Logger.log('workArray[' + i + ']=' + workArray[i]);
      var valuesSum = [["Total:",strSum,"Radio[h]:","0","",""]];
      rangeTotalSum.setValues(valuesSum);
      applyFormating(rangeTotalSum,darkBlue,lightTextColor); 
      if(first){
  
        if(i!=0 && info.vreme){
          var it = calculateSlots(info);          
          var rangeTimeTablePrvi = currentWorksheet.getRange(rangePomocnaVreme.getRow(), kolicinaKolonaPoOsobi*i +1 + i, 1, 1);
          var rangeTimeTable = currentWorksheet.getRange(rangePomocnaVreme.getRow()+1, kolicinaKolonaPoOsobi*i +1 + i , it-1, 1);
          makeTimeTable(currentWorksheet,info,rangeTimeTablePrvi,rangeTimeTable);
          applyFormating(currentWorksheet.getRange(rangePomocnaVreme.getRow()-2, kolicinaKolonaPoOsobi*i +1 + i, 2, 1),darkBlue,lightTextColor); 
          applyFormating(currentWorksheet.getRange(rangeTotalSum.getRow(), kolicinaKolonaPoOsobi*i +1 + i, 1, 1),darkBlue,lightTextColor); 
        }
        
      rangePerson.merge();
      rangePerson.setBorder(false, true, true, true, false, false, darkTextColor, SpreadsheetApp.BorderStyle.SOLID);
      rangePerson.setValue(info.radnici[i]);

      rangeHeaders.setValues(valuesPerPerson);  
      rangeHeaders.setBorder(true, true, false, true, false, false, darkTextColor, SpreadsheetApp.BorderStyle.SOLID);
      applyFormating(rangeHeaders,columnDataBojaRadnika[i],darkTextColor);
      applyFormating(rangePerson,columnDataBojaRadnika[i],darkTextColor);
  
      rangeRasporedPerson.setBorder(false, true, false, true, false, false, darkTextColor, SpreadsheetApp.BorderStyle.SOLID);
   
      rangeCenaTable.setFormula('=IF(INDIRECT("RC[-1]",0)=0,"",VLOOKUP(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1),Cenovnik,2,FALSE))');
      rangeUslugaTable.setDataValidation(rule);

      //rangeVremeTable.setFormula('=IF(INDIRECT("RC[-4]",0)=0,"",VLOOKUP(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-4),Cenovnik,3,FALSE))');
      rangeVremeTable.setDataValidation(ruleRadnik);
      rangePomocnaUslugaTable.setDataValidation(rule);
      
      }
      else 
      {
        rangePrethodni.copyTo(rangePrethodni.offset(0, danZauzima));
        rangePrethodni = rangePrethodni.offset(0,kolicinaKolonaPoOsobi);
        applyFormating(currentWorksheet.getRange(rangeTotalSum.getRow(), rangeTotalSum.getColumn()-1, 1, 1),darkBlue,lightTextColor); 
     
        //bandingsArray[i].copyTo(rangePrethodni);
      }
      rangeCenaTable = rangeCenaTable.offset(0, kolicinaKolonaPoOsobi + additionalDist);
      rangeUslugaTable = rangeUslugaTable.offset(0, kolicinaKolonaPoOsobi + additionalDist);
      rangeVremeTable = rangeVremeTable.offset(0, kolicinaKolonaPoOsobi + additionalDist);
      rangePomocnaUslugaTable = rangePomocnaUslugaTable.offset(0,kolicinaKolonaPoOsobi + additionalDist);
      rangeRasporedPerson  = rangeRasporedPerson.offset(0, kolicinaKolonaPoOsobi + additionalDist);
      
      rangeTotalSum = rangeTotalSum.offset(0,kolicinaKolonaPoOsobi + additionalDistTotal);
      rangePerson = rangePerson.offset(0, kolicinaKolonaPoOsobi + additionalDist);    
      rangeHeaders = rangeHeaders.offset(0, kolicinaKolonaPoOsobi + additionalDist);
    }
    if(first && info.vreme){
      kolicinaKolonaPoOsobi = kolicinaKolonaPoOsobi + 1;
      danZauzima = info.radnici.length * kolicinaKolonaPoOsobi;
      additionalDistTotal = 0;
      rangePrethodni = currentWorksheet.getRange(4, 1, hoursCount+2, kolicinaKolonaPoOsobi)  
      rangeDays = currentWorksheet.getRange(2, 2, 1, danZauzima);
      rangeDates = currentWorksheet.getRange(3, 2, 1, danZauzima);
    }
    first =false;
    rangeDays = rangeDays.offset(0, danZauzima);
    rangeDates = rangeDates.offset(0, danZauzima);
  }

  
    //START COPY PASTE
  Logger.clear();
  var dodato = 0;
  var inace = 0;
  if(info.vreme) {
    dodato = 1;
    --kolicinaKolonaPoOsobi
  }else 
    inace = 1;
 
  
  for (var i = 1; i <= info.radnici.length; ++i) {
    
    rangeRasporedPerson = currentWorksheet.getRange(6,1 + inace + kolicinaKolonaPoOsobi*(i-1) + dodato*i ,hoursCount,kolicinaKolonaPoOsobi );
    
    var banding = rangeRasporedPerson.applyColumnBanding(SpreadsheetApp.BandingTheme.BLUE, false, false); //
    banding.setFirstColumnColor(columnDataBojaRadnika[i-1])
    banding.setSecondColumnColor("White");
    
    for (var days = 0; days < stvarno-1; ++days) {
      rangeRasporedPerson  = rangeRasporedPerson.offset(0, danZauzima);
      banding.copyTo(rangeRasporedPerson);
      
    }
    //Pregledati formulu
    
    Logger.log("Kolona=" + kolicinaKolonaPoOsobi + " i="+i);
    Logger.log("BOJA NA =" + rangeRasporedPerson.getA1Notation());
  }
  

  //END COMPY PASTE
  
  
  for(var fill = 0; fill < info.radnici.length; ++fill){
    sumsArray[fill] = sumsArray[fill] + ")";
    workArray[fill] = workArray[fill] + ")";
  }
  Logger.log('sumsArray[' + fill + ']=' + sumsArray[i]);
  Logger.log('workArray[' + fill + ']=' + workArray[i]);
  rangeTotalSum = rangeTotalSum.offset(1,-6);
  for(var fill = 0; fill < info.radnici.length; ++fill){
    applyFormating(rangeTotalSum,darkBlue,lightTextColor);
    rangeTotalSum.setValues([["RADNIK:",info.radnici[fill] + ":", "ZARADIO:" , sumsArray[fill], "RADIO:",workArray[fill]]]);
    rangeTotalSum = rangeTotalSum.offset(1,0);
   
  }
}


function makeScheduleYear(form){
  if(form == undefined){
    form = new Object();
    form.start = "";
    form.date = "";
    form.name = "";
    form.end = "";
    form.int = "";
    form.date = "2017-01-01";
  }
  var godina = form.date.split("-")[0];
  form.name = "";
  for(var i=0; i<12;++i){
    form.date = godina + "-" + (i+1) + "-1";
    makeSchedule(form);
  }
}
// Kod Script Editora |-> Edit ? Current Project’s triggers
// Novi SKROZ DOLE DESNO
function timeTrigger(e){
  
  //var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  //for(var i = 0; i< sheets.length; ++i){
  //    if(A2 == NOW().MONTH()) 
  //     nadji tacan dan u RASPOREDU
  //       nadji sve kolone ZA LJUDE
  //         VLOOKUP ZA EMAIL
  //           CALENDAR OBAVESTI
   //         
   //   radi nesto
  //}  
}

//function posaljiCalendar(){
//  var email = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings").getRange("B6").getValue();
//  var event = CalendarApp.createAllDayEvent("Uuups salon frizer", new Date(2019,10,17,10,0,0,0));
//  Logger.log(email);
//  event.addEmailReminder(60);
//  event.addGuest(email);
//  event.addPopupReminder(60);
//  event.setColor(CalendarApp.EventColor.PALE_BLUE);
//  event.addEmailReminder(600);
//  event.addPopupReminder(600);
//  event.setDescription("Frizerski salon ups zakazan termin u 10:10");
//  event.setLocation("Kosovska 1,Novi Sad,Serbia");
 // event.setTime( new Date(2019,10,17,10,0,0,0),  new Date(2019,10,17,11,0,0,0));
//  event.setTag("Frizer", "Uups");
//}

function togglePastDays(toggleHidden){
 if(toggleHidden == undefined)
   toggleHidden = true;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()

  var range = sheet.getRange(3, 2, 1,sheet.getDataRange().getLastColumn()-1);


  if(toggleHidden){
    var hiding = true;
    var skrivaj = 0;
    range.getValues()[0].forEach(function(value){
      Logger.log(value);
      
      if(value != "")
      {
        var datum = new Date(value);
        Logger.log(datum.getDate());
        var today = new Date();
        today.setHours(0);
        today.setMinutes(0);
        today.setDate(today.getDate() - 1);
        if(datum > today)
        {
          hiding = false;
          Logger.log("Prestao=" + datum + " DANAS=" + today);
        }
        
      }
      if(hiding)
        ++skrivaj;
    });
    sheet.hideColumns(2, skrivaj);
    Logger.log("SKRIVAJ = " + skrivaj);  
  }else{
    sheet.showColumns(2, sheet.getMaxColumns()-2);
  }
}


function makeSchedule(form){
  Logger.log("Dobijene iz forme = " + form);
  var info = getValuesFromForm(form);
  Logger.log("Dobijene DATUM ZASTO PIKA =" + info.date);
  Logger.log("!!!!!!!!!!!!!!! Dobijeni radnici checkirani" + info.radnici);
  var currentWorksheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(info.name);
  if(currentWorksheet == null) 
    currentWorksheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(info.name);
  currentWorksheet.setFrozenColumns(1);
  currentWorksheet.setFrozenRows(5)
  var hoursCount = makeHours(currentWorksheet,info);
  makePeople(currentWorksheet,info,hoursCount);

}

function makeSettings(){
  var settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  if(settingsSheet == null) 
    settingsSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Settings");
  settingsSheet.setTabColor(darkBlue);
  var valuesMusterije = [["Ime","Prezime","Nadimak","Kontakt","Beleske","Rodendan"],["Marko","Markovic","Markic","0604445566","Voli da kasni","2011/01/01"],["Jelena","Jelanovic","jelenakomsiluk","0612223344","Blans 4B","2005/11/15"]];
  var valuesCenovnik = [["Usluga","Cena","Vreme[min]"],["Muško šisanje","700","30"],["Zensko sišanje","1000","45"],["Feniranje","100","10"]];
  var valuesNeradniDani = [["Dan","Razlog"],["2019/01/01","Nova Godina"],["2019/01/02","Nova Godina"]];
  var valuesRadnici = [["Ime","Boja"],["Sofija","#add8e6"],["Bojan","#ddd8e6"]];
  
  var dist = 1;
  
  var rangeMusterijeHeader = settingsSheet.getRange(1, 1,1,valuesMusterije[0].length);
  var rangeMusterija = settingsSheet.getRange(rangeMusterijeHeader.getRowIndex()+1,1,valuesMusterije.length,valuesMusterije[0].length);
  //settingsSheet.getRange(row, column, numRows, numColumns)
  
  var rangeCenovnikHeader = settingsSheet.getRange(1, rangeMusterijeHeader.getWidth() + rangeMusterijeHeader.getColumn() + dist, 1, valuesCenovnik[0].length); 
  var rangeCenovnik = settingsSheet.getRange(rangeCenovnikHeader.getRowIndex()+1, rangeMusterijeHeader.getWidth() + rangeMusterijeHeader.getColumn() + dist, valuesCenovnik.length, valuesCenovnik[0].length); 
  var rangeCenovnikNamed = settingsSheet.getRange(rangeCenovnikHeader.getRowIndex()+2, rangeMusterijeHeader.getWidth() + rangeMusterijeHeader.getColumn() + dist, settingsSheet.getMaxRows() - rangeCenovnikHeader.getRowIndex()-2, valuesCenovnik[0].length);

  var range = SpreadsheetApp.getActive().getSheetByName("Settings").getRange('A5:A10');
  SpreadsheetApp.getActive().setNamedRange('Cenovnik', rangeCenovnikNamed);
  
  //var namedPused = named.push(rangeCenovnikNamed);

  
  
  var rangeNeradniDaniHeader = settingsSheet.getRange(1, rangeCenovnikHeader.getWidth() + rangeCenovnikHeader.getColumn() + dist, 1, valuesNeradniDani[0].length); 
  var rangeNeradniDani = settingsSheet.getRange(rangeNeradniDaniHeader.getRowIndex()+1, rangeCenovnikHeader.getWidth() + rangeCenovnikHeader.getColumn() + dist, valuesNeradniDani.length, valuesNeradniDani[0].length); 
  
  var rangeRadniciHeader = settingsSheet.getRange(1, rangeNeradniDaniHeader.getWidth()  + rangeNeradniDaniHeader.getColumn() + dist, 1, valuesRadnici[0].length); 
  var rangeRadnici = settingsSheet.getRange(rangeRadniciHeader.getRowIndex()+1, rangeNeradniDaniHeader.getWidth() + rangeNeradniDaniHeader.getColumn() + dist, valuesRadnici.length, valuesRadnici[0].length); 
  
  var bandingThema = SpreadsheetApp.BandingTheme.BLUE;
  
  //Musterije Table
  makeHeaderForTable(rangeMusterijeHeader,"Mušterije");
  rangeMusterija.setValues(valuesMusterije);
  try{
    rangeMusterija.applyRowBanding(bandingThema, true, false);
    var criteria = SpreadsheetApp.newFilterCriteria();
    var slicer = settingsSheet.insertSlicer(rangeMusterija, 5, 5);
    slicer.setTitle("Pretraga Musterija"); 
    slicer.setColumnFilterCriteria(1, criteria);
    slicer.setApplyToPivotTables(false);
    
  } catch(err){
    throw "Za uspesno ponovno pokretanje unistite Settings list"
  }
  //Cenovnik Table
  makeHeaderForTable(rangeCenovnikHeader,"Cenovnik");
  rangeCenovnik.setValues(valuesCenovnik);
  if(!moreRunes){
    rangeCenovnik.applyRowBanding(bandingThema, true, false);
  }
  //Neradni dani Table
  makeHeaderForTable(rangeNeradniDaniHeader,"Neradni Dani");
  rangeNeradniDani.setValues(valuesNeradniDani);
  if(!moreRunes){
    rangeNeradniDani.applyRowBanding(bandingThema, true, false);
  }
  //Radnici Table
  makeHeaderForTable(rangeRadniciHeader,"Radnici");
  rangeRadnici.setValues(valuesRadnici);
  if(!moreRunes){
    rangeRadnici.applyRowBanding(bandingThema, true, false);
  }

}

//END Verzija 5