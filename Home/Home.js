(function () {
    "use strict";

    var messageBanner;

    const tamnoPlava = "#2d5898";
    const fillYellow = "#fff200";
    const fillGreen = "#247348";

    // The initialize function must be run each time a new page is loaded

    // =$A$2+(0-WEEKDAY($A$2,2)+1) Pocetni dan bude ponedeljak

    Office.initialize = function (reason) {
        $(document).ready(function () {
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            //prepare();

            $('#radniciButton').click(zaSvakogRadnika);

            $('#get-data-from-selection').click(prepare);
            $('#make-headers').click(makeHeaders);
            $('#test').click(jestePraznik);
            zaSvakogRadnika();
        });
    }

    function daysInMonth(month, year) {
        return new Date(year, month, 0).getDate();
    }

    function jestePraznik(date, praznici) {
       for (var i = 0; i < praznici.length; ++i) {
            if (date.getFullYear() == praznici[i].getFullYear() && date.getMonth() == praznici[i].getMonth() && date.getDate() == praznici[i].getDate()) {
                return true;
            }
        }
        return false;
    }

    function zaSvakogRadnika() {
        Excel.run(function (context) {

            var settingsSheet = context.workbook.worksheets.getItem("settings");
            var radniciTable = settingsSheet.tables.getItem("Radnici");
            var columnDataRadnici = radniciTable.columns.getItem("Radnici").getDataBodyRange();
            var range = columnDataRadnici;
            range.load("values"); // https://docs.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-preview
            return context.sync().then(function () {
                $("#radnici").html("");
                range.values.forEach(function (item) {
                    $('#radnici').append('<input type="checkbox" name="' + item + '" value="' + item + '">' + item + ' <br>');
                });
                context.sync();
            });
          


        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    function tableHeader(range, text , settingsSheet){
        var cenovnikHeader = settingsSheet.getRange(range);
        cenovnikHeader.merge(true);
        cenovnikHeader.values = text;
        cenovnikHeader.format.fill.color = tamnoPlava;
        cenovnikHeader.format.horizontalAlignment = "Center";
        cenovnikHeader.format.font.color = "white";
    }

    function prepare() {
        Excel.run(function (context) {                

            var settingsSheet = context.workbook.worksheets.add("settings");

            //Radnici
            tableHeader("A1:B1", "Radnici", settingsSheet);

            var radniciTable = settingsSheet.tables.add("A2:B2", true /*hasHeaders*/);
            radniciTable.name = "Radnici";

            // Neradni Dani Table

            tableHeader("E1:F1", "Neradni Dani", settingsSheet);


            var neRadniDaniTable = settingsSheet.tables.add("E2:F2", true);
            neRadniDaniTable.name = "NeradniDani";

            neRadniDaniTable.getHeaderRowRange().values =
                [["Razlog", "Datum"]];

            neRadniDaniTable.columns.getItemAt(1).getDataBodyRange().numberFormat = "dd.mm.yyyy";


                // Cenovnik

            tableHeader("M1:O1", "Cenovnik", settingsSheet);

            var cenovnikTable = settingsSheet.tables.add("M2:O2", true /*hasHeaders*/);
            cenovnikTable.name = "Cenovnik";

            cenovnikTable.getHeaderRowRange().values =
                [["Usluga", "Cena", "Vreme"]];

            cenovnikTable.rows.add(null, [
                ["Muško šisanje", 700, 30],
                ["Zensko šišanje", 1000, 45],
            ]);

            // Musterije

            tableHeader("H1:K1", "Mušterije", settingsSheet);

            var musterijeTable = settingsSheet.tables.add("H2:K2", true);
            musterijeTable.name = "Mušterije";

            musterijeTable.getHeaderRowRange().values =
                [["Ime", "Prezime","Kontakt", "Beleške"]];

            musterijeTable.columns.getItemAt(1).getDataBodyRange().numberFormat = "0000000000#";


            musterijeTable.rows.add(null, [["Marko", "Markovic", "0614564546", "Voli da kasni"]]);

            var current = new Date();
            var novaGod1 = new Date(current.getFullYear(), 1, 1);
            var novaGod2 = new Date(current.getFullYear(), 1, 2);
            var bozic = new Date(current.getFullYear(), 1, 7);
            var drzavnost1 = new Date(current.getFullYear(), 2, 15);
            var drzavnost2 = new Date(current.getFullYear(), 2, 16);
            var prRada1 = new Date(current.getFullYear(), 4, 1);
            var prRada2 = new Date(current.getFullYear(), 4, 2);
            
            /*var snovaGod1 = "=DATE(" + novaGod1.getFullYear() + "," + novaGod1.getMonth() + "," + novaGod1.getDate() + ")"; 
            var snovaGod2 = "=DATE(" + novaGod2.getFullYear() + "," + novaGod2.getMonth() + "," + novaGod2.getDate() + ")";
            var sbozic = "=DATE(" + bozic.getFullYear() + "," + bozic.getMonth() + "," + bozic.getDate() + ")";
            var sdrazvnosti1 = "=DATE(" + drzavnost1.getFullYear() + "," + drzavnost1.getMonth() + "," + drzavnost1.getDate() + ")";
            var sdrazvnosti2 = "=DATE(" + drzavnost2.getFullYear() + "," + drzavnost2.getMonth() + "," + drzavnost2.getDate() + ")";
            var sprRada1 = "=DATE(" + prRada1.getFullYear() + "," + prRada1.getMonth() + "," + prRada1.getDate() + ")";
            var sprRada2 = "=DATE(" + prRada2.getFullYear() + "," + prRada2.getMonth() + "," + prRada2.getDate() + ")";*/

            var snovaGod1 = novaGod1.getFullYear() + "/" + novaGod1.getMonth() + "/" + novaGod1.getDate(); 
            var snovaGod2 = novaGod2.getFullYear() + "/" + novaGod2.getMonth() + "/" + novaGod2.getDate();
            var sbozic = bozic.getFullYear() + "/" + bozic.getMonth() + "/" + bozic.getDate();
            var sdrazvnosti1 = drzavnost1.getFullYear() + "/" + drzavnost1.getMonth() + "/" + drzavnost1.getDate();
            var sdrazvnosti2 = drzavnost2.getFullYear() + "/" + drzavnost2.getMonth() + "/" + drzavnost2.getDate();
            var sprRada1 = prRada1.getFullYear() + "/" + prRada1.getMonth() + "/" + prRada1.getDate();
            var sprRada2 = prRada2.getFullYear() + "/" + prRada2.getMonth() + "/" + prRada2.getDate();

          
            neRadniDaniTable.rows.add(null, [ //BUG
                ["Nova Godina", snovaGod1],
                ["Nova Godina", snovaGod2],
                ["Bozic",sbozic],
                ["Dan drzavnosti", sdrazvnosti1],
                ["Dan drzavnosti", sdrazvnosti2],
                ["Praznik rada", sprRada1],
                ["Praznik rada", sprRada2],
            ]); 

            radniciTable.getHeaderRowRange().values =
                [["Radnici", "Boja"]];

            radniciTable.rows.add(null, [
                ["Bojan", "#23b049"],
                ["Sofija", "#22bfa8"],
                ["Radnik3", "#bdb942"]
                ]);

         
            //var nameSourceRange = cenovnikTable.columns.getItemAt(0).getDataBodyRange();
            //nameSourceRange.name = "CenovnikLista";
           
            return context.sync().then(function () {
                zaSvakogRadnika();
            });

        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

     
    }

    function addToRangeMonthConditionalFormat(rangeDataMonth) {

        var conditionalFormat = rangeDataMonth.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
        );
        var conditionalFormat2 = rangeDataMonth.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
        );
        var conditionalFormat3 = rangeDataMonth.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
        );

        conditionalFormat3.cellValue.rule = {
            formula1: "=EDATE(NOW(),1)-DAY(EDATE(NOW(),1))+1", formula2: "=DATE(YEAR(EDATE(NOW(),1)),MONTH(EDATE(NOW(),1))+1,1)-1", operator: "Between",
        };
        conditionalFormat3.custom.format.fill.color = fillYellow;
        conditionalFormat3.custom.format.font.color = "black";


        conditionalFormat2.cellValue.rule = {
            formula1: "=NOW()-DAY(NOW())", formula2: "=DATE(YEAR(NOW()),MONTH(NOW())+1,1)-1", operator: "Between",
        };
        conditionalFormat2.custom.format.fill.color = fillGreen;
        conditionalFormat2.custom.format.font.color = "white";

        conditionalFormat.cellValue.rule = {
            formula1: "=NOW()-DAY(EOMONTH(TODAY(),0))", operator: "LessThan",
        };
        conditionalFormat.custom.format.fill.color = "red";
        conditionalFormat.custom.format.font.color = "white";
    }


    function addDayConditionalFormat(rangeDataInput) {

        var fconditionalFormat = rangeDataInput.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
        );
        var fconditionalFormat2 = rangeDataInput.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
        );
        var fconditionalFormat3 = rangeDataInput.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
        );

        fconditionalFormat.cellValue.rule = {
            formula1: "=NOW()-1", operator: "LessThan",
        };
        fconditionalFormat.custom.format.fill.color = "red";
        fconditionalFormat.custom.format.font.color = "white";


        fconditionalFormat2.cellValue.rule = {
            formula1: "=NOW()-1", formula2: "=NOW()", operator: "Between",
        };
        fconditionalFormat2.custom.format.fill.color = fillGreen;
        fconditionalFormat2.custom.format.font.color = "white";


        fconditionalFormat3.cellValue.rule = {
            formula1: "=NOW()", formula2: "=NOW()+4", operator: "Between",
        };
        fconditionalFormat3.custom.format.fill.color = fillYellow;
        fconditionalFormat3.custom.format.font.color = "black";

    }





    function addToRangeWeekdayConditionalFormat(rangeDataInput) {

        var dconditionalFormat = rangeDataInput.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
        );
        var dconditionalFormat2 = rangeDataInput.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
        );
        var dconditionalFormat3 = rangeDataInput.conditionalFormats.add(
            Excel.ConditionalFormatType.cellValue
        );

        dconditionalFormat.cellValue.rule = {
            formula1: "=NOW()+(0-WEEKDAY(NOW(),2)+1)", operator: "LessThan",
        };
        dconditionalFormat.custom.format.fill.color = "red";
        dconditionalFormat.custom.format.font.color = "white";


        dconditionalFormat2.cellValue.rule = {
            formula1: "=NOW()-1+(0-WEEKDAY(NOW(),2)+1)", formula2: "=NOW()+7+(0-WEEKDAY(NOW(),2)+1)", operator: "Between",
        };
        dconditionalFormat2.custom.format.fill.color = fillGreen;
        dconditionalFormat2.custom.format.font.color = "white";


        dconditionalFormat3.cellValue.rule = {
            formula1: "=NOW()+7-1+(0-WEEKDAY(NOW(),2)+1)", formula2: "=NOW()+14+(0-WEEKDAY(NOW(),2)+1)", operator: "Between",
        };
        dconditionalFormat3.custom.format.fill.color = fillYellow;
        dconditionalFormat3.custom.format.font.color = "black";
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




    function makeHeaders() {
        Excel.run(function (context) {

   

            var pocetak = $("#pocetak").val();
            var datum = $("#mesec").val();
            var kraj = $("#kraj").val();
            var interval = $("#interval").val();
            var sheetName = $("#ime").val();

           

            if (pocetak == "")
                pocetak = "8";
            if (datum == "")
                datum = "1.1.2019";
            if (kraj == "" ) {          
                kraj = "20";
            }
            if (interval == "")
                interval = "30";


                
            var naslov = " \t   Popunite podatke";
            var message1 = "    Unesite dobro datum";
            var splitDate = datum.split(".");
            var datumZahtevani;
            try {
                datumZahtevani = new Date(splitDate[2], splitDate[1] - 1, splitDate[0]);
            }
            catch (error) {
                showNotification(naslov, message1);
                return;
            }

            if (sheetName == "") {
                var nazivDatuma = translate(datumZahtevani);
                sheetName = nazivDatuma + splitDate[2];
            } 

            if (datum.search(".") == -1) {
                showNotification(naslov, message1);
                return;
            } else if (typeof (splitDate) == "undefined") {
                showNotification(naslov, message1);
                return;
            } else if (splitDate.length != 3) {
                showNotification(naslov, message1);
                return;
            }

            var currentWorksheet = context.workbook.worksheets.add(sheetName);
            var range = currentWorksheet.getRange("A1:I2");

            range.load("values");
            range.format.fill.color = tamnoPlava;
            range.format.font.color = "white";
            range.format.horizontalAlignment = "Center";
            range.numberFormat = [[null, "dd.mm.yyyy", null, "hh:mm", null, null, null,null, "hh:mm"], ["dd.mm.yyyy", null, null, null, null, null,null,null,null]];
            range.format.columnWidth = 60;

            var rangeFix = currentWorksheet.getRange("A3:A5");
            rangeFix.values = [[null], [null], ["Vreme:"]];
            rangeFix.format.horizontalAlignment = "Center";
            rangeFix.format.fill.color = tamnoPlava;
            rangeFix.format.font.color = "white";

            var od = new Date();
            od.setHours(pocetak);
            od.setMinutes(0);
            od.setSeconds(0);
            var dok = new Date();
            dok.setHours(kraj);
            dok.setMinutes(0);
            dok.setSeconds(0);

                               
            var it = 0;
            for (var i = od; i < dok; ) {
                it++;
                i = new Date(i.getTime() + interval * 60000);
            }
                   
            it = it + 7;
            var rangeTimeTablePrvi = currentWorksheet.getRange("A6");
            var rangeTimeTable = currentWorksheet.getRange("A7:A" + it);
            rangeTimeTable.load("values");
            rangeTimeTable.load("formulas")
              
            var settingsSheet = context.workbook.worksheets.getItem("settings");
            var radniciTable = settingsSheet.tables.getItem("Radnici");
            var columnDataRadnici = radniciTable.columns.getItem("Radnici").getDataBodyRange();
            var columnDataBojaRadnika = radniciTable.columns.getItem("Boja").getDataBodyRange();
            columnDataRadnici.load("values"); 
            columnDataBojaRadnika.load("values");

        
            var neradniDanTable = settingsSheet.tables.getItem("NeradniDani");
            var rangeDatuma = neradniDanTable.columns.getItemAt(1).getDataBodyRange();
            rangeDatuma.load("text");
      


            return context.sync().then(function () {

                currentWorksheet = context.workbook.worksheets.getItem(sheetName); //DUNNO zasto treba

                var prazniciArray = new Array(rangeDatuma.text.length);
                var brojNeradnih = 0;
                rangeDatuma.text.forEach(function (item) {

                    var splitDate = item[0].split(".");
                    var provera = new Date(splitDate[2], splitDate[1]-1, splitDate[0]);
                    prazniciArray[brojNeradnih] = provera;
                    ++brojNeradnih;

                });

                

              
                rangeTimeTablePrvi.values = "=$D$1";
                rangeTimeTablePrvi.numberFormat = "hh:mm";
                rangeTimeTablePrvi.format.horizontalAlignment = "Center";

                rangeTimeTable.formulas = '=IF(A6="","",IF(A6+TIME(0,$G$1,0)<=$I$1,A6+TIME(0,$G$1,0),""))';
                rangeTimeTable.numberFormat = "hh:mm";
                rangeTimeTable.format.horizontalAlignment = "Center";
                rangeTimeTable.format.borders.getItem('EdgeBottom').color = tamnoPlava;
                rangeTimeTable.format.borders.getItem('EdgeBottom').weight = "Hairline";

                range.values = [["DATUM:", "=$A$2", "POČETAK:", "=TIME(" + pocetak + ",0,0)", "INTERVAL:", null,interval, "Kraj:", "=TIME(" + kraj + ",0,0)"],
                                ["=DATE(" + splitDate[2] + "," + splitDate[1] + "," + splitDate[0] + ")", null , null ,null ,null, null,null,null,null]
                ];

                var rangeDate = currentWorksheet.getRange("A2");

                addToRangeMonthConditionalFormat(rangeDate);

                var rangeDate2 = currentWorksheet.getRange("B1");

                //addToRangeWeekdayConditionalFormat(rangeDate2); //Deprecated

                addToRangeMonthConditionalFormat(rangeDate2);

                //Pocetak kreiranja headera za tabelu
                var names = [];
                var iter = 0;
                $("#radnici :checkbox").each(function () {
                    if (this.checked) {
                        names[iter] = this.getAttribute("name");
                        iter++;
                    }
                });
                /*names.sort(function (a, b) {
                    return ('' + a.attr).localeCompare(b.attr);
                })*/

                var kolicinaKolonaPoOsobi = 5; //TODO promenuti ako ima vise kolona

                var danZauzima = names.length * kolicinaKolonaPoOsobi;
                
                var dat = new Date(splitDate[2], splitDate[1]-1, splitDate[0]);


                var month = dat.getMonth() + 1;
                var year = dat.getFullYear();
                var edays = daysInMonth(month, year);

                var radniciArray = new Array(names.length);
                var bojaRadnikaArray = new Array(names.length);

                var brojac = 0;
                columnDataRadnici.values.forEach(function (item) {
                    radniciArray[brojac] = item;
                    ++brojac;
                });
                brojac = 0;
                columnDataBojaRadnika.values.forEach(function (item) {
                    bojaRadnikaArray[brojac] = item;
                    ++brojac;
                });
                var imaDana = 0;
                for (var days = 0; days < edays; ++days) { //prebroj dane
                    var check = new Date(dat.getTime() + days * 60 * 24 * 60000);
                    var danUNedelji = check.getDay();
                   
                    if (danUNedelji == 0) { //SKIP NEDELJA
                        continue;
                    }
                    try {
                        var jeliPraznik = jestePraznik(check, prazniciArray);
                        if (jeliPraznik == true) {
                            continue;
                        }
                    } catch (err) {
                        showNotification("Greska prilikom citanja neradnih dana", "Greska u 530 liniji");
                    }
                    ++imaDana;
                }

                var rangeTable = currentWorksheet.getRangeByIndexes(4, 1, it - 7, danZauzima * imaDana); //* edays
                var rasporedTable = currentWorksheet.tables.add(rangeTable, false /*hasHeaders*/);
                rasporedTable.name = "Raspored" + sheetName;
                rasporedTable.showBandedRows = false;
                rasporedTable.showBandedColumns = true;
                rasporedTable.showHeaders = false;
                rasporedTable.load("values");

                //rasporedTable.style = 'TableStyleMedium9' //https://stackoverflow.com/questions/44787595/how-do-i-set-table-style-with-office-js-for-excel-on-mac

                var cenovnikTableSettings = context.workbook.worksheets.getItem("settings").tables.getItem("Cenovnik"); //Ucitavanje
                var nameSourceRange = cenovnikTableSettings.columns.getItemAt(0).getDataBodyRange();
                nameSourceRange.load("address");


      

                return context.sync().then(function () {

                    var rangeUslugaTable = rasporedTable.columns.getItemAt(0).getDataBodyRange();
                    var rangeCenaTable = rasporedTable.columns.getItemAt(1).getDataBodyRange();
                    var rangeVremeTable = rasporedTable.columns.getItemAt(4).getDataBodyRange();
                    //rangeUslugaTable.values = 1;

                               
                    var check = nameSourceRange.address;
                    var novi = "";
                    var naleteo = 0;
                    for (var c = 0; c < check.length; ++c) {
                        if (check[c].toUpperCase() == check[c].toLowerCase()) {
                            if (naleteo == 0 || naleteo == 2) {
                                novi = novi + check[c] + "$";
                            }
                            else if (naleteo == 1 || naleteo == 3) {
                                novi = novi + "$" + check[c];
                            }
                            ++naleteo;
                        } else {
                            novi = novi + check[c];
                        }

                    }

                    var rangeDays = currentWorksheet.getRangeByIndexes(1, 1, 1, danZauzima);
                    var rangeDates = currentWorksheet.getRangeByIndexes(2, 1, 1, danZauzima);
                    var rangePerson = currentWorksheet.getRangeByIndexes(3, 1, 1, kolicinaKolonaPoOsobi);
                    var rangeHeaders = currentWorksheet.getRangeByIndexes(4, 1, 1, kolicinaKolonaPoOsobi);


                    for (var days = 0; days < edays; ++days) {
                        var check = new Date(dat.getTime() + days * 60 * 24 * 60000);
                        var danUNedelji = check.getDay();
                        if (danUNedelji == 0) { //SKIP NEDELJA
                            continue;
                        } 
                        if (jestePraznik(check, prazniciArray) == true) {
                            continue;
                        }

                        //var formulaDays = '=UPPER(TEXT(=$A$2+' + days + ',"dddd"))'; =SWITCH(UPPER(TEXT($A$2+1; "dddd")); "MONDAY"; "PONEDELJAK"; "TUESDAY"; "UTORAK"; "WEDNESDAY"; "SREDA"; "THURSDAY"; "ČETVRTAK"; "FRIDAY"; "PETAK"; "SATURDAY"; "SUBOTA"; "SUNDAY"; "NEDELJA"; UPPER(TEXT($A$2+1;"dddd")) )
                        var formulaDays = '=SWITCH(UPPER(TEXT($A$2+' + days +', "dddd")), "MONDAY", "PONEDELJAK", "TUESDAY", "UTORAK", "WEDNESDAY", "SREDA", "THURSDAY", "ČETVRTAK", "FRIDAY", "PETAK", "SATURDAY", "SUBOTA", "SUNDAY", "NEDELJA",  UPPER(TEXT($A$2+' +days +',"dddd")) )';
                        rangeDays.merge(true)
                        rangeDays.numberFormatLocal = 'dddd';
                        rangeDays.formulas = formulaDays;
                        rangeDays.format.horizontalAlignment = "Center";
                        rangeDays.format.fill.color = tamnoPlava;
                        rangeDays.format.font.color = "white";
                        rangeDays.format.borders.getItem('EdgeLeft').color = 'Black';
                        rangeDays.format.borders.getItem('EdgeRight').color = 'Black';


                        var formulaDates = "=$A$2+" + days;
                        rangeDates.formulas = formulaDates;
                        rangeDates.merge(true);
                        rangeDates.numberFormatLocal = 'dd/mm/yyyy';
                        rangeDates.format.horizontalAlignment = "Center";
                        rangeDates.format.fill.color = tamnoPlava;
                        rangeDates.format.font.color = "white";
                        rangeDates.format.borders.getItem('EdgeLeft').color = 'Black';
                        rangeDates.format.borders.getItem('EdgeRight').color = 'Black';
                        

                        for (var i = 0; i < names.length; ++i) {
                            rangePerson.merge();
                            rangePerson.format.horizontalAlignment = "Center";
                            rangePerson.format.font.color = "white";
                            rangePerson.format.borders.getItem('EdgeLeft').color = 'Black';
                            rangePerson.format.borders.getItem('EdgeRight').color = 'Black';
                            rangeHeaders.format.borders.getItem('EdgeBottom').color = 'Black';

                            rangeHeaders.values = [["Usluga", "Cena", "Kupac", "Beleške", "Vreme[min]"]];
                            rangeHeaders.format.borders.getItem('EdgeLeft').color = 'Black';
                            rangeHeaders.format.borders.getItem('EdgeRight').color = 'Black';
                            rangeHeaders.format.font.color = "white";
                            rangeHeaders.format.horizontalAlignment = "Center";
                            rangeHeaders.format.borders.getItem('EdgeTop').color = 'Black';

                            for (var j = 0; j < names.length; ++j) {
                                if (names[i] == radniciArray[j]) {
                                    rangePerson.format.fill.color = bojaRadnikaArray[j][0];
                                    rangeHeaders.format.fill.color = bojaRadnikaArray[j][0];
                                }

                            }

                            rangeCenaTable.formulas = '=IF(INDIRECT("RC[-1]",0)=0,"",VLOOKUP(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1),Cenovnik,2,FALSE))';

                            rangeUslugaTable.dataValidation.rule = {
                                list: {
                                    inCellDropDown: true,
                                    source: "=" + novi
                                }
                            };

                            addDayConditionalFormat(rangeDates);

                            rangeVremeTable.formulas = '=IF(INDIRECT("RC[-4]",0)=0,"",VLOOKUP(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-4),Cenovnik,3,FALSE))';

                            rangeCenaTable = rangeCenaTable.getOffsetRange(0, kolicinaKolonaPoOsobi);
                            rangeUslugaTable = rangeUslugaTable.getOffsetRange(0, kolicinaKolonaPoOsobi);
                            rangeVremeTable = rangeVremeTable.getOffsetRange(0, kolicinaKolonaPoOsobi);

                            rangePerson.values = names[i];
                            rangePerson = rangePerson.getOffsetRange(0, kolicinaKolonaPoOsobi);

                            rangeHeaders = rangeHeaders.getOffsetRange(0, kolicinaKolonaPoOsobi);
                        }
                        rangeDays = rangeDays.getOffsetRange(0, danZauzima);
                        rangeDates = rangeDates.getOffsetRange(0, danZauzima);
                    }

                    

                   
                    // Freeze the first two columns in the worksheet.
                    currentWorksheet.freezePanes.freezeColumns(1);
                    currentWorksheet.freezePanes.freezeRows(5);
                    return context.sync();
                });
            });
        })
            .catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    }

   

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        if (Office.context.document.getSelectedDataAsync) {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        showNotification('The selected text is:', '"' + result.value + '"');

                    } else {
                        showNotification('Error:', result.error.message);
                    }
                }
            );
        } else {
            app.showNotification('Error:', 'Reading selection data is not supported by this host application.');
        }
    }
    
    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();