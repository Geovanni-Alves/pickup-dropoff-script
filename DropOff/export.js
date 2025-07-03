function ExportDropOffTime() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sheetName = sheet.getName();
  
  if (sheetName != "ETA_Parents") {
    return;
  }
  
  var vehicleListRange = sheet.getRange("H2:I"); 
  var vehicleListValues = vehicleListRange.getValues();
  
  // Loop through each row in the vehicle list
  for (var i = 0; i < vehicleListValues.length; i++) {
    var vehicleName = vehicleListValues[i][0]; // Vehicle name in column H
    var isTrue = vehicleListValues[i][1]; // True/false value in column I
    
    if (isTrue === true) { // Check if the value is true
      var vehicleCell = sheet.getRange("C6");
      vehicleCell.setValue(vehicleName); // Update the vehicle name in C6
      // //Utilities.sleep(3000)


      // var dateOfRoute = sheet.getRange("C4").getValue();
      // var selectRanges = sheet.getRange("B2:C25").activate();
      
           // Wait until the range is not "loading..."
      var selectRanges = null;
      while (selectRanges === null || selectRanges == "loading...") {
        Utilities.sleep(1000); // Sleep for 1 second
        selectRanges = sheet.getRange("B2:C25").activate(); // Get values of the range
      }

      // // Activate the range
      // sheet.getRange("B2:C25").activate();

      var dateOfRoute = sheet.getRange("C4").getValue();
      
      exportPartAsPDF(" Time - " + vehicleName + " - (" + formattedDate(dateOfRoute) + ")");
      
 
      
    }
  }

  const htmlOutput = HtmlService 
      .createHtmlOutput('<p><p>') // <a href="' + pdfFile.getUrl() + '" target="_blank">' + fileName + '</a></p>')
      .setWidth(300)
      .setHeight(80);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Successful');
      selectRanges = sheet.getRange("A5").activate();
      return; // Exit the function once a true value is found and vehicle name is updated
  
  // If no true value is found in the list, you might want to handle this case
  // For example, display an error message or take appropriate action
  SpreadsheetApp.getUi().alert("No vehicle with 'TRUE' status found in the list.");
}




// function ExportDropOffTime() 
// {
//   var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//   var sheetName = sheet.getName();
//   //console.log(sheet.getName());
//   if (sheetName != "ETA_Parents"){
//     return;
//   }
//   var vehicle = sheet.getRange("C6").getValue();
//   var dateOfRoute = sheet.getRange("C4").getValue();
//   var selectRanges = sheet.getRange("B2:C25").activate();
  
//   exportPartAsPDF(" Time - " + vehicle + " - (" + formattedDate(dateOfRoute) + ")");
  
  
  
//   const htmlOutput = HtmlService 
//      .createHtmlOutput('<p><p>') // <a href="' + pdfFile.getUrl() + '" target="_blank">' + fileName + '</a></p>')
//      .setWidth(300)
//     .setHeight(80)
  
//   SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Successful')
//   selectRanges = sheet.getRange("A5").activate();
// }


// function saveExportRoute() 
// {
//   var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//   var sheetName = sheet.getName();
//   //console.log(sheet.getName());
//   if (sheetName != "MONDAY" &&  sheetName != "TUESDAY" && sheetName != "WEDNESDAY" && sheetName != "THURSDAY" && sheetName != "FRIDAY"){
//     return;
//   }
  
//   //const savedRoutes = SpreadsheetApp.getActive().getSheetByName('savedRoutes');
//   var lastRow = sheet.getLastRow();

//   let rotaDiaVeiculo = sheet.getRange("A2:N" + lastRow+1).getValues();
//   //let saveRota = [];
//   //let vehicles = [];
//   let inicio = [];
//   let fim = [];
//   let vehicles = [];
//   let drivers = [];
//   var iniCount = 0;
//   //let addressRoute = [];
//   //let totalKm = []
//   var dateOfRoute = rotaDiaVeiculo[0][7];


//   for (var i = 0; i<rotaDiaVeiculo.length; i++){ // loop for pass all the rotadia sheet
//     if (rotaDiaVeiculo[i][0] == 1 && rotaDiaVeiculo[i][2] != "") {
//       var routeOrigin = rotaDiaVeiculo[i-3][8]
//       var driverName = rotaDiaVeiculo[i-2][9]
//       var busName = rotaDiaVeiculo[i-2][13]

//       drivers.push(driverName)
//       inicio.push(i + 1);
//       fim.push(getFirstEmptyRow(sheet,inicio[iniCount]) - 1);
//       //console.log(inicio[iniCount], fim[iniCount])
//       //var plateText = rotaDiaVeiculo[i-2][0];
//       //var plateIndex = plateText.indexOf("- PLATE");
//       //if (plateIndex !== -1) {
//       //  var busName = plateText.substring(0,plateIndex).trim();
//       //}
//       vehicles.push(busName)
//       //var vehicleAddress = busName
//       iniCount++;
//     } 
//     // if (rotaDiaVeiculo[i][0] >= 1 && rotaDiaVeiculo[i][8] != "") {
//     //   addressRoute.push([rotaDiaVeiculo[i][4],vehicleAddress])
//     //   totalKm.push(rotaDiaVeiculo[i][10])  
//     // } 

//   }
 
//   // for (var r = 0; r<vehicles.length; r++){
//   //   var totalAllKm = 0;
//   //   for (var a = 0;a<addressRoute.length;a++){
//   //     if (addressRoute[a][1] == vehicles[r]){
//   //       totalAllKm += (totalKm[a] * 2);
//   //       saveRota.push([formattedDate(dateOfRoute),vehicles[r],routeOrigin,addressRoute[a],totalKm[a], totalAllKm]);
//   //     }
//   //   }
//   // }
// // Check if the date already exists in "savedRoutes" sheet
// var formattedDateOfRoute = Utilities.formatDate(dateOfRoute, "GMT-0700", "EEE MMM dd yyyy HH:mm:ss 'GMT'XXX");

// //var dateExists = false;
// //var dataRange = savedRoutes.getDataRange().getValues();
// //var rowsToDelete = [];
// //for (var i = 1; i < dataRange.length; i++) { // Start from row 2 to skip header
// //   var rowData = dataRange[i][0];
// //   var formattedRowDate = Utilities.formatDate(rowData, "GMT-0700", "EEE MMM dd yyyy HH:mm:ss 'GMT'XXX");
// //   if (formattedRowDate == formattedDateOfRoute) {
// //     // Date already exists, mark row for deletion
// //     rowsToDelete.push(i + 1); // Add 1 because sheets are 1-indexed
// //     dateExists = true;
// //   }
// // }
// // if (dateExists == true) {console.log('ja existe')}
// // // Delete rows marked for deletion in reverse order to avoid index issues
// // for (var j = rowsToDelete.length - 1; j >= 0; j--) {
// //   savedRoutes.deleteRow(rowsToDelete[j]);
// // }
  

//   // var lastLineSaved=savedRoutes.getLastRow()+1
//   // if (lastLineSaved == 1 ){ lastLineSaved = 2}
//   // savedRoutes.getRange(lastLineSaved,1,saveRota.length, 6).setValues(saveRota);
//   //saveRota.push([dateOfRoute,vehicles[],addressRoute[i]]);
//   //
//   // Extract and format the date
//   //console.log(dateOfRoute);
//   //console.log('dateRoute', dateRoute)
//   for (var j = 0; j< inicio.length; j++){ //loop for how many times nees to save at pdf
//     //console.log(inicio[j],fim[j])
//     //console.log("vehicles", vehicles[j]);
//     inicio[j] = inicio[j] - 2
//     var selectRanges = sheet.getRange("A" + inicio[j] + ":H" + fim[j]).activate();
    
//     exportPartAsPDF(" - " + sheet.getName() + " - " + vehicles[j] + " - " 
//     + drivers[j] + " (" + formattedDate(dateOfRoute) + ")");
    
//   }
  
//   const htmlOutput = HtmlService 
//      .createHtmlOutput('<p><p>') // <a href="' + pdfFile.getUrl() + '" target="_blank">' + fileName + '</a></p>')
//      .setWidth(300)
//     .setHeight(80)
  
//   SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Successful')
//   selectRanges = sheet.getRange("A5").activate();
// }

function saveExportRoute() 
{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sheetName = sheet.getName();
  //console.log(sheet.getName());
  if (sheetName != "MONDAY" &&  sheetName != "TUESDAY" && sheetName != "WEDNESDAY" && sheetName != "THURSDAY" && sheetName != "FRIDAY"){
    return;
  }
  

  var vehicleStatusRange = sheet.getRange("P3:Q"); // Range containing vehicle names and their true/false status
  var vehicleStatusValues = vehicleStatusRange.getValues();

  // Filter out vehicles marked as false
  var trueVehicles = [];
  for (var i = 0; i < vehicleStatusValues.length; i++) {
    if (vehicleStatusValues[i][1] !== null && vehicleStatusValues[i][1] === true) {
      trueVehicles.push(vehicleStatusValues[i][0]);
    }
  }

  //const savedRoutes = SpreadsheetApp.getActive().getSheetByName('savedRoutes');
  var lastRow = sheet.getLastRow();

  let rotaDiaVeiculo = sheet.getRange("A2:N" + lastRow+1).getValues();
  //let saveRota = [];
  //let vehicles = [];
  let inicio = [];
  let fim = [];
  let vehicles = [];
  let drivers = [];
  var iniCount = 0;
  //let addressRoute = [];
  //let totalKm = []
  var dateOfRoute = rotaDiaVeiculo[0][7];


  for (var i = 0; i<rotaDiaVeiculo.length; i++){ // loop for pass all the rotadia sheet
    if (rotaDiaVeiculo[i][0] == 1 && rotaDiaVeiculo[i][2] != "") {
      var routeOrigin = rotaDiaVeiculo[i-3][8]
      var driverName = rotaDiaVeiculo[i-2][9]
      var busName = rotaDiaVeiculo[i-2][13]

      if (trueVehicles.includes(busName)) {
        drivers.push(driverName)
        inicio.push(i + 1);
        fim.push(getFirstEmptyRow(sheet,inicio[iniCount]) - 1);
        //console.log(inicio[iniCount], fim[iniCount])
        //var plateText = rotaDiaVeiculo[i-2][0];
        //var plateIndex = plateText.indexOf("- PLATE");
        //if (plateIndex !== -1) {
        //  var busName = plateText.substring(0,plateIndex).trim();
        //}
        vehicles.push(busName)
        //var vehicleAddress = busName
        iniCount++;
      }

    } 
    // if (rotaDiaVeiculo[i][0] >= 1 && rotaDiaVeiculo[i][8] != "") {
    //   addressRoute.push([rotaDiaVeiculo[i][4],vehicleAddress])
    //   totalKm.push(rotaDiaVeiculo[i][10])  
    // } 

  }
 
  // for (var r = 0; r<vehicles.length; r++){
  //   var totalAllKm = 0;
  //   for (var a = 0;a<addressRoute.length;a++){
  //     if (addressRoute[a][1] == vehicles[r]){
  //       totalAllKm += (totalKm[a] * 2);
  //       saveRota.push([formattedDate(dateOfRoute),vehicles[r],routeOrigin,addressRoute[a],totalKm[a], totalAllKm]);
  //     }
  //   }
  // }
// Check if the date already exists in "savedRoutes" sheet
var formattedDateOfRoute = Utilities.formatDate(dateOfRoute, "GMT-0700", "EEE MMM dd yyyy HH:mm:ss 'GMT'XXX");

//var dateExists = false;
//var dataRange = savedRoutes.getDataRange().getValues();
//var rowsToDelete = [];
//for (var i = 1; i < dataRange.length; i++) { // Start from row 2 to skip header
//   var rowData = dataRange[i][0];
//   var formattedRowDate = Utilities.formatDate(rowData, "GMT-0700", "EEE MMM dd yyyy HH:mm:ss 'GMT'XXX");
//   if (formattedRowDate == formattedDateOfRoute) {
//     // Date already exists, mark row for deletion
//     rowsToDelete.push(i + 1); // Add 1 because sheets are 1-indexed
//     dateExists = true;
//   }
// }
// if (dateExists == true) {console.log('ja existe')}
// // Delete rows marked for deletion in reverse order to avoid index issues
// for (var j = rowsToDelete.length - 1; j >= 0; j--) {
//   savedRoutes.deleteRow(rowsToDelete[j]);
// }
  

  // var lastLineSaved=savedRoutes.getLastRow()+1
  // if (lastLineSaved == 1 ){ lastLineSaved = 2}
  // savedRoutes.getRange(lastLineSaved,1,saveRota.length, 6).setValues(saveRota);
  //saveRota.push([dateOfRoute,vehicles[],addressRoute[i]]);
  //
  // Extract and format the date
  //console.log(dateOfRoute);
  //console.log('dateRoute', dateRoute)
  for (var j = 0; j< inicio.length; j++){ //loop for how many times nees to save at pdf
    //console.log(inicio[j],fim[j])
    //console.log("vehicles", vehicles[j]);
    inicio[j] = inicio[j] - 2
    var selectRanges = sheet.getRange("A" + inicio[j] + ":H" + fim[j]).activate();
    
    exportPartAsPDF(" - " + sheet.getName() + " - " + vehicles[j] + " - " 
    + drivers[j] + " (" + formattedDate(dateOfRoute) + ")");
    
  }
  
  const htmlOutput = HtmlService 
     .createHtmlOutput('<p><p>') // <a href="' + pdfFile.getUrl() + '" target="_blank">' + fileName + '</a></p>')
     .setWidth(300)
    .setHeight(80)
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Successful')
  selectRanges = sheet.getRange("A5").activate();
}

function exportPartAsPDF(fileSuffix, predefinedRanges) {
  var ui = SpreadsheetApp.getUi()
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  //"A2:G19"
  var selectedRanges
  var fileSuffix
  if (predefinedRanges) {
    selectedRanges = predefinedRanges
    fileSuffix = '-predefined'
  } else {
    //var activeRangeList = spreadsheet.setActiveRangeList(predefinedRanges)
    var activeRangeList = spreadsheet.getActiveRangeList()
    if (!activeRangeList) {
      ui.alert('Please select at least one range to export')
      return
    }
    selectedRanges = activeRangeList.getRanges()
    //console.log(selectedRanges);
    //fileSuffix = '-selected'
  }
  //var ranges = activeRangeList
  //selectedRanges = spreadsheet.getRanges("A2:G19");
  if (selectedRanges.length === 1) {
    // special export with formatting
    var currentSheet = selectedRanges[0].getSheet()
    var blob = _getAsBlob(spreadsheet.getUrl(), currentSheet, selectedRanges[0]);
    
    var fileName = "Drop-Off" + fileSuffix //spreadsheet.getName() + fileSuffix
    _exportBlob(blob, fileName, spreadsheet)
    return
  }
  
  var tempSpreadsheet = SpreadsheetApp.create(spreadsheet.getName() + fileSuffix)
  if (!saveToRootFolder) {
    DriveApp.getFileById(tempSpreadsheet.getId()).moveTo(DriveApp.getFileById(spreadsheet.getId()).getParents().next())
  }
  var tempSheets = tempSpreadsheet.getSheets()
  var sheet1 = tempSheets.length > 0 ? tempSheets[0] : undefined
  SpreadsheetApp.setActiveSpreadsheet(tempSpreadsheet)
  tempSpreadsheet.setSpreadsheetTimeZone(spreadsheet.getSpreadsheetTimeZone())
  tempSpreadsheet.setSpreadsheetLocale(spreadsheet.getSpreadsheetLocale())
  
  for (var i = 0; i < selectedRanges.length; i++) {
    var selectedRange = selectedRanges[i]
    var originalSheet = selectedRange.getSheet()
    var originalSheetName = originalSheet.getName()
    
    var destSheet = tempSpreadsheet.getSheetByName(originalSheetName)
    if (!destSheet) {
      destSheet = tempSpreadsheet.insertSheet(originalSheetName)
    }
    
    Logger.log('a1notation=' + selectedRange.getA1Notation())
    var destRange = destSheet.getRange(selectedRange.getA1Notation())
    destRange.setValues(selectedRange.getValues())
    destRange.setTextStyles(selectedRange.getTextStyles())
    destRange.setBackgrounds(selectedRange.getBackgrounds())
    destRange.setFontColors(selectedRange.getFontColors())
    destRange.setFontFamilies(selectedRange.getFontFamilies())
    destRange.setFontLines(selectedRange.getFontLines())
    destRange.setFontStyles(selectedRange.getFontStyles())
    destRange.setFontWeights(selectedRange.getFontWeights())
    destRange.setHorizontalAlignments(selectedRange.getHorizontalAlignments())
    destRange.setNumberFormats(selectedRange.getNumberFormats())
    destRange.setTextDirections(selectedRange.getTextDirections())
    destRange.setTextRotations(selectedRange.getTextRotations())
    destRange.setVerticalAlignments(selectedRange.getVerticalAlignments())
    destRange.setWrapStrategies(selectedRange.getWrapStrategies())
  }
  
  // remove empty Sheet1
  if (sheet1) {
    Logger.log('lastcol = ' + sheet1.getLastColumn() + ',lastrow=' + sheet1.getLastRow())
    if (sheet1 && sheet1.getLastColumn() === 0 && sheet1.getLastRow() === 0) {
      tempSpreadsheet.deleteSheet(sheet1)
    }
  }
  
  exportAsPDF()
  SpreadsheetApp.setActiveSpreadsheet(spreadsheet)
  DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true)
}


/**
 * @license MIT
 * 
 * © 2020 xfanatical.com. All Rights Reserved.
 * 
 * @since 1.3.4 Add retrying for printing many sheets
 * @since 1.3.3 Add an option to save PDF files in the same folder
 * @since 1.3.2 Fix time changed issue
 * @since 1.3.1 Fix a landscape problem
 * @since 1.3.0 Support printing the entire spreadsheet as separate pdf files
 * @since 1.2.3 Add formatting for a single predefined area
 * @since 1.2.2 Add formatting for a single selected area
 * @since 1.2.1 Fix a reference issue of printing current sheet
 * @since 1.2.0 Support printing current sheet as pdf
 * @since 1.1.1 Fix an error for multi-language
 * @since 1.1.0 Support printing predefined areas as a pdf
 * @since 1.0.0 Support printing the entire spreadsheet as a pdf
 *              Support printing the selected areas as a pdf
 */

// By default, PDFs are saved in your Drive Root folder
// To save in the same folder as the spreadsheet, change the value to 'false' without the single quote pair
// You must have EDIT permission to the same folder
var saveToRootFolder = false

/*
function onOpen() {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Export all sheets', 'exportAsPDF')
    .addItem('Export all sheets as separate files', 'exportAllSheetsAsSeparatePDFs')
    .addItem('Export current sheet', 'exportCurrentSheetAsPDF')
    .addItem('Export selected area', 'exportPartAsPDF')
    .addItem('Export predefined area', 'exportNamedRangesAsPDF')
    .addToUi()
}
*/
function _exportBlob(blob, fileName, spreadsheet) {
  blob = blob.setName(fileName)
  //var folder = saveToRootFolder ? DriveApp : DriveApp.getFileById(spreadsheet.getId()).getParents().next()
  var folder = DriveApp.getFoldersByName('DROPOFF_PDF_VANCOUVER').next(); 
  //var folder = DriveApp.getFolderById(folderId);
  //console.log(folder);
  var pdfFile = folder.createFile(blob);
  //var pdfFile = folder.createFile(blob)
  
  // Display a modal dialog box with custom HtmlService content.
  /*const htmlOutput = HtmlService
    .createHtmlOutput('<p><p>') // <a href="' + pdfFile.getUrl() + '" target="_blank">' + fileName + '</a></p>')
    .setWidth(300)
   .setHeight(80)
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Successful')
  */
}

function exportAsPDF() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var blob = _getAsBlob(spreadsheet.getUrl())
  var folderId = "172OwTARFfdVKg-RG5baO0CqpW878WADs"
  _exportBlob(blob, spreadsheet.getName(), folderId)
}

function _getAsBlob(url, sheet, range) {
  var rangeParam = ''
  var sheetParam = ''
  if (range) {
    rangeParam =
      '&r1=' + (range.getRow() - 1)
      + '&r2=' + range.getLastRow()
      + '&c1=' + (range.getColumn() - 1)
      + '&c2=' + range.getLastColumn()
  }
  if (sheet) {
    sheetParam = '&gid=' + sheet.getSheetId()
  }
  // A credit to https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
  // these parameters are reverse-engineered (not officially documented by Google)
  // they may break overtime.
  var exportUrl = url.replace(/\/edit.*$/, '')
      + '/export?exportFormat=pdf&format=pdf'
      + '&size=LETTER'
      + '&portrait=true'
      + '&fitw=true'       
      + '&top_margin=0.75'              
      + '&bottom_margin=0.75'          
      + '&left_margin=0.7'             
      + '&right_margin=0.7'           
      + '&sheetnames=false&printtitle=false'
      + '&pagenum=UNDEFINED' // change it to CENTER to print page numbers
      + '&gridlines=true'
      + '&fzr=FALSE'      
      + sheetParam
      + rangeParam
      
  Logger.log('exportUrl=' + exportUrl)
  var response
  var i = 0
  for (; i < 5; i += 1) {
    response = UrlFetchApp.fetch(exportUrl, {
      muteHttpExceptions: true,
      headers: { 
        Authorization: 'Bearer ' +  ScriptApp.getOAuthToken(),
      },
    })
    if (response.getResponseCode() === 429) {
      // printing too fast, retrying
      Utilities.sleep(3000)
    } else {
      break
    }
  }
  
  if (i === 5) {
    throw new Error('Printing failed. Too many sheets to print.')
  }
  
  return response.getBlob()
}

function exportAllSheetsAsSeparatePDFs() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var files = []
  var folder = saveToRootFolder ? DriveApp : DriveApp.getFileById(spreadsheet.getId()).getParents().next()
  spreadsheet.getSheets().forEach(function (sheet) {
    spreadsheet.setActiveSheet(sheet)
    
    var blob = _getAsBlob(spreadsheet.getUrl(), sheet)
    var fileName = sheet.getName()
    blob = blob.setName(fileName)
    var pdfFile = folder.createFile(blob)
    
    files.push({
      url: pdfFile.getUrl(),
      name: fileName,
    })
  })
  
  /*
  const htmlOutput = HtmlService
    .createHtmlOutput('<p>Click to open PDF files</p>'
      + '<ul>'
      + files.reduce(function (prev, file) {
        prev += '<li><a href="' + file.url + '" target="_blank">' + file.name + '</a></li>'
        return prev
      }, '')
      + '</ul>')
    .setWidth(300)
    .setHeight(150)
  */
  //SpreadsheetApp.getUi().showModalDialog('Export Successful')//htmlOutput, 'Export Successful')
}

function exportCurrentSheetAsPDF() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var currentSheet = SpreadsheetApp.getActiveSheet()
  
  var blob = _getAsBlob(spreadsheet.getUrl(), currentSheet)
  _exportBlob(blob, currentSheet.getName(), spreadsheet)
}



function exportNamedRangesAsPDF() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var allNamedRanges = spreadsheet.getNamedRanges()
  var toPrintNamedRanges = []
  for (var i = 0; i < allNamedRanges.length; i++) {
    var namedRange = allNamedRanges[i]
    if (/^print_area_.*$/.test(namedRange.getName())) {
      Logger.log('found named range ' + namedRange.getName())
      toPrintNamedRanges.push(namedRange.getRange())
    }
  }
  if (toPrintNamedRanges.length === 0) {
    SpreadsheetApp.getUi().alert('No print areas found. Please add at least one \'print_area_1\' named range in the menu Data > Named ranges.')
    return
  } else {
    toPrintNamedRanges.sort(function (a, b) {
      return a.getSheet().getIndex() - b.getSheet().getIndex()
    })
    exportPartAsPDF(toPrintNamedRanges)
  }
}



