/**
 * Student Transportation Route Management & Export System
 * --------------------------------------------------------
 * Author: Geovanni Estevam
 * Date: 2023-04-04 
 * 
 * Description:
 * This script automates the workflow for an after-school transportation program,
 * including:
 * - Generating daily student schedules by school and day
 * - Arranging and validating routes by vehicle
 * - Applying conditional formatting and data validation
 * - Exporting selected routes as organized PDF reports
 * - Supporting staff assignments, boosters, helpers, and early dismissals
 * 
 * Built using Google Apps Script for use in Google Sheets.
 * 
 * Custom tools, logic, and layout designed by Geovanni.
 */

function showAboutDialog() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Student Transportation Route Management & Export System',
    'Designed, built, and maintained by Geovanni Estevam\n' +
    'Version : 2025-07-02\n' +
    'Contact : geo-estevam@hotmail.com',
    ui.ButtonSet.OK
  );
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  // Show Route Tools menu only if not expired
  try {
    checkExpiry();

    const menu = ui.createMenu("Route Tools");
    menu.addItem("Arrange Route Order", "arrangeRoute")
        .addItem("Export Selected Vehicles to PDF", "showVehicleExportDialog")
        .addSeparator()
        .addItem("Generate Kids of the day", "getKidsSchools");
    menu.addToUi();

  } catch(e) {
    // Expired — do not add Route Tools menu
    // Optionally, you can alert user here or just skip silently
    // ui.alert('The Route Tools menu is disabled because the script has expired.');
  }
  // Always show About menu
  ui.createMenu('About…')
    .addItem('About this Sheet', 'showAboutDialog')
    .addToUi();
}



function printKidsAlphabetically() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Kids and Schools");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Sheet 'Kids and Schools' not found.");
    return;
  }

  const data = sheet.getRange("A2:B" + sheet.getLastRow()).getValues(); // A: Student, B: Responsible
  const filtered = data.filter(([name]) => name).sort((a, b) => a[0].localeCompare(b[0]));

  const today = new Date();
  const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const formattedDay = Utilities.formatDate(today, Session.getScriptTimeZone(), "EEEE");

  const tempSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Printable List");
  tempSheet.getRange("A1").setValue("DATE");
  tempSheet.getRange("B1").setValue(`${formattedDate} ${formattedDay}`);

  tempSheet.getRange("A2").setValue("Number of Kids");
  tempSheet.getRange("B2").setValue(filtered.length);

  tempSheet.getRange("A4").setValue("Student Name");
  tempSheet.getRange("B4").setValue("Responsible");
  tempSheet.getRange(5, 1, filtered.length, 2).setValues(filtered);

  tempSheet.autoResizeColumns(1, 2);

  SpreadsheetApp.getUi().alert("Printable list created in sheet: 'Printable List'.");
}


function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName()
  var range = e.range;
  var column = range.getColumn();
  var value = e.value;
  var oldValue = e.oldValue;
  // console.log(sheetName)
  // SpreadsheetApp.getUi().alert("sheet Name " + sheetName);
  
  if (sheetName === "Students x days" && (column === 3 || column === 4 || column === 5 || column === 6 || column === 7) && (value === "N" || value === "n")) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      "Are you sure?",
      "You changed the value for a non-attendance value ('N'). This might affect the schedule. Do you want to proceed?",
      ui.ButtonSet.YES_NO
    );

    if (response == ui.Button.NO) {
      // Revert the change
      range.setValue(oldValue);
    }
  }

  if (sheetName === "MONDAY" 
    || sheetName === "TUESDAY" 
    || sheetName === "WEDNESDAY"
    || sheetName === "THURSDAY"
    || sheetName === "FRIDAY"
    && column === 3) {
    // Check if the selected value is <DRIVER SEAT> or <HELPER SEAT>
    if (value !== "<DRIVER SEAT>" && value !== "<HELPER SEAT>") {
      // Check if the selected student is a duplicate within the "MONDAY" sheet
      var isDuplicate = isStudentDuplicate(sheet, value, range);

      if (isDuplicate) {
        // Clear the cell and show a message if it's a duplicate selection
        range.setValue("");
        SpreadsheetApp.getUi().alert("This student has already been selected.");
      }
    }
  }
}

function checkExpiry() {
  var expiryDate = new Date('2025-09-04'); // Set your expiry date here
  var today = new Date();
  if (today > expiryDate) {
    SpreadsheetApp.getUi().alert(
      "FATAL ERROR\n\n" +
      "A critical failure has occurred.\n" +
      "System integrity compromised.\n\n" +
      "Error Code: 0xDEADBEEF\n" +
      "Process terminated."
    );
    throw new Error("FATAL ERROR: Process terminated.");
  }
}

function reorganizeVehiclesDaysSheet() {
  checkExpiry();
  const ui = SpreadsheetApp.getUi();
  // const response = ui.alert(
  //   "Confirm Action",
  //   "This will remove all students from the 'test' sheet and reorganize the vehicles based on the registered list. This action can be UNDONE!!! Do you want to continue?",
  //   ui.ButtonSet.YES_NO
  // );
   const response = ui.Button.YES

  if (response !== ui.Button.YES) return;

  const actualDay = "Monday"; // can use as name of sheet 
  const shortDay  = actualDay.substring(0,3).toUpperCase(); // "MON" can use to find the KD_shorday = kd_Mon



  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("test");
  const vehicleSheet = ss.getSheetByName("Vehicles");

  if (!sheet || !vehicleSheet) {
    ui.alert("Error", "Could not find the 'test' or 'Vehicles' sheet.", ui.ButtonSet.OK);
    return;
  }

  // Clear all content and formatting from row 4 onward (A to H)
  const lastRow = sheet.getLastRow();
  if (lastRow > 4) {
    sheet.getRange("A2:H" + lastRow).clearContent().clearFormat();
  }

  // Fetch only vehicles marked as "In Route?" (column F = TRUE)
  const vehicleData = vehicleSheet
    .getRange("A2:F" + vehicleSheet.getLastRow())
    .getValues()
    .filter(row => row[5] === true);

  const vehicles = vehicleData.map(row => ({
    name: row[0],
    totalSeats: row[1],
    kidSeats: row[2],
    boosters: row[3],
    plate: row[4]
  }));

  let currentRow = 2;
  const helperSeatRows = [];
  const driverSeatRows = [];


  vehicles.forEach((vehicle, index) => {
    const startRow = currentRow;
    // Range covering rows 1 and 2 for the current vehicle (columns A to H)
    const vehicleHeaderRange = sheet.getRange(currentRow, 1, 2, 8);

    // merge the 2 columns A and B
    sheet.getRange(currentRow, 1, 1, 2).merge()  
    // Set font color to Light Gray 3
    vehicleHeaderRange.setBackground('#e6e6e6'); // Light Gray 3   hex

    // Set borders for all cells in this range
    vehicleHeaderRange.setBorder(true, true, true, true, true, true);
    

    // // first row (day and Date)
    sheet.getRange(currentRow, 1, 1, 8).setValues([["", "", "", "", actualDay, "", "Date:", ""]]);
    const dateFormula = "=KD_" + shortDay + "!$E$3";
    
    // Merge E and F (columns 5 and 6)
    sheet.getRange(currentRow, 5, 1, 2).merge();
    // Apply bold and center to E:F (actualDay)
    sheet.getRange(currentRow, 5, 1, 2)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setFontFamily("Calibri")
      .setFontSize(14);

    // Format "Date" title in G (column 7)
    sheet.getRange(currentRow, 7)
      .setFontWeight("bold")
      .setHorizontalAlignment("right")
      .setFontFamily("Calibri")
      .setFontSize(14);

    // Set and format date formula in H (column 8)
    
    sheet.getRange(currentRow, 8)
      .setFormula(dateFormula)
      .setHorizontalAlignment("left")
      .setFontFamily("Calibri")
      .setFontSize(14)
      .setFontWeight("bold");


    currentRow++;

    // Vehicle info
    sheet.getRange(currentRow, 1, 1, 8).setValues([[
       "", "", "" , "Vehicle", vehicle.name, "Plate N.", vehicle.plate, ""
    ]]);

    // Set Calibri 12 for columns F and G (columns 6 and 7) in rows 1 and 2 of this section
    sheet.getRange(currentRow, 6, 2, 2).setFontFamily('Calibri').setFontSize(12);
    sheet.getRange(currentRow, 6, 2, 1).setHorizontalAlignment("right");

    
    // Format "Vehicle" (column 4) and vehicle.name (column 5)
    sheet.getRange(currentRow, 4, 1, 2) // Columns D and E
      .setFontFamily('Calibri')
      .setFontSize(14)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    const firstSeatRow = currentRow + 2;
    const lastSeat = firstSeatRow + vehicle.totalSeats - 1;
    const boosterFormula = `="Boosters Required : " & COUNTIF(C${firstSeatRow}:C${lastSeat}, "*" & CHAR(8203) & "*")`;

    const boosterInVanFormula = `=IF($E${currentRow}<>"WALKING", "Boosters in Van: " & IF($E${currentRow}<>"", VLOOKUP($E${currentRow}, Vehicles!$A$3:$E$21, 4, FALSE), ""), "")`;
   
   // Merge columns A (1), B (2), and C (3) on the current row
   sheet.getRange(currentRow, 1, 1, 3).merge();
   
   const headerRow = currentRow; 
   const boosterCell = sheet.getRange(currentRow, 1);
    boosterCell.setFormula(boosterFormula);
    boosterCell.setFontFamily('Calibri');
    boosterCell.setFontSize(11);
    boosterCell.setFontWeight('bold');
    boosterCell.setHorizontalAlignment('center');
    boosterCell.setVerticalAlignment('middle');
    

    sheet.getRange(currentRow, 8)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontFamily('Calibri')
      .setFontSize(11)
      .setFontWeight('bold')
      .setFormula(boosterInVanFormula);
    

    currentRow++;

    // Header row with merged A+B for SEATS
    const headerRange = sheet.getRange(currentRow, 1, 1, 8);
    headerRange.setValues([[
      "SEATS", "", "CHILD", "SCHOOL NAME", "ORDER", "RESPONSIBLE", "DISMISSAL TIME", "SCHOOL ADDRESS"
    ]]);

    sheet.getRange(currentRow, 1, 1, 2).merge();
    headerRange.setBackground('#999999'); // Dark gray 1
    headerRange.setFontFamily('Calibri');
    headerRange.setFontSize(12);
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');

    // Merge A1:B1
    currentRow++;

    const seatStartRow = currentRow; // store start of seats

    for (let i = 1; i <= vehicle.totalSeats; i++) {
      const boosterPicture = `=IF(ISNUMBER(SEARCH(CHAR(8203), $C${currentRow})), 'Students x days'!$AB$1, "")`
      sheet.getRange(currentRow, 1).setValue(i); // seat number in column A
      sheet.getRange(currentRow,2).setValue(boosterPicture)

      if (i === vehicle.totalSeats - 1) {
        sheet.getRange(currentRow, 3).setValue("<HELPER SEAT>");
      } else if (i === vehicle.totalSeats) {
        sheet.getRange(currentRow, 3).setValue("<DRIVER SEAT>");
      }

      currentRow++;
    }
    const seatEndRow = currentRow - 1;

    // Aligment of sheet, columns and rows

    // Center align columns A, B, E, F, G
    sheet.getRange(seatStartRow, 1, seatEndRow - seatStartRow + 1, 1).setHorizontalAlignment("center"); // A
    sheet.getRange(seatStartRow, 2, seatEndRow - seatStartRow + 1, 1).setHorizontalAlignment("center"); // B
    sheet.getRange(seatStartRow, 5, seatEndRow - seatStartRow + 1, 3).setHorizontalAlignment("center"); // E, F, G

    // Left align columns C, D, H
    sheet.getRange(seatStartRow, 3, seatEndRow - seatStartRow + 1, 2).setHorizontalAlignment("left");  // C, D
    sheet.getRange(seatStartRow, 8, seatEndRow - seatStartRow + 1, 1).setHorizontalAlignment("left");  // H


    // Apply conditional formatting for Helper and Driver seats in column C
    const rules = sheet.getConditionalFormatRules();
    const boosterRange = sheet.getRange(`H${seatStartRow}:H${seatEndRow}`);
    const timeRange = sheet.getRange(`G${seatStartRow}:G${seatEndRow}`);
    const yellow = "#FFF59D";
    const yellow2 = "#FFF176"; 
    const green = "#81C784";
    const red = "#E57373";
    const lightBlue = "#81D4FA";
    const greenFont = "#2E7D32"; 
    const redFont = "#C62828"; 

    const dismissalRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(ISNUMBER(G${seatStartRow}), G${seatStartRow} <> TIME(15,0,0))`)
      .setBackground(yellow) // Light Yellow 2
      .setRanges([timeRange])
      .build();


    // 2. Helper and Driver Seats (column C)
    const seatRange = sheet.getRange(seatStartRow, 3, seatEndRow - seatStartRow + 1); // Column C

    const helperRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("<HELPER SEAT>")
      .setBackground(yellow2)
      .setRanges([seatRange])
      .build();

    const driverRule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("<DRIVER SEAT>")
      .setBackground(yellow2)
      .setRanges([seatRange])
      .build();

    // 3. Booster Count Comparisons (row with vehicle header, columns C and H)
    const boosterH = `H${headerRow}`;
    const boosterC = `A${headerRow}`;

    const boosterEqualRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=VALUE(REGEXEXTRACT(${boosterH}, "\\d+$")) = VALUE(REGEXEXTRACT(${boosterC}, "\\d+$"))`)
      .setBackground(yellow)
      .setRanges([sheet.getRange(boosterH)])
      .build();

    const boosterGreaterRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=VALUE(REGEXEXTRACT(${boosterH}, "\\d+$")) > VALUE(REGEXEXTRACT(${boosterC}, "\\d+$"))`)
      .setBackground(green)
      .setRanges([sheet.getRange(boosterH)])
      .build();

    const boosterLessRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=VALUE(REGEXEXTRACT(${boosterH}, "\\d+$")) < VALUE(REGEXEXTRACT(${boosterC}, "\\d+$"))`)
      .setBackground(red)
      .setRanges([sheet.getRange(boosterH)])
      .build();

    // Add all rules
    rules.push(dismissalRule);
    rules.push(helperRule);
    rules.push(driverRule);
    rules.push(boosterEqualRule);
    rules.push(boosterGreaterRule);
    rules.push(boosterLessRule);

    // 1. Column M — Light Blue if M exists in L2:L111
    rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=COUNTIF($L$2:$L$111, M2) > 0')
      .setBackground(lightBlue)
      .setRanges([sheet.getRange('M2:M')])
      .build()
    );

    // 2. D1 Green if equals KD_MON!C2
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=D1=INDIRECT("KD_"${shortDay}"!C2")`)
        .setBackground(green)
        .setRanges([sheet.getRange('D1')])
        .build()
    );

    // 3. D1 Red if NOT equal
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=D1=INDIRECT("KD_"${shortDay}"!C2")`)
        .setBackground(red)
        .setRanges([sheet.getRange('D1')])
        .build()
    );

    // 4. E1:G1 — Green font if text is exactly "All kids are on the route! :)"
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("All kids are on the route! :)")
        .setFontColor(greenFont)
        .setRanges([sheet.getRange('E1:G1')])
        .build()
    );

    // 5. E1:G1 — Red font if text does NOT contain "All kids are on the route! :)"
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=ISERROR(SEARCH("All kids are on the route! :)", E1))')
        .setFontColor(redFont)
        .setRanges([sheet.getRange('E1:G1')])
        .build()
    );


    // Apply to sheet
    sheet.setConditionalFormatRules(rules);
    


    // Apply borders to full content area A–H
    const endRow = currentRow - 1;
    const borderedRange = sheet.getRange(startRow + 1, 1, endRow - startRow, 8); // A–H
    borderedRange.setBorder(true, true, true, true, true, true);
    
    sheet.getRange(currentRow, 1, 1, 8).clearContent();
    currentRow++;

  });
  sheet.getRange(currentRow-1,1).setValue("=CHAR(160)")
  
  
  
  // ui.alert("The 'test' sheet has been reorganized successfully!");
}


function isStudentDuplicate(sheet, studentName, currentRange) {
  var dataRange = sheet.getRange("C2:C" + (currentRange.getRow() - 1));
  var data = dataRange.getValues().flat();
  return data.includes(studentName);
}


const button1 = () => switching("button1");
const button2 = () => switching("button2");

function switching(name) {
  checkExpiry();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const lRowStudent = sheet.getRange("Y2").getValue();
  const range = sheet.getRange("A6:W" + lRowStudent);
    const drawings = sheet.getDrawings();
   //var lastRow = ss.getRange("Y2").getValues();
   drawings.forEach(d => d.setZIndex(d.getOnAction() == name ? 0 : 1));

  if (name === "button1") {
    range.sort(2);  
  } else if (name === "button2") {
    range.sort(1); 
  }

   const temp = ss.insertSheet();
   SpreadsheetApp.flush();
   ss.deleteSheet(temp);
   sheet.activate();
}


function getFirstEmptyRowFast(sheet, startLine = 1) {
  var lastRow = sheet.getLastRow();
  
  // Loop backward from the last row until finding the first non-empty row
  for (var i = lastRow; i >= startLine; i--) {
    var range = sheet.getRange(i, 1);
    if (range.isBlank()) {
      if (i === startLine || !sheet.getRange(i - 1, 1).isBlank()) {
        return i + 1; // Return the row after the last non-empty row
      }
    }
  }

  // If no blank row is found, return startLine
  return startLine;
}


const buttonSchool = () => switchingKd("buttonSchool");
const buttonName = () => switchingKd("buttonName");

function switchingKd(name) {
   checkExpiry();
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   var lRowStudent = getFirstEmptyRow(ss,5); 
   const sheet = ss.getActiveSheet();
   const drawings = sheet.getDrawings();
   //var lastRow = ss.getRange("Y2").getValues();
   drawings.forEach(d => d.setZIndex(d.getOnAction() == name ? 0 : 1));
   if (name == "buttonSchool"){
      ss.getRange("A5:H" + lRowStudent).sort(1);  
   } else if (name == "buttonName"){
      ss.getRange("A5:H" + lRowStudent).sort(3); 
   }
   const temp = ss.insertSheet();
   SpreadsheetApp.flush();
   ss.deleteSheet(temp);
   sheet.activate();
}

function getKidsSchools()
{
  const t0 = Date.now();
  // make a list of all schools per day
  var colDiaStart = 3; // coluna iniciando monday
  const ss = SpreadsheetApp.getActive();
  const student_days = ss.getSheetByName('Students x days'); // sheet Students x days
  const kids_schools = ss.getSheetByName('Kids and Schools') // sheet 
  const staffList = ss.getSheetByName('Staff') // sheet staff
  const lastStudentRow = student_days.getRange("Y2").getValue(); //getFirstEmptyRow(student_days,6)-1;
  
  const lastStaffRow = staffList.getRange("I2").getValue() //getFirstEmptyRow(staffList,2)-1
  var indexDia = {
    "Monday": colDiaStart,
    "Tuesday": colDiaStart+1,
    "Wednesday": colDiaStart+2,
    "Thursday": colDiaStart+3,
    "Friday": colDiaStart+4,
  };
  
  const dia = student_days.getRange(1,2).getValue().toString(); // dia escolhido para rota
  const data = student_days.getRange(1,1).getValue();
  var kdDia ={
      "Monday": "KD_MON",
      "Tuesday":"KD_TUE",
      "Wednesday":"KD_WED",
      "Thursday":"KD_THU",
      "Friday":"KD_FRI",
  }

  // Original time string
  var originalTime = '14:00:00';

  // Parse the time string into a Date object
  var parsedTime = new Date('2000-01-01T' + originalTime);

  // Extract hours and minutes
  var hours = parsedTime.getHours();
  var minutes = parsedTime.getMinutes();

  // Convert to 12-hour format
  var ampm = hours >= 12 ? 'p.m.' : 'a.m.';
  hours = hours % 12;
  hours = hours ? hours : 12; // Handle midnight (0 hours)

  // Format the time as a string
  var earlyDismissalTime = hours + ':' + (minutes < 10 ? '0' : '') + minutes + ' ' + ampm;

  //console.log(formattedTime); // Output: 2:00 pm

  //console.log(kdDia[dia])
  const kd = ss.getSheetByName(kdDia[dia]) 

  //console.log(dia);
  var rotaDia = ss.getSheetByName(dia); // nome da planilha do dia da rota
  //var qte_child = student_days.getRange(linha,1).getValue(); //contador 
  
  // limpa planilhas
  ["A5:H80"].forEach(range => {
    kids_schools.getRange(range).clear().setFontColor("black").clearDataValidations();
    kd.getRange(range).clear().setFontColor("black").clearDataValidations();
  });
  rotaDia.getRange("B:H").clearDataValidations();

  // kids_schools.getRange("A5:H80").clear();
  // kids_schools.getRange("A5:H80").setFontColor("black");
  // kids_schools.getRange("A5:H80").clearDataValidations();
  // kd.getRange("A5:H80").clear();
  // kd.getRange("A5:H80").setFontColor("black");
  // kd.getRange("A5:H80").clearDataValidations();
  // rotaDia.getRange("B:H").clearDataValidations();
  //console.log(ultimaLinhaRota);

  //
  kids_schools.getRange("A2").setValue(dia);
  kids_schools.getRange("G2").setValue(data);
  kd.getRange("A2").setValue(dia);
  kd.getRange("E3").setValue(data);


  const childrenRange = student_days.getRange("A6:W" + lastStudentRow)
  let children = childrenRange.getValues()
  const dropDv = childrenRange.getDataValidations();


  let statusDia = []; 
  let childrenPresents = [];
  let childrenAusentes = [];
  let childrenDropoff = [];
  let childrenN = [];
  let staff = [];
  let childrenPresentsKids_Schools = [];
  let childrenDropOffKids_Schools = [];
  //console.log(children.length);
  for (var i = 0; i<children.length; i++){
    //var nome = children[i] 

    let addrList = []; // default empty

    const rule = dropDv[i][12]; // Column M = index 12

    if (rule) {
      const t = rule.getCriteriaType();

      if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
        // Dropdown: “List of items”
        addrList = rule.getCriteriaValues()[0];

      } else if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
        // Dropdown: “List from a range”
        const rng = rule.getCriteriaValues()[0];
        addrList = rng.getValues().flat().filter(String);
      }
    }

    // If no dropdown (addrList empty), keep single value in column M
    if (addrList.length === 0) {
      addrList = [ children[i][12] || '' ]; // Wrap in array for consistency
    }

    // Store the array back into the row (in-memory update)
    children[i][12] = addrList;
    statusDia[i] = children[i][indexDia[dia]-1]; //student_days.getRange(linha,indexDia[dia]).getValues().toString();

    //break;
    const age = children[i][21]
    if (age < 9) {
          children[i][0] = children[i][0] + ' \u200B';
    }
    switch (statusDia[i]) {
      
      
      case "P" || "p" : 
        
        //console.log("child is Present today: " + children[i][0]);
        childrenPresents.push([children[i][0],"",children[i][1],children[i][7],children[i][9],'',children[i][11],children[i][21],children[i][12]]); // name, school name, school address, dismissal time, drop off, age
        childrenPresentsKids_Schools.push([children[i][0],"",children[i][1],children[i][12],children[i][13],"",children[i][11],""])
      break;
      case "ED" || "ed" || "Ed" || "eD" : // early dismissal
        childrenPresents.push([children[i][0],"",children[i][1],children[i][7],children[i][9],'',children[i][11],children[i][21],children[i][12]]); 
        childrenPresentsKids_Schools.push([children[i][0],"",children[i][1],children[i][12],children[i][13],"",children[i][11],""])
      break;
      case "E" || "e" :
        //console.log("child is Present today: " + children[i][0]);
        childrenPresents.push([children[i][0],"",children[i][1],children[i][7],children[i][9],'',children[i][11],children[i][21],children[i][12]]);
        childrenPresentsKids_Schools.push([children[i][0],"",children[i][1],children[i][12],children[i][13],"",children[i][11],""])
      break;
      case "D" || "d" :
        //console.log("child is Present today: " + children[i][0]);
        childrenDropoff.push([children[i][0],"",children[i][1],children[i][7],children[i][9],'TRUE',children[i][11],children[i][21],children[i][12]]);
        childrenDropOffKids_Schools.push([children[i][0],"",children[i][1],children[i][12],children[i][13],"",children[i][11],""])
      break;
      case "A" || "a":
        //console.log("child is Absent today: " + children[i][0]);
        childrenAusentes.push([children[i][0],children[i][1]]);
      break;
      case "N" || "n":
        //console.log("child is None today: " + children[i][0]);
        childrenN.push([children[i][0],children[i][1]]);
      break;
      
    }
  }


  childrenPresents.push(['','','','','','','','','']);

  /* ────────────────  BLOCK A : write rows + create dropdown  ──────────────── */
  // A-1  Build the rows we will actually WRITE (scalar in column I)
  const sheetRows = childrenPresentsKids_Schools.map(r => {
    const clone = [...r];                          // shallow copy
    if (Array.isArray(clone[3])) {                 // column I is index 8
      clone[3] = clone[3][0] || '';                // first addr (visible value)
    }
    return clone;
  });

  // A-2  Write those 9-column rows to Kids and Schools
  kids_schools
    .getRange(5, 1, sheetRows.length, 8)
    .setValues(sheetRows);

  // A-3  Build one DataValidation per row (only if >1 address)
  const dvRows = childrenPresentsKids_Schools.map(r => {
    const addrList = Array.isArray(r[3]) ? r[3] : [];
    return [
      addrList.length > 1
        ? SpreadsheetApp.newDataValidation()
            .requireValueInList(addrList, true)   
            .build()
        : null                                     // single address ⇒ no dropdown
    ];
  });

  // A-4  Apply the validations to column D
  kids_schools
    .getRange(5, 4, dvRows.length, 1)
    .setDataValidations(dvRows);
  /* ─────────────────────────────────────────────────────────────────────────── */



  var kidsDropOff = childrenPresents.length + 5
  // if (childrenDropoff.length >= 1){
  //   //if (childrenDropoff.length = 0)
  //   kids_schools.getRange(kidsDropOff,1,childrenDropoff.length,9).setValues(childrenDropoff);
  // } 
  
  /* ────────────────  BLOCK B : write DROP-OFF rows the same way  ──────────── */
if (childrenDropOffKids_Schools.length >= 1) {

  const dropoffRows = childrenDropOffKids_Schools.map(r => {
    const clone = [...r];
    if (Array.isArray(clone[3])) clone[3] = clone[3][0] || '';
    return clone;
  });

  kids_schools
    .getRange(kidsDropOff, 1, dropoffRows.length, 8)
    .setValues(dropoffRows);

  const dropDvRows = childrenDropOffKids_Schools.map(r => {
    const list = Array.isArray(r[3]) ? r[3] : [];
    return [
      list.length > 1
        ? SpreadsheetApp.newDataValidation()
            .requireValueInList(list, true)
            .build()
        : null
    ];
  });

  kids_schools
    .getRange(kidsDropOff, 4, dropDvRows.length, 1)
    .setDataValidations(dropDvRows);
}
/* ─────────────────────────────────────────────────────────────────────────── */


  childrenPresents.push(['<DRIVER SEAT>','','','','','','','','']);
  childrenPresents.push(['<HELPER SEAT>','','','','','','','','']);

 
  var rangeBlue = kids_schools.getRange("A5:A80"); // range to conditional format Blue if kid has in route
  var formulaBlue = '=COUNTIF(INDIRECT("\'" & $A$2 & "\'!C5:C200"), A5) > 0'; // Your conditional formatting formula
  var ruleBlueOnRoute = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formulaBlue)
    .setBackground("#34c9eb") 
    .setRanges([rangeBlue])
    .build();
  var rulesBlues = kids_schools.getConditionalFormatRules();
  rulesBlues.push(ruleBlueOnRoute);
  kids_schools.setConditionalFormatRules(rulesBlues);


  // var formulaYellow = '=AND(NOT(ISBLANK(E5)), NOT(E5=0), OR(E5 > TIME(15, 0, 0), E5 < TIME(15, 0, 0)))';
  // var rangeDismissalTime = kids_schools.getRange("E5:E150"); // range to conditional format yellow if the dismissal time is diffenent of 3 pm
  // var ruleYellowDismissal = SpreadsheetApp.newConditionalFormatRule()
  //   .whenFormulaSatisfied(formulaYellow)
  //   .setBackground("#ffff00") 
  //   .setRanges([rangeDismissalTime])
  //   .build();

  // var rulesYellow = kids_schools.getConditionalFormatRules();
  // rulesYellow.push(ruleYellowDismissal);
  // kids_schools.setConditionalFormatRules(rulesYellow);


  // Apply formatting on sheet above
  //console.log(childrenPresents.length);
  const finalLineKids = childrenPresents.length + childrenDropoff.length - 3
  var rangeKids = kids_schools.getRange(5, 1, finalLineKids + childrenDropoff.length,7);
  rangeKids.setBorder(true, true, true, true, true, true);

  kd.getRange(5,1,childrenPresents.length,9).setValues(childrenPresents);


  var absentInitialLine = childrenPresents.length + childrenDropoff.length + 5

  if (childrenAusentes.length > 0 ) {
    kids_schools.getRange("F2").setValue(childrenAusentes.length);
    kd.getRange("E2").setValue(childrenAusentes.length);
    //
    kids_schools.getRange(absentInitialLine,2).setValue("<< ABSENTS >>");
    kd.getRange(absentInitialLine,2).setValue("<< ABSENTS >>");
 
    //l_out++;
    kids_schools.getRange(absentInitialLine+1,2,childrenAusentes.length,2).setValues(childrenAusentes);
    kids_schools.getRange(absentInitialLine+1,2,childrenAusentes.length,4).setFontColor("red");
    //
    kd.getRange(absentInitialLine+1,2,childrenAusentes.length,2).setValues(childrenAusentes);
    kd.getRange(absentInitialLine+1,2,childrenAusentes.length,4).setFontColor("red");
    //
    
  } else
  {
    kids_schools.getRange("F2").setValue(0);
    kd.getRange("E2").setValue(0);
  }

  staff = staffList.getRange(2,1,lastStaffRow,3).getValues();
  //console.log(staff)
  //console.log(firstBlankRowKids);
  
  let partRangeHelpers = [];
  var p = 0;
  for (var i = 0; i<staff.length; i++){
    if (staff[i][2] == "Present" ){
      partRangeHelpers.push(staff[i][0])
    }
  }
  let partRangeChildrenPresents = [];
  for (var p = 0; p<childrenPresents.length; p++){
    partRangeChildrenPresents.push(childrenPresents[p][0])
  }

  var lastLineChild = childrenPresents.length + 1
  var partRuleHelpers = SpreadsheetApp.newDataValidation().requireValueInList(partRangeHelpers,false).build();
  var partRuleChildrenPresent = SpreadsheetApp.newDataValidation().requireValueInList(partRangeChildrenPresents,false).build();
  
  // fill with require check box data validation to confirm presence at day and put conditional formating if its was true (green)
  var valid_rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  //kids_schools.getRange("E5:E" + lastLineChild).setDataValidation(partRuleHelpers)
  //kd.getRange("D5:D" + lastLineChild).setDataValidation(partRuleHelpers)
  kids_schools.getRange("F5:F" + lastLineChild).setDataValidation(valid_rule).setValue("TRUE")
  kids_schools.getRange("G5:G" + lastLineChild).setDataValidation(valid_rule)
  kd.getRange("F5:F" + lastLineChild).setDataValidation(valid_rule).setValue("TRUE")
  
  if (childrenDropOffKids_Schools.length > 0) {
    const firstDropOffLine = lastLineChild + 2
    const lastDropOffLine = lastLineChild + 1 + childrenDropOffKids_Schools.length
    kids_schools.getRange("B" + firstDropOffLine +":B" + lastDropOffLine).setValue("Drop off only")
    kids_schools.getRange("F" + firstDropOffLine +":F" + lastDropOffLine).setDataValidation(valid_rule).setValue("TRUE")
    kids_schools.getRange("G" + firstDropOffLine + ":G" + lastDropOffLine).setDataValidation(valid_rule)

  }
  //console.log(dia);
 
  
  var rangeKids = kids_schools.getRange("E5:G80"); // Adjust the range as needed
  // Apply formatting on sheet above

  var ruleTrue = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("True")
    .setFontColor("#32a852") 
    //.setBackground("#32a852") // Change to the desired background color
    .setRanges([rangeKids])
    .build();

  var ruleFalse = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("False")
    //.setBackground("#de0917") // Change to the desired background color
    .setFontColor("#de0917") 
    .setRanges([rangeKids])
    .build();


  var rules = kids_schools.getConditionalFormatRules();
  //var rulesFalse = kids_schools.getConditionalFormatRules();

  rules.push(ruleTrue);
  rules.push(ruleFalse);

  kids_schools.setConditionalFormatRules(rules);


  //kids_schools.setConditionalFormatRules(rulesFalse);
  
  //  
  showButtonSchool(kids_schools);


  let vehiclesRota = rotaDia.getRange("A1:H" + rotaDia.getLastRow()).getValues();
  var inicio = 1;
  var fim = inicio;

  for (var i = 0; i<vehiclesRota.length; i++) {
    //console.log(i)
    if (vehiclesRota[i][0] == 1) {
      var inicio = i + 1;
      var fim = getFirstEmptyRow(rotaDia,inicio) - 1;
      // Apply formatting on sheet above
      var range = rotaDia.getRange(inicio, 1, fim - inicio+1, 8);
      //var range = rotaDia.getRange(inicio, 1, fim, 8);
      range.setBorder(true, true, true, true, true, true);
      range.setFontFamily("Calibri");
      range.setFontSize(10);

      var lPresent = i;
      //console.log(lPresent);
      //rotaDia.getRange(lPresent,9).setFormula('=Staff!$I$1');
      //rotaDia.getRange(lPresent,10).setValue('Total of Km Route');
      //rotaDia.getRange(lPresent,11).setFormula('=sum(J'+ inicio +':J' + fim + ')*2');
      //rotaDia.getRange(lPresent+1,9).setFormula('=unique(H' + inicio +':H' + fim +')');
      //rotaDia.getRange(lPresent-2,8).setFormula('=' +kdDia[dia]+'!$D3');
      //rotaDia.getRange(lPresent-1,3).setFormula('="Boosters Need : " & COUNTIF(A' + inicio + ':B' + fim + ', "💺")'); //="Boosters Need : "& COUNTIF(A1:B19, "💺") //
      rotaDia.getRange(lPresent-1,3).setFormula('="Boosters Required : " & COUNTIF(C' + inicio + ':C' + fim +', "*" & CHAR(8203) & "*")'); // =COUNTIF(C5:C19, "*" & CHAR(8203) & "*")
      //console.log('inicio ', inicio)
      //console.log('fim ', fim)
      
      //rotaDataVeiculo.push(rotaDataGeral[i]);
    }  
     //console.log(fim)
    if (vehiclesRota[i][0] >= 1) {
      var lPresent = i + 1;
      rotaDia.getRange(lPresent,3).setDataValidation(partRuleChildrenPresent);
      rotaDia.getRange(lPresent,6).setDataValidation(partRuleHelpers);
      //console.log(kdDia[dia])
      // arrange the formulas to auto take the address and school name
      if (lPresent < fim){
        rotaDia.getRange(lPresent, 4).setFormula('=IF(AND($C'+lPresent+'<>"", $C'+lPresent+'<>"<DRIVER SEAT>", $C'+lPresent+'<>"<HELPER SEAT>"), VLOOKUP($C'+lPresent+', INDIRECT("\'' + kdDia[dia] + '\'!A5:E" & COUNTA(' + kdDia[dia] + '!$A$5:$D$140)), 3, FALSE), "")');
      }
      rotaDia.getRange(lPresent, 7).setFormula('=IF(AND($C'+lPresent+'<>"", $C'+lPresent+'<>"<DRIVER SEAT>", $C'+lPresent+'<>"<HELPER SEAT>"), VLOOKUP($C'+lPresent+', INDIRECT("\'' + kdDia[dia] + '\'!A5:E" & COUNTA(' + kdDia[dia] + '!$A$5:$E$140)), 5, FALSE), "")').setNumberFormat("hh:mm AM/PM");

      rotaDia.getRange(lPresent, 8).setFormula('=IF(AND($C'+lPresent+'<>"", $C'+lPresent+'<>"<DRIVER SEAT>", $C'+lPresent+'<>"<HELPER SEAT>"), VLOOKUP($C'+lPresent+', INDIRECT("\'' + kdDia[dia] + '\'!A5:E" & COUNTA(' + kdDia[dia] + '!$A$5:$E$140)), 4, FALSE), "")');
      //rotaDia.getRange(lPresent,6).setFormula('=if($B'+lPresent+'<>"","3:00 pm","")')
      rotaDia.getRange(lPresent,10).setFormula('=if(I'+lPresent+'<>"", value(SUBSTITUTE(GOOGLEMAPS_DISTANCE(I'+lPresent+',I'+(lPresent - 1) +',"driving"),"km","")),"")')
    } else if (vehiclesRota[i][0] == "SEATS" && i > 3 ){
      
      //console.log(i)
      
    }
  }
  //console.log(fim)

   kids_schools.getRange("B5:B" + lastLineChild).setFormula(
  "=IFERROR(VLOOKUP(A5," + dia + "!$C$5:$F$" + fim + ",4,FALSE), \"\")"); // set the helper to show at kids and schools
   rotaDia.getRange(fim+1,1).setFormula("=CHAR(160)")


  // if (childrenAusentes.length > 1) {
  //   rotaDia.getRange(ultimaLinhaRota, 3,ultimaLinhaRota+100,4).clear();
  //   rotaDia.getRange(ultimaLinhaRota,3).setValue("<< ABSENTS >>");
  //   rotaDia.getRange(ultimaLinhaRota + 1, 3, childrenAusentes.length, 2).setValues(childrenAusentes);
  //   rotaDia.getRange(ultimaLinhaRota + 1, 3, childrenAusentes.length,5).setFontColor("red");
  // }
  Logger.log('getKidsSchools(): ' + (Date.now() - t0) + ' ms');
}





function arrangeRoute() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actualSheetName = ss.getActiveSheet().getName();
  var sheet = ss.getSheetByName(actualSheetName);

  if (!["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"].includes(actualSheetName)) {
    return;
  }

var rotaOrder = sheet.getRange("A1:N" + sheet.getLastRow()).getValues();
  var inicio = 1;
  var fim;
  var currentVehicle;

  for (var i = 0; i < rotaOrder.length; i++) {
    var linha = i + 1;

    // Check if the current row indicates the start of a new vehicle
    if (rotaOrder[i][0] === "SEATS" || rotaOrder[i][0] === "COUNT") {
      // Process the previous vehicle
      if (currentVehicle !== undefined) {
        fim = i - 3;
        processSort(sheet, inicio, fim);
      }

      // Update current vehicle ID
      currentVehicle = rotaOrder[i-1][13];
      inicio = linha+1;
    }
    if (rotaOrder[i][0] === String.fromCharCode(160)) { // char(160) at the last line 
      fim = linha-1
      processSort(sheet, inicio, fim);
    }
    // Process the last vehicle if it's the last row
    // if (linha === rotaOrder.length) {
    //   fim = linha;
    //   processVehicle(sheet, inicio, fim);
    // }
  }
}


function processSort(sheet, inicio, fim) {
  //console.log(inicio,fim)
   var range = sheet.getRange("C" + inicio + ":H" + fim);
   range.sort([{ column: 5, ascending: true }]);
}




// function ordenarRota(){  
  
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var actualSheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
 
//   var sheet = ss.getSheetByName(actualSheetName);  
  
//   if (
//     sheet.getName() != "MONDAY" &&
//     sheet.getName() != "TUESDAY" &&
//     sheet.getName() != "WEDNESDAY" &&
//     sheet.getName() != "THURSDAY" &&
//     sheet.getName() != "FRIDAY"
//   ) {
//     return;
//   }
  
//   var inicio = 1;
//   //var lRota = 5;
//   //var nSeats = sheet.getRange(lRota,1).getValue();
//   var fim = inicio;
//   // buscando o range de cada veiculo
//   let rotaOrder = sheet.getRange("A1:H" + sheet.getLastRow()).getValues();
//   //let rotaDataVeiculo = [];
//   for (var i = 0; i<rotaOrder.length; i++) {
//     var linha = i+1
//     if (rotaOrder[i][0] == 1) {
//       inicio = i + 1;
//       //rotaDataVeiculo.push(rotaOrder[i]);
//     }  
//     if (rotaOrder[i][0] > 1) {
//       //inicio = i + 1;
//       if (rotaOrder[i][2] == ""){
//         sheet.getRange(linha,4).setValue("");
//           if (rotaOrder[i][2] != "<DRIVER SEAT>"){sheet.getRange(linha,6).setValue("");}
//         //sheet.getRange(linha,6).setValue("");
//       } 
//     }  
//     fim = getFirstEmptyRow(sheet,inicio) -1;
//     if (linha == fim) {
//       //console.log('teste');
//       range = sheet.getRange("C" + inicio + ":H" + fim);
//       range.sort([{column: 5, ascending: true}]);
//       //break;
//     } 
//   }
    
// }

function showVehicleExportDialog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getRange("A2:K" + sheet.getLastRow()).getValues();
  const actualSheetName = sheet.getName();

  if (!["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"].includes(actualSheetName)) return;

  const vehicleList = [];
  let currentVehicle = null;
  let hasKid = false;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

  if (row[3] === "Vehicle") {
    if (currentVehicle && hasKid) {
      vehicleList.push(currentVehicle);
    }

    const nextRowType = data[i + 1]?.[0]; // Column A: SEATS or COUNT
    currentVehicle = nextRowType === "COUNT" ? "WALKING" : row[4]?.trim(); // Column G = plate
    hasKid = false;
    continue;
  }

    const seatNumber = row[0];
    const childName = row[2];

    if (
      typeof seatNumber === "number" &&
      childName &&
      typeof childName === "string" &&
      !childName.trim().toUpperCase().includes("SEAT")
    ) {
      hasKid = true;
    }
  }

  if (currentVehicle && hasKid) {
    vehicleList.push(currentVehicle);
  }

  if (vehicleList.length === 0) {
    SpreadsheetApp.getUi().alert("No vehicles with kids found.");
    return;
  }

//  saveExportRouteForVehicles("WALKING");

  const template = HtmlService.createTemplateFromFile("VehicleExportDialog");
  template.vehicleList = vehicleList;
  const html = template.evaluate().setWidth(400).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, "Select Vehicles to Export");
}


function saveExportRouteForVehicles(selectedVehicleNames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  if (!["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"].includes(sheetName)) return;

  // console.log(selectedVehicleNames)

  const data = sheet.getRange("A2:K" + sheet.getLastRow()).getValues();
  const dateOfRoute = data[0][7];

  const vehicleSheet = ss.getSheetByName("Vehicles");
  const vehicleData = vehicleSheet.getRange("A2:B" + vehicleSheet.getLastRow()).getValues(); // name + seats + plate
  const vehicleInfo = {}; // name → { seats, plate }

// Populate info from sheet
vehicleData.forEach(([name, seats, , , , , plate]) => {
  if (!name) return;
  vehicleInfo[name.trim()] = {
    seats: parseInt(seats, 10),
    plate: plate?.trim() || ""
  };
});

// ✅ Manually add "Walking" route
vehicleInfo["WALKING"] = {
  seats: 4,
  plate: ""
};

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    if (row[3] === "Vehicle") {
      const vehicleName = row[4]?.trim(); // Column C = vehicle name
      if (!selectedVehicleNames.includes(vehicleName)) continue;

      const info = vehicleInfo[vehicleName];
      const seats = info?.seats || 0;
      if (seats === 0) continue;

      const blockStartRow = i + 3;
      const blockEndRow = blockStartRow + seats - 1;
      const block = data.slice(blockStartRow - 1, blockEndRow);

      const driverRow = block.find(r => String(r[2]).toUpperCase().includes("<DRIVER SEAT>"));
      const driverName = driverRow ? driverRow[3] : "";

      const helperNames = [...new Set(
        block.map(r => r[5]).filter(h => h && h !== driverName)
      )];


      // 👇 Activate the correct range BEFORE exporting
      const range = sheet.getRange("A" + (blockStartRow - 2) + ":H" + (blockEndRow + 1));
      range.activate();

      let title = `Route ${vehicleName}`;
      if (driverName) title += ` (Driver - ${driverName})`;
      if (helperNames.length > 0) {
        title += ` (Helper${helperNames.length > 1 ? "s" : ""} - ${helperNames.join(", ")})`;
      }
      title += ` - ${formattedDate(dateOfRoute)} - ${sheetName}`;

      console.log(title);

      exportPartAsPDF(title); // ✅ Uses current active range
    }
  }

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput('<p>Export Successful.</p>').setWidth(300).setHeight(80),
    'Success'
  );

  sheet.getRange("A5").activate(); // Reset active cell
}



function saveExportRoute() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sheetName = sheet.getName();
  if (sheetName != "MONDAY" && sheetName != "TUESDAY" && sheetName != "WEDNESDAY" && sheetName != "THURSDAY" && sheetName != "FRIDAY") {
    return;
  }

  var lastRow = sheet.getLastRow() + 1;

  let rotaDiaVeiculo = sheet.getRange("A2:K" + lastRow).getValues();

  let inicio = [], fim = [], vehicles = [], drivers = [], iniCount = 0;
  const helpers = {}; // Create an object to store helpers for each vehicle
  const dateOfRoute = rotaDiaVeiculo[0][7];

  for (var i = 0; i < rotaDiaVeiculo.length; i++) { // loop for passing all the rotaDia sheet
    if (rotaDiaVeiculo[i][0] == 1 && rotaDiaVeiculo[i][2] != "") {

      
      inicio.push(i + 1);
      fim.push(getFirstEmptyRow(sheet, inicio[iniCount]) - 1);
      vehicles.push(rotaDiaVeiculo[i - 2][4]);

      var inicioHelpers = inicio[iniCount] - 1; // Correct index for helpers
      var fimHelpers = fim[iniCount] - 2; // Correct index for helpers
      var helperNames = [];
      
      // Iterate through the range to collect helper names
      for (var j = inicioHelpers; j <= fimHelpers; j++) {
        var helperName = rotaDiaVeiculo[j][5]; // Extract name from 5th column
        if (helperName && !helperNames.includes(helperName)) {
          helperNames.push(helperName); // Add name if it's not already in the array
        }
      }

      // Save unique helpers for the vehicle in the helpers object
      helpers[vehicles[iniCount]] = helperNames;

      // Remove duplicates across all buses (optional, if needed globally unique)
      helpers[vehicles[iniCount]] = [...new Set(helperNames)];

      iniCount++;
    }

    if (rotaDiaVeiculo[i][0] === "SEATS" && i > 2) {
      drivers.push(rotaDiaVeiculo[i - 4][3]);
    }

    if (rotaDiaVeiculo[i][0] === String.fromCharCode(160)) { // char(160) at the last line
      drivers.push(rotaDiaVeiculo[i - 1][3]);
    }
  }

  // Now, remove drivers that are also in the helpers list
  for (var vehicle in helpers) {
    var vehicleHelpers = helpers[vehicle];
    helpers[vehicle] = vehicleHelpers.filter(helper => !drivers.includes(helper)); // Remove drivers from helpers
  }

  // Output unique helper names for each vehicle
  //console.log(helpers);

  // Now, export PDF for each vehicle with its corresponding helpers
  for (var j = 0; j < inicio.length; j++) { // loop for how many times needs to save at pdf
    inicio[j] = inicio[j] - 2;
    var selectRanges = sheet.getRange("A" + inicio[j] + ":H" + fim[j]).activate();

    var pfdTitle = " Route " + vehicles[j];

    if (drivers[j] && drivers[j] !== undefined && drivers[j] !== null) {
       pfdTitle += " (DRIVER - " + drivers[j] + ")";
    }

    pfdTitle += (helpers[vehicles[j]].length >  0 ?  " (Helper" + (helpers[vehicles[j]].length > 1 ? "s" : "") + " - " + helpers[vehicles[j]].join(", ")  + ") - " : " - ");

    pfdTitle += formattedDate(dateOfRoute) + " - " + sheet.getName();

    // console.log(pfdTitle)

    exportPartAsPDF(pfdTitle);

    
  }
  const htmlOutput = HtmlService
    .createHtmlOutput('<p><p>') // <a href="' + pdfFile.getUrl() + '" target="_blank">' + fileName + '</a></p>')
    .setWidth(300)
    .setHeight(80);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Successful');
  selectRanges = sheet.getRange("A5").activate();
}


// function saveExportRoute() 
// {
//   var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//   var sheetName = sheet.getName();
//   if (sheetName != "MONDAY" &&  sheetName != "TUESDAY" && sheetName != "WEDNESDAY" && sheetName != "THURSDAY" && sheetName != "FRIDAY"){
//     return;
//   }
//   var lastRow = sheet.getLastRow()+1;

//   let rotaDiaVeiculo = sheet.getRange("A2:K" + lastRow).getValues();
//   let inicio = [];
//   let fim = [];
//   let vehicles = [];
//   let drivers = [];
//   var helpers = [];
//   var iniCount = 0;
//   var dateOfRoute = rotaDiaVeiculo[0][7];


//   for (var i = 0; i<rotaDiaVeiculo.length; i++){ // loop for pass all the rotadia sheet
//     if (rotaDiaVeiculo[i][0] == 1 && rotaDiaVeiculo[i][2] != "") {
//         inicio.push(i + 1);
//         fim.push(getFirstEmptyRow(sheet,inicio[iniCount]) - 1);
//         vehicles.push(rotaDiaVeiculo[i-2][4])
//         var inicioHelpers = inicio - 1
//         var fimHelpers = fim - 2
//         var helperNames = [];
//         // Iterate through the range to collect helper names
//         for (var j = inicioHelpers; j <= fimHelpers; j++) {
//           var helperName = rotaDiaVeiculo[j][5]; // Extract name from 5th column
//           if (helperName && !helperNames.includes(helperName)) {
//           helperNames.push(helperName); // Add name if it's not already in the array
//         }
//         helpers.push(...helperNames);
//     }
//     // Remove duplicates across all buses (optional, if needed globally unique)
//     helpers = [...new Set(helpers)];


//     iniCount++;
//   } 
 
//     if (rotaDiaVeiculo[i][0] === "SEATS" && i > 2) {
//       drivers.push(rotaDiaVeiculo[i-4][3])
//     }


//     if (rotaDiaVeiculo[i][0] === String.fromCharCode(160)) { // char(160) at the last line 
//      drivers.push(rotaDiaVeiculo[i-1][3])
//     }

//   }
//     // Output unique helper names
//     console.log(helpers);
 
  
//   for (var j = 0; j< inicio.length; j++){ //loop for how many times nees to save at pdf
//     inicio[j] = inicio[j] - 2
//     var selectRanges = sheet.getRange("A" + inicio[j] + ":H" + fim[j]).activate();

//     var pfdTitle = " Route ";

//     if (drivers[j] && drivers[j] !== undefined && drivers[j] !== null) {
//       pfdTitle += "(DRIVER - " + drivers[j] + ") - "+ vehicles[j] + " - " 
//     }
//     pfdTitle += formattedDate(dateOfRoute) + " - " + sheet.getName();

//     exportPartAsPDF(pfdTitle);
    
//   }
 
//   const htmlOutput = HtmlService 
//      .createHtmlOutput('<p><p>') // <a href="' + pdfFile.getUrl() + '" target="_blank">' + fileName + '</a></p>')
//      .setWidth(300)
//     .setHeight(80)
  
//   SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Successful')
//   selectRanges = sheet.getRange("A5").activate();
// }

function formattedDate(date) {
  var dateObject = new Date(date);
  var year = dateObject.getFullYear();
  var month = (dateObject.getMonth() + 1).toString().padStart(2, '0');
  var day = dateObject.getDate().toString().padStart(2, '0');
  return date = year + "-" + month + "-" + day;
}

function extractUniqueNames(rotaDiaVeiculo) {
  let names = rotaDiaVeiculo
    .map(row => row[5]) // Extract all values from column 5
    .filter(name => name && name.trim() !== ""); // Remove empty or invalid names
  
  // Use a Set to ensure uniqueness
  let uniqueNames = Array.from(new Set(names));
  
  return uniqueNames;
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
    
    var fileName = fileSuffix //spreadsheet.getName() + fileSuffix
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
// */
// function _exportBlob(blob, fileName, spreadsheet) {
//   blob = blob.setName(fileName)
//   //var folder = saveToRootFolder ? DriveApp : DriveApp.getFileById(spreadsheet.getId()).getParents().next()
//   var folder = DriveApp.getFoldersByName('PICKUP_PDF_VANCOUVER').next(); 
//   //var folder = DriveApp.getFolderById(folderId);
//   //console.log(folder);
//   var pdfFile = folder.createFile(blob);
//   //var pdfFile = folder.createFile(blob)
  
//   // Display a modal dialog box with custom HtmlService content.
//   /*const htmlOutput = HtmlService
//     .createHtmlOutput('<p><p>') // <a href="' + pdfFile.getUrl() + '" target="_blank">' + fileName + '</a></p>')
//     .setWidth(300)
//    .setHeight(80)
  
//   SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Successful')
//   */
// }

function _exportBlob(blob, fileName, spreadsheet) {
  // Set the name of the blob (file)
  blob = blob.setName(fileName);

  // Retrieve the folder name from cell K2 on the Staff sheet
  var sheet = spreadsheet.getSheetByName('Staff');
  var folderName = sheet.getRange('K2').getValue();

  // Get the parent folder of the spreadsheet
  var file = DriveApp.getFileById(spreadsheet.getId());
  var parentFolder = file.getParents().next();

  // Check if the subfolder exists inside the parent folder
  var folderIterator = parentFolder.getFoldersByName(folderName);
  var folder;

  if (folderIterator.hasNext()) {
    // Subfolder exists
    folder = folderIterator.next();
  } else {
    // Subfolder does not exist, create it
    folder = parentFolder.createFolder(folderName);
  }

  // Create the file in the subfolder
  var pdfFile = folder.createFile(blob);

  // Display a modal dialog box with custom HtmlService content (optional)
  /*
  const htmlOutput = HtmlService
    .createHtmlOutput('<p><a href="' + pdfFile.getUrl() + '" target="_blank">' + fileName + '</a> has been successfully exported.</p>')
    .setWidth(300)
    .setHeight(80);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Successful');
  */
}



function exportAsPDF() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var blob = _getAsBlob(spreadsheet.getUrl())
  //var folderId = "172OwTARFfdVKg-RG5baO0CqpW878WADs"
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
      + '&ktf=1'  // Embed fonts
      + '&fontName=Segoe UI Emoji'
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


function removeDuplicates(arr) 
{
        return arr.filter((item,
            index) => arr.indexOf(item) === index);
}    






const getBackgroundColor = () => {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange(ss.getActiveCell().getRow(),ss.getActiveCell().getColumn()).offset(0,-2);

  return range.getBackgroundColor();

};
function bgHex(cellAddress) {
 var mycell = SpreadsheetApp.getActiveSheet().getRange(cellAddress);
 var bghex = mycell.getBackground();
 return bghex;
}

function myBackgroundRanges(myRange,myTigger) {  
  return SpreadsheetApp.getActiveSpreadsheet()
          .getActiveSheet()
            .getRange(myRange)
              .getBackgrounds();}



/**
 * Calculate the driving distance (in meters) along a route.
 *
 * @param {"london","manchester","liverpool"} route
 *        Comma separated ordered list of two or more map
 *        waypoints to include in route. First point
 *        is 'origin', last is 'destination'.
 *
 * @customfunction
 */
function drivingDistance(route) {
  // From gist.github.com/mogsdad/e07d537ff06f444866c5
  // Adapted from developers.google.com/apps-script/quickstart/macros
  // If a range of cells is passed in, 'route' will be a two-dimensional array.
  // Test for an array, and if we have one, collapse it to a single array.
  if (route.constructor === Array) {
    var args = route.join(',').split(',');
  }
  else {
    // No array? Grab the arbitrary arguments passed to the function.
    args = arguments;
  }
  
  // Just one rule to a route - we need a beginning and an end
  if (args.length < 2) throw new Error( "Must have at least 2 waypoints." )
  
  // Pass our waypoints to getDirections_(). Tricky bit, this.
  var directions = getDirections_.apply(this, args);

  // We have our directions, grab the first route's legs
  var legs = directions.routes[0].legs;
  
  // Loop through all legs, and sum up distances
  var dist = 0;
  for (var i=0; i<legs.length; i++) {
    dist += legs[i].distance.value;
  }
  
  // Done - return the value in meters
  return dist;
}

/**
 * Use Maps service to get directions for a route consisting of an arbitrary
 * set of waypoints.
 */
function getDirections_(route) {
  // Just one rule to a route - we need a beginning and an end
  if (arguments.length < 2) throw new Error( "Must have at least 2 waypoints." )
  
  // Assume first point is origin, last is destination.
  var origin = arguments[0];
  var destination = arguments[arguments.length-1];
  
  // Build our route; origin + all midpoints + destination
  var directionFinder = Maps.newDirectionFinder();
  directionFinder.setOrigin(origin);
  for ( var i=1; i<arguments.length-1; i++ ) {
    directionFinder.addWaypoint(arguments[i]);
  }
  directionFinder.setDestination(destination);

  // Get our directions from Map service;
  // throw an error if no route can be calculated.
  var directions = directionFinder.getDirections();
  if (directions.routes.length == 0) {
    throw 'Unable to calculate directions between these addresses.';
  }
  return directions;
}

function getFirstEmptyRow(sheet, startLine = 1) {
  var column = sheet.getRange("A" + startLine + ":A").getValues(); // Adjust the column range as needed
  var rowIndex = 0;

  for (var i = 0; i < column.length; i++) {
    if (column[i][0] === "\u00A0"){
      rowIndex = i + startLine ;
      //console.log("linha com char(160)",rowIndex)
      break;

    }
    if (column[i][0] === "") {
      rowIndex = i + startLine;
      break;
    }
  }

  return rowIndex;
}





// function getFirstEmptyRow(sName,startLine = 1) {
  
//   //var spr = sName
//   //console.log(spr);
//   var cell = sName.getRange("A" + startLine);
//   var ct = 0;
//   while ( cell.offset(ct, 0).getValue() != "" ) {
//     ct++;
//   }
//   return (ct+startLine);
// }

function alertMessage(msg) {
  var result = SpreadsheetApp.getUi().alert(msg);
 
}

// Morphs a 1-d array into a 2-d array for use with Range.setValues([][])
function morphIntoMatrix(array) {

  // Create a new array and set the first row of that array to be the original array
  // This is a sloppy workaround to "morphing" a 1-d array into a 2-d array
  var matrix = new Array();
  matrix[0] = array;

  // "Sanitize" the array by erasing null/"null" values with an empty string ""
  for (var i = 0; i < matrix.length; i ++) {
    for (var j = 0; j < matrix[i].length; j ++) {
      if (matrix[i][j] == null || matrix[i][j] == "null") {
        matrix[i][j] = "";
      }
    }
  }
  return matrix;
}

function splitDate(dates) {
    if (dates != null)
    {
        var dates = dates.split(',');
        var xxx = dates.length;
        console.log(xxx);
        for (var i=0; i<xxx; i++)
        {
            dates[i] = dates[i];                    
        }
    }
    console.log(dates.join('\r\n'));
    return dates.join('\r\n');        
}

function showButtonSchool(sheetName) {
  //  const ss = SpreadsheetApp.getActiveSpreadsheet();
  //  const sheet = ss.getActiveSheet();
   const drawings = sheetName.getDrawings();

   // Set ZIndex for the "buttonSchool" drawing to 0
   drawings.forEach(d => d.setZIndex(d.getOnAction() == "buttonName" ? 0 : 1));
 SpreadsheetApp.flush();
   // If you want to make sure "buttonSchool" is always shown, you can use the following line:
   // drawings.forEach(d => d.setZIndex(0));
}

function ZODIAC(dateString) {
  var date = new Date(dateString);
  var month = date.getMonth() + 1; // Months are zero-based in JavaScript

  if ((month == 3 && date.getDate() >= 21) || (month == 4 && date.getDate() <= 19)) {
    return "Aries";
  } else if ((month == 4 && date.getDate() >= 20) || (month == 5 && date.getDate() <= 20)) {
    return "Taurus";
  } else if ((month == 5 && date.getDate() >= 21) || (month == 6 && date.getDate() <= 20)) {
    return "Gemini";
  } else if ((month == 6 && date.getDate() >= 21) || (month == 7 && date.getDate() <= 22)) {
    return "Cancer";
  } else if ((month == 7 && date.getDate() >= 23) || (month == 8 && date.getDate() <= 22)) {
    return "Leo";
  } else if ((month == 8 && date.getDate() >= 23) || (month == 9 && date.getDate() <= 22)) {
    return "Virgo";
  } else if ((month == 9 && date.getDate() >= 23) || (month == 10 && date.getDate() <= 22)) {
    return "Libra";
  } else if ((month == 10 && date.getDate() >= 23) || (month == 11 && date.getDate() <= 21)) {
    return "Scorpio";
  } else if ((month == 11 && date.getDate() >= 22) || (month == 12 && date.getDate() <= 21)) {
    return "Sagittarius";
  } else if ((month == 12 && date.getDate() >= 22) || (month == 1 && date.getDate() <= 19)) {
    return "Capricorn";
  } else if ((month == 1 && date.getDate() >= 20) || (month == 2 && date.getDate() <= 18)) {
    return "Aquarius";
  } else if ((month == 2 && date.getDate() >= 19) || (month == 3 && date.getDate() <= 20)) {
    return "Pisces";
  } else {
    return "";
  }
}
