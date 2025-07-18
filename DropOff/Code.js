/**
 * Drop-Off Route Management & Reporting System
 * -------------------------------------------
 * Author : Geovanni Estevam  
 * Version: 1.0
 * Date   : 2023-04-04
 *
 * Overview
 * --------
 * Google Apps Script that automates after-school drop-off logistics inside
 * Google Sheets.  Key features include:
 *   • Dynamic “Route Controls” menu on sheet open
 *   • Fast lookup of first blank rows for data insertion
 *   • Route block sorting / validation per vehicle
 *   • One-click Google Maps direction links for selected vehicles
 *   • PDF export of full routes and drop-off-time tables
 *   • Distance / duration look-ups, reverse geocoding, & caching
 *
 * Sheet / Column Assumptions
 * --------------------------
 *   • Day-specific tabs named MONDAY … FRIDAY
 *   • Vehicle blocks follow the COUNT / SEATS pattern
 *   • Helper & driver names in columns F and D (0-based relative to script)
 *
 * API & Quotas
 * ------------
 *   • Requires an active Google Maps API key
 *   • Direction, Distance & Geocode calls count against Maps quota
 *
 * License
 * -------
 * MIT – do anything you like, but keep this header and give credit.
 *
 * “Built by Geovanni Estevam. 
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
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  try {
    checkExpiry();

  ui.createMenu('Route Controls')
      .addItem('Arrange', 'arrangeRoute')
      .addItem('Create Google maps links', 'generateMapLinkWithDirections')
      .addItem('Export route to PDF', 'saveExportRoute')
      .addItem('Export Drop-Off time as PDF', 'ExportDropOffTime')
      // .addItem('Export Selected Drop-Offs to PDF',"showDropOffExportDialog")
      //.addSeparator()
      //.addSubMenu(ui.createMenu('Sub-menu')
      //    .addItem('Second item', 'menuItem2'))
      .addToUi();
  
  } catch (e) {
    // Expired — do not add Route Tools menu
    // Optionally, you can alert user here or just skip silently
    // ui.alert('The Route Tools menu is disabled because the script has expired.');
  }
  // var sheet = getWeekDayName();
  // var ss = SpreadsheetApp
  // ss.getActiveSpreadsheet().getSheetByName(sheet).activate();
  ui.createMenu('About…')
    .addItem('About this Sheet', 'showAboutDialog')
    .addToUi();
}

function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  var column = range.getColumn();
  var value = e.value;

  if (sheet.getName() === "MONDAY" 
    || sheet.getName() === "TUESDAY" 
    || sheet.getName() === "WEDNESDAY"
    || sheet.getName() === "THRUSDAY"
    || sheet.getName() === "FRIDAY"
    && column === 2) {
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

function firstBlankRow(sheetName, columnName, startRow) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var columnValues = sheet.getRange(columnName + startRow + ":" + columnName + sheet.getMaxRows()).getValues();
  
  for (var i = 0; i < columnValues.length; i++) {
    if (!columnValues[i][0]) {
      return startRow + i;
    }
  }

  return "No blank row found";
}

function generateMapLinkWithDirections() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (
    sheet.getName() != "MONDAY" &&
    sheet.getName() != "TUESDAY" &&
    sheet.getName() != "WEDNESDAY" &&
    sheet.getName() != "THURSDAY" &&
    sheet.getName() != "FRIDAY"
  ) {
    return;
  }

  //var apiKey = GOOGLE_MAPS_API_KEY 
  var baseLink = "https://www.google.com/maps/dir/";
  var startLocation = sheet.getRange("I2").getValue(); // Assuming starting address is in cell B2

  var start = 1;
  var end = start;
  let vehicle = [];
  let driver = [];
  var actualVehicle = [];
  var actualDriver = []; 
  var v = 0;
 
  // buscando o range de cada veiculo
  let rotaOrder = sheet.getRange("A1:N" + sheet.getLastRow()).getValues();

  var vehicleStatusRange = sheet.getRange("P3:Q"); // Assuming the range for vehicle status is P3:Q
  var vehicleStatusValues = vehicleStatusRange.getValues();

  var trueVehicles = [];
  for (var i = 0; i < vehicleStatusValues.length; i++) {
    if (vehicleStatusValues[i][1] === true) {
      trueVehicles.push(vehicleStatusValues[i][0]);
    }
  }

  for (var i = 0; i<rotaOrder.length; i++) {
    actualVehicle = [];
    actualDriver = []; 
    var linha = i + 1
    if (rotaOrder[i][0] == 1 && rotaOrder[i][2] != "") {
      //vehicle = [];
      //driver = [];
      vehicle = rotaOrder[i-2][13]
      driver = rotaOrder[i-2][9]
      start = i + 1;
      v++;
    }  
    end = getFirstEmptyRow(sheet,start) - 1;
    if (linha == end ) {
      actualVehicle = vehicle
      actualDriver = driver
      if (trueVehicles.includes(actualVehicle)) {
        var data = sheet.getRange("F" + start + ":F" + end).getValues(); // Assuming addresses are in column F


        // Get the starting location (assuming it's in cell B2)
        var mapLink = baseLink + encodeURIComponent(startLocation) + "/";

        var addressSet = new Set(); // Use a Set to keep track of unique addresses

        for (var a = 0; a < data.length; a++) { // Start from index 1 since the first row is the starting location
          var address = data[a][0];

          // Skip empty rows and check for duplicate addresses
          if (address.trim() !== "" && !addressSet.has(address)) {
            // Append the address to the link
            mapLink += encodeURIComponent(address) + "/";
            addressSet.add(address); // Add the address to the set to track duplicates
          }
        }
        mapLink += "?travelmode=driving";
        
        
        //console.log(mapLink)
        SpreadsheetApp.getUi().alert("Directions (" + actualVehicle + ") - DRIVER " + actualDriver +":\n" + mapLink);
        //console.log(actualDriver)

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


function getBestRoute() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actualSheetName = ss.getActiveSheet().getName();
  var sheet = ss.getSheetByName(actualSheetName);

  if (
    sheet.getName() != "MONDAY" &&
    sheet.getName() != "TUESDAY" &&
    sheet.getName() != "WEDNESDAY" &&
    sheet.getName() != "THURSDAY" &&
    sheet.getName() != "FRIDAY"
  ) {
    return;
  }

  var gracieBarraEnd = sheet.getRange("L2").getValue(); // Get the origin address in "latitude, longitude" format
  var enderecos = sheet.getRange("M5:M17").getValues(); // Get the list of coordinates

  // Ensure that gracieBarraEnd and enderecos are in the correct format with latitude and longitude
  var waypoints = [];

  for (var i = 0; i < enderecos.length; i++) {
    if (enderecos[i][0]) {
      // Check if the cell is not empty
      var coordinate = enderecos[i][0].toString();
      var latLng = coordinate.split(','); // Split the coordinate into latitude and longitude
      if (latLng.length === 2) {
        // Ensure it's a valid coordinate with both latitude and longitude
        waypoints.push({
          latitude: parseFloat(latLng[0]),
          longitude: parseFloat(latLng[1])
        });
      }
    }
  }

  // Structure the addresses array with origin and waypoints
  var origin = {
    latitude: parseFloat(gracieBarraEnd.split(',')[0]),
    longitude: parseFloat(gracieBarraEnd.split(',')[1])
  };

  var addresses = [origin, ...waypoints];
  console.log(addresses);
  // Call the Google Maps Routes API to calculate the route.
  var route = calculateRoute(addresses); // Get both route data and shareable URL

  // Log the best route.
  Logger.log('Best Route:');
  Logger.log(route.routeData); // Log the route data
  Logger.log('Shareable URL:');
  Logger.log(route.shareableUrl); // Log the shareable URL

  // You can now use route.routeData and route.shareableUrl as needed in your script.
}


function calculateRoute(addresses) {
  // Replace 'YOUR_GOOGLE_MAPS_API_KEY' with your actual Google Maps API key.
  var apiKey = GOOGLE_MAPS_API_KEY  //'AIzaSyBFaCGSLr8WImcQuCEBgHk0Bn5GeKm2E58';
  // Replace 'YOUR_GOOGLE_MAPS_API_KEY' with your actual Google Maps API key.
  
  // Prepare the request payload for the Google Maps Routes API.
  //console.log(addresses)
  var payload = {
    "origin": addresses[0].latitude + "," + addresses[0].longitude,
    "destination": addresses[addresses.length - 1].latitude + "," + addresses[addresses.length - 1].longitude,
    "mode": "driving",
    "avoid": ["tolls", "ferries"],
    "language": "en-US",
    "units": "imperial",
    "waypoints": addresses.slice(1, -1).map(function(coord) {
      return coord.latitude + "," + coord.longitude;
    })
  };


  // Create an array of intermediates using latitude and longitude
  var intermediates = addresses.slice(1, -1).map(function(coord) {
    return {
      "latitude": coord.latitude,
      "longitude": coord.longitude
      }
  });

  // Add the list of intermediates to the payload
  payload.intermediates = intermediates;

  // Make the request to the Google Maps Routes API.
  //var url = 'https://routes.googleapis.com/route-optimizer/v1/routeOptimization?key=' + apiKey;
  //var url = 'https://routes.googleapis.com/directions/v2:computeRoutes?key=' + apiKey;
  var url = 'https://maps.googleapis.com/maps/api/directions/json?key=' + apiKey;

  var headers = {
    'X-Goog-FieldMask': 'routes.duration,routes.distanceMeters,routes.polyline.encodedPolyline'
  };

  var options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'headers': headers
  };
  console.log(url, options)
  var response = UrlFetchApp.fetch(url, options);

  var responseData = JSON.parse(response.getContentText());
console.log(responseData)
  // Check if the API request was successful
  if (responseData && responseData.routes && responseData.routes.length > 0) {
    // Extract route information from responseData
    var route = responseData.routes[0];
    var distance = route.distanceMeters;
    var duration = route.duration;
    var encodedPolyline = route.polyline.encodedPolyline;

    // Construct the shareable URL for the route
    var origin = addresses[0];
    var destination = addresses[addresses.length - 1];
    var waypoints = addresses.slice(1, -1);

    var originString = origin.latitude + ',' + origin.longitude;
    var destinationString = destination.latitude + ',' + destination.longitude;
    var waypointsString = waypoints.map(function(coord) {
      return coord.latitude + ',' + coord.longitude;
    }).join('|');
    
    var shareableUrl = 'https://www.google.com/maps/dir/?api=1&origin=' + originString + '&destination=' + destinationString + '&waypoints=' + waypointsString + '&travelmode=driving';

    // Return the extracted route data and the shareable URL
    return { routeData: responseData, shareableUrl: shareableUrl };
  } else {
    // Handle the case where the response is unexpected or empty
    return 'Error: Unable to fetch route data from Google Maps API.';
  }
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
   range.sort([{ column: 4, ascending: true }]);
}


// function arrangeRoute() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var actualSheetName = ss.getActiveSheet().getName();
//   var sheet = ss.getSheetByName(actualSheetName);

//   if (!["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"].includes(actualSheetName)) {
//     return;
//   }

//   var rotaOrder = sheet.getRange("A1:N" + sheet.getLastRow()).getValues();
//   var inicio = 1;
//   var fim;
//   var currentVehicle;

//   for (var i = 0; i < rotaOrder.length; i++) {
//     var linha = i + 1;

//     // Check if the current row indicates the start of a new vehicle
//     if (rotaOrder[i][0] === "COUNT") {
//       // Process the previous vehicle
//       if (currentVehicle !== undefined) {
//         fim = i - 3;
//         processVehicle(sheet, inicio, fim);
//       }

//       // Update current vehicle ID
//       currentVehicle = rotaOrder[i-1][13];
//       inicio = linha+1;
//     }

//     // Process the last vehicle if it's the last row
//     // if (linha === rotaOrder.length) {
//     //   fim = linha;
//     //   processVehicle(sheet, inicio, fim);
//     // }
//   }
// }



// function processVehicle(sheet, inicio, fim) {
//   var range = sheet.getRange("C" + inicio + ":H" + fim);
//   range.sort([{ column: 4, ascending: true }]);
// }


// function ordenarRota() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var actualSheetName = ss.getActiveSheet().getName();
//   var sheet = ss.getSheetByName(actualSheetName);

//   if (!["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY"].includes(actualSheetName)) {
//     return;
//   }

// var rotaOrder = sheet.getRange("A1:N" + sheet.getLastRow()).getValues();
//   var inicio = 1;
//   var fim;
//   var currentVehicle;

//   for (var i = 0; i < rotaOrder.length; i++) {
//     var linha = i + 1;

//     // Check if the current row indicates the start of a new vehicle
//     if (rotaOrder[i][0] === "SEATS") {
//       // Process the previous vehicle
//       if (currentVehicle !== undefined) {
//         fim = i - 3;
//         processSort(sheet, inicio, fim);
//       }

//       // Update current vehicle ID
//       currentVehicle = rotaOrder[i-1][13];
//       inicio = linha+1;
//     }
//     if (rotaOrder[i][0] === String.fromCharCode(160)) { // char(160) at the last line 
//       fim = linha-1
//       processSort(sheet, inicio, fim);
//     }
//     // Process the last vehicle if it's the last row
//     // if (linha === rotaOrder.length) {
//     //   fim = linha;
//     //   processVehicle(sheet, inicio, fim);
//     // }
//   }
// }


// function processSort(sheet, inicio, fim) {
//   //console.log(inicio,fim)
//    var range = sheet.getRange("C" + inicio + ":H" + fim);
//    range.sort([{ column: 5, ascending: true }]);
// }


function isStudentDuplicate(sheet, studentName, currentRange) {
  var dataRange = sheet.getRange("B2:B" + (currentRange.getRow() - 1));
  var data = dataRange.getValues().flat();
  return data.includes(studentName);
}


function GetSheetName() {
return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

const getBackgroundColor = () => {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[2];
  var range = sheet.getRange(ss.getActiveCell().getRow(),ss.getActiveCell().getColumn()).offset(0,-2);

  return range.getBackgroundColor();

};

function myBackgroundRanges(myRange,myTigger) {  
  return SpreadsheetApp.getActiveSpreadsheet()
          .getActiveSheet()
            .getRange(myRange)
              .getBackgrounds();}

function getWeekDayName() {
  
  // get current date
  var exampleDate = new Date();
  //Logger.log('exampleDate is: ' + exampleDate);
  
  // get the weekday number from the current date
  var dayOfWeek = exampleDate.getDay();
  //Logger.log('weekday number is: ' + dayOfWeek);
  
  // use a 'switch' statement to calculate the weekday name from the weekday number
  switch (dayOfWeek) {
  case 0:
    day = "Sunday";
    break;
  case 1:
    day = "Monday";
    break;
  case 2:
     day = "Tuesday";
    break;
  case 3:
    day = "Wednesday";
    break;
  case 4:
    day = "Thursday";
    break;
  case 5:
    day = "Friday";
    break;
  case 6:
    day = "Saturday";
  
  
}
  
  // log the output to show the weekday name
  //Logger.log('weekday name is: ' + day);
  return day;
}

 


function formattedDate(date) {
  var dateObject = new Date(date);
  var year = dateObject.getFullYear();
  var month = (dateObject.getMonth() + 1).toString().padStart(2, '0');
  var day = dateObject.getDate().toString().padStart(2, '0');
  return date = year + "-" + month + "-" + day;
}



function getFirstEmptyRow(sName,startLine = 1) {
  
  //var spr = sName
  //console.log(spr);
  var cell = sName.getRange("A" + startLine);
  var ct = 0;
  while ( cell.offset(ct, 0).getValue() != "" ) {
    ct++;
  }
  return (ct+startLine);
}


/**
 * DROP-OFF ▸ open checkbox dialog that lists every BUS block
 * with at least one child.
 */
function showDropOffExportDialog() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const tab   = sheet.getName();

  if (!["MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY"].includes(tab)) {
    SpreadsheetApp.getUi().alert("Open a weekday drop-off tab first.");
    return;
  }

  const data        = sheet.getRange("A2:H" + sheet.getLastRow()).getValues();
  const vehicleList = [];
  let   currentBus  = null;
  let   hasKid      = false;

  data.forEach(row => {
    const colA = row[0];              // COUNT / SEATS / seat number
    const colB = row[1];              // header text “BUS 1 – PLATE …”
    const child = row[2];             // kid name

    // ── header rows ──────────────────────────────────────────────
    if (colA === "COUNT" || colA === "SEATS") {
      if (currentBus && hasKid) vehicleList.push(currentBus);
      currentBus = (colB || "").trim();
      hasKid     = false;
      return;
    }

    // ── kid rows (seat number is numeric) ───────────────────────
    if (typeof colA === "number" && child && !/SEAT/i.test(child)) {
      hasKid = true;
    }
  });
  if (currentBus && hasKid) vehicleList.push(currentBus);

  if (vehicleList.length === 0) {
    SpreadsheetApp.getUi().alert("No buses with kids found.");
    return;
  }

  /* reuse the same HTML file you already have */
  const tpl         = HtmlService.createTemplateFromFile("VehicleExportDialog");
  tpl.vehicleList   = vehicleList;
  tpl.dialogTitle   = "Select Buses to Export as PDF";
  tpl.buttonLabel   = "✅ Export Selected";
  tpl.callbackName  = "saveExportRouteForVehicles";
  SpreadsheetApp.getUi()
    .showModalDialog(tpl.evaluate().setWidth(400).setHeight(400),
                     "Select Buses");
}



/**
 * ------------------------------------------------------------------
 *  DROP-OFF  ▸  step 2: export the checked vehicles
 *             (same name as pick-up version so HTML works untouched)
 * ------------------------------------------------------------------
 */
function saveExportRouteForVehicles(selectedVehicleNames) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const tab   = sheet.getName();
  if (!["MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY"].includes(tab)) return;

  const data        = sheet.getRange("A2:K" + sheet.getLastRow()).getValues();
  const dateOfRoute = data[0][7];                             // col H = date

  /* ----------------------------------------------------------------
   * 1. Build a quick lookup  (Vehicles sheet:  name │ seats │ plate)
   * ---------------------------------------------------------------- */
  const vSheet = ss.getSheetByName("Vehicles");
  const vData  = vSheet.getRange("A2:G" + vSheet.getLastRow()).getValues();
  const vInfo  = {};                                          // name → {seats,plate}
  vData.forEach(([name,seats,,,,plate]) => {
    if (name) vInfo[name.trim()] = { seats:Number(seats)||0, plate:plate?.trim()||"" };
  });
  vInfo["WALKING"] = { seats:4, plate:"" };                   // always include walk-route

  /* ----------------------------------------------------------------
   * 2. Walk the sheet once, locate each vehicle block
   * ---------------------------------------------------------------- */
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (row[3] !== "Vehicle") continue;                       // not a header

    const vehicleName = data[i+1]?.[0] === "COUNT"
                      ? "WALKING"
                      : (row[6] || "").trim();                // col G = plate
    if (!selectedVehicleNames.includes(vehicleName)) continue;

    const { seats } = vInfo[vehicleName] || {};
    if (seats === 0) continue;

    const blockStart = i + 3;            // first kid row
    const blockEnd   = blockStart + seats - 1;
    const block      = data.slice(blockStart-1, blockEnd);

    const driverRow  = block.find(r => String(r[2]).toUpperCase().includes("<DRIVER SEAT>"));
    const driverName = driverRow ? driverRow[3] : "";

    const helperNames = [...new Set(
      block.map(r => r[5]).filter(h => h && h !== driverName)
    )];

    // Activate target range and export
    sheet.getRange("A" + (blockStart-2) + ":H" + (blockEnd+1)).activate();

    let title = `Drop-Off ${vehicleName}`;
    if (driverName) title += ` (Driver – ${driverName})`;
    if (helperNames.length)
      title += ` (Helper${helperNames.length>1?"s":""} – ${helperNames.join(", ")})`;
    title += ` – ${formattedDate(dateOfRoute)} – ${tab}`;

    exportPartAsPDF(title);
  }

  SpreadsheetApp.getUi()
    .showModalDialog(HtmlService.createHtmlOutput('<p>Export successful.</p>')
    .setWidth(300).setHeight(80), 'Done');

  sheet.getRange("A5").activate();                            // reset cursor
}





///
//
////. google maps funtions
//.  
///

// const GOOGLE_MAPS_API_KEY = 'AIzaSyBFaCGSLr8WImcQuCEBgHk0Bn5GeKm2E58';

const md5 = (key = '') => {
  const code = key.toLowerCase().replace(/\s/g, '');
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, code)
    .map((char) => (char + 256).toString(16).slice(-2))
    .join('');
};

const getCache = (key) => {
  return CacheService.getDocumentCache().get(md5(key));
};

// Store the results for 6 hours
const setCache = (key, value) => {
  const expirationInSeconds = 6 * 60 * 60;
  CacheService.getDocumentCache().put(md5(key), value, expirationInSeconds);
};

// function test_google  () {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var actualSheetName = ss.getActiveSheet().getName();
//   var sheet = ss.getSheetByName(actualSheetName);

//   var origin = '2440 Main St, Vancouver, BC V5T 3E2';
//   var destination = '480 Robson St #707, Vancouver, BC V6B 1S1';
//   var time = sheet.getRange("H3").getValue();
//   console.log(time)
//   //console.log('distance: ',GOOGLEMAPS_DISTANCE(origin,destination,'driving'));
//   console.log('distance: ',GOOGLEMAPS_DURATION(origin,destination));

// }

// const GOOGLEMAPS_DURATION = (origin, destination, mode = 'driving') => {
//   // Generate a unique cache key for this request
//   //const departureTimestamp = new Date(`1970-01-01T${departureTime}`).getTime() / 1000;
//   const cacheKey = ['duration', origin, destination, mode].join(',');
  

//   // Check if the result is in the cache
//   const cachedResult = getCache(cacheKey);
//   if (cachedResult !== null) {
//     return cachedResult;
//   }

//   const apiKeyParam = `key=${GOOGLE_MAPS_API_KEY}`;
//   const apiUrl = `https://maps.googleapis.com/maps/api/directions/json?${apiKeyParam}&origin=${encodeURIComponent(origin)}&destination=${encodeURIComponent(destination)}&mode=${mode}&departure_time=now`;
  
//   // &departure_time=${departureTimestamp}`;
//   const response = UrlFetchApp.fetch(apiUrl);
//   //console.log(response.getContentText());
//   const data = JSON.parse(response.getContentText());

//   if (data.status !== 'OK' || data.routes.length === 0) {
//     throw new Error('No route found!');
//   }

//   const { legs: [{ duration: { text: time } } = {}] } = data.routes[0];
//   // Extract the numerical part of the duration and convert it to a number
//   const numericalTime = parseFloat(time.replace(' mins', '').replace('min',''));
//   setCache(cacheKey, numericalTime);
//   return numericalTime;
// };


// const GOOGLEMAPS_DISTANCE = (origin, destination, mode = 'driving') => {
//   const cacheKey = ['distance', origin, destination, mode].join(',');

//   // Check if the result is in the cache
//   const cachedResult = getCache(cacheKey);
//   if (cachedResult !== null) {
//     return cachedResult;
//   }

//   const apiKeyParam = `key=${GOOGLE_MAPS_API_KEY}`;
//   const apiUrl = `https://maps.googleapis.com/maps/api/directions/json?${apiKeyParam}&origin=${encodeURIComponent(origin)}&destination=${encodeURIComponent(destination)}&mode=${mode}`;

//   const response = UrlFetchApp.fetch(apiUrl);
//   //console.log(response.getContentText());
//   const data = JSON.parse(response.getContentText());

//   if (data.status !== 'OK' || data.routes.length === 0) {
//     throw new Error('No route found!');
//   }
//  const { legs: [{ distance: { text: distance } } = {}] } = data.routes[0];

//   // Cache the result for future use
//   setCache(cacheKey, distance);
//   return distance

// };








/* Calculate the distance between two
 * locations on Google Maps.
 *
 * =GOOGLEMAPS_DISTANCE("NY 10005", "Hoboken NJ", "walking")
 *
 * @param {String} origin The address of starting point
 * @param {String} destination The address of destination
 * @param {String} mode The mode of travel (driving, walking, bicycling or transit)
 * @return {String} The distance in miles
 * @customFunction
 */
const GOOGLEMAPS_DISTANCE = (origin, destination, mode) => {
  const key = ['duration', origin, destination, mode].join(',');
  // Is result in the internal cache?
  const value = getCache(key);
  // If yes, serve the cached result
  if (value !== null) return value;
  const { routes: [data] = [] } = Maps.newDirectionFinder()
    .setOrigin(origin)
    .setDestination(destination)
    .setMode(mode)
    .getDirections();

  if (!data) {
    throw new Error('No route found!');
  }

  const { legs: [{ distance: { text: distance } } = {}] = [] } = data;
    setCache(distance);
  return distance;
};

/**
 * Use Reverse Geocoding to get the address of
 * a point location (latitude, longitude) on Google Maps.
 *
 * =GOOGLEMAPS_REVERSEGEOCODE(latitude, longitude)
 *
 * @param {String} latitude The latitude to lookup.
 * @param {String} longitude The longitude to lookup.
 * @return {String} The postal address of the point.
 * @customFunction
 */

const GOOGLEMAPS_REVERSEGEOCODE = (latitude, longitude) => {
  const { results: [data = {}] = [] } = Maps.newGeocoder().reverseGeocode(latitude, longitude);
  return data.formatted_address;
};

/**
 * Get the latitude and longitude of any
 * address on Google Maps.
 *
 * =GOOGLEMAPS_LATLONG("10 Hanover Square, NY")
 *
 * @param {String} address The address to lookup.
 * @return {String} The latitude and longitude of the address.
 * @customFunction
 */
const GOOGLEMAPS_LATLONG = (address) => {
  const { results: [data = null] = [] } = Maps.newGeocoder().geocode(address);
  if (data === null) {
    throw new Error('Address not found!');
  }
  const { geometry: { location: { lat, lng } } = {} } = data;
  return `${lat}, ${lng}`;
};

/**
 * Find the driving direction between two
 * locations on Google Maps.
 *
 * =GOOGLEMAPS_DIRECTIONS("NY 10005", "Hoboken NJ", "walking")
 *
 * @param {String} origin The address of starting point
 * @param {String} destination The address of destination
 * @param {String} mode The mode of travel (driving, walking, bicycling or transit)
 * @return {String} The driving direction
 * @customFunction
 */
const GOOGLEMAPS_DIRECTIONS = (origin, destination, mode = 'driving') => {
  const { routes = [] } = Maps.newDirectionFinder()
    .setOrigin(origin)
    .setDestination(destination)
    .setMode(mode)
    .getDirections();
  if (!routes.length) {
    throw new Error('No route found!');
  }
  return routes
    .map(({ legs }) => {
      return legs.map(({ steps }) => {
        return steps.map((step) => {
          return step.html_instructions.replace(/<[^>]+>/g, '');
        });
      });
    })
    .join(', ');
};



/**
 * Calculate the travel time between two locations
 * on Google Maps.
 *
 * =GOOGLEMAPS_DURATION("NY 10005", "Hoboken NJ", "walking")
 *
 * @param {String} origin The address of starting point
 * @param {String} destination The address of destination
 * @param {String} mode The mode of travel (driving, walking, bicycling or transit)
 * @return {String} The time in minutes
 * @customFunction
 */
const GOOGLEMAPS_DURATION = (origin, destination, mode = 'driving') => {
  const key = ['duration', origin, destination, mode].join(',');
  // Is result in the internal cache?
  const value = getCache(key);
  // If yes, serve the cached result
  if (value !== null) return value;
  const { routes: [data] = [] } = Maps.newDirectionFinder()
    .setOrigin(origin)
    .setDestination(destination)
    .setMode(mode)
    .getDirections();
  if (!data) {
    throw new Error('No route found!');
  }
  const { legs: [{ duration: { text: time } } = {}] = [] } = data;
  // Store the result in internal cache for future
   const numericalTime = parseFloat(time.replace(' mins', '').replace('min',''));
  setCache(key, numericalTime);
  return numericalTime;
};


const geoCodeAddress = (address) => {
  const key = [address].join(',');
  // Is result in the internal cache?
  const value = getCache(key);
  // If yes, serve the cached result
  if (value !== null) return value;
  var response = Maps.newGeocoder().geocode(address);
  
  if (response.status === 'OK') {
    var location = response.results[0].geometry.location;
    var lat = location.lat;
    var lng = location.lng;
    var latLng = lat + ", " + lng;
    setCache(latLng);
    return latLng
  } else {
    return "Error";
  }
}











