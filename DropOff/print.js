function doGet() {
  var spreadsheetId = '1Q13xpsb1xgqd1UvXLXQ1Pa36mbFiX7_qoXHVqFO1fCI'; // Replace with your spreadsheet ID
  var dropOffPerDaySheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName('DROP-OFF LIST PER DAY');
  var sheetName = dropOffPerDaySheet.getRange("B2").getValue(); 
  
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  var linkSheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Links");
  var lastLine = sheet.getLastRow();
  var data = sheet.getRange("A2:N" + lastLine).getValues();
  var gateKeeper = linkSheet.getRange("A4").getValue();

  var lists = [];
  var currentList = null;
  var k = 1;

  for (var j = 0; j < data.length; j++) {
    if (data[j][0] === 1 && data[j][2] != "") {
      if (currentList) {
        lists.push(currentList);
      }
      currentList = {
        listName: "Vehicle: " + data[j-2][13],
        driver: data[j-2][9],
        helper: data[j-2][10] + ': ' + data[j-2][11],
        kids: []
      };
    }

    if (currentList && data[j][2] !== "" && data[j][2] !== "CHILD") {
      currentList.kids.push([data[j][0], data[j][2], data[j][4]]);
      k++;
    }
  }

  var totalKids = k - 1;
  if (currentList) {
    lists.push(currentList);
  }

  // Return the HTML optimized for printing
  let htmlOutput = `
  <html>
  <head>
    <style>
      body { 
        font-family: Arial, sans-serif; 
        font-size: 12px; /* Smaller font size */
      }
      .container {
        padding: 3px; /* Reduced padding */
        border: 1px solid #ccc; /* Thinner border */
        border-radius: 5px; /* Smaller border radius */
        margin: 3px; /* Reduced margin */
        font-size: 12px; /* Smaller font size */
      }
      .header {
        font-size: 14px; /* Smaller header font */
        font-weight: bold;
        text-align: center;
      }
      .kids-list {
        list-style-type: none;
        padding-left: 0;
        font-size: 18px; /* Smaller font for kids' names */
      }
      .grid-container {
        display: grid;
        grid-template-columns: repeat(3, 1fr); /* Three columns */
        gap: 3px; /* Less space between columns */
      }
      .image-title {
        font-size: 30px; /* Slightly smaller title */
        text-align: center;
        font-weight: bold;
        color: white;
        background-color: red;
        padding: 10px; /* Reduced padding */
        margin-bottom: 10px; /* Reduced bottom margin */
        border-radius: 10px;
      }
      .print-btn {
        font-size: 16px; /* Smaller button font */
        padding: 5px 10px; /* Reduced padding */
        background-color: green;
        color: white;
        border: none;
        cursor: pointer;
        border-radius: 5px;
      }
      @media print {
        .print-btn {
          display: none; 
        }
      }
    </style>
  </head>
  <body>
    <div class="image-title">
      DROP-OFF LIST - ${sheetName}
    </div>
    <button class="print-btn" onclick="window.print()">Print</button> 
    <div class="grid-container">
  `;

  lists.forEach(list => {
    htmlOutput += `
    <div class="container">
      <div class="header">${list.listName}</div>
      <div class="driver-helper">
        <div style="display: flex;">
          <p style="margin-right: 5px; font-weight: bold;">Driver:</p>
          <p style="margin-right: 5px;">${list.driver}</p>
          <p style="font-weight: bold;"> ${list.helper}</p>
        </div>
      </div>
      <div class="info">List of Kids: </div>
      <ul class="kids-list">
    `;

    list.kids.forEach(childInfo => {
      const fullName = childInfo[1].replace(/\u200B/g, '').replace(/\s+/g, ' ').trim();
      const nameParts = fullName.split(' ');
      const firstName = nameParts[0];
      const lastName = nameParts.length > 1 ? nameParts[nameParts.length - 1] : '';
      const formattedName = lastName ? `${firstName} ${lastName}` : firstName;

      htmlOutput += `<li>${childInfo[0]} - ${formattedName}</li>`;
    });

    htmlOutput += `
      </ul>
    </div>
    `;
  });

  htmlOutput += `
    </div> <!-- Close grid-container -->
    <div style="text-align: center;">
      <p style="margin-top: 20px; font-weight: bold; font-size: 20px;">Total of Kids: ${totalKids}</p>
      <p style="margin-top: 20px; font-weight: bold; font-size: 20px;">Drop-off Manager: ${gateKeeper}</p>
    </div>
  </body>
  </html>
  `;

  return HtmlService.createHtmlOutput(htmlOutput);
}
