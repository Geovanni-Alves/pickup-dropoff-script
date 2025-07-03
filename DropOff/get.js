function doGet() {
  var spreadsheetId = '1Q13xpsb1xgqd1UvXLXQ1Pa36mbFiX7_qoXHVqFO1fCI'; // Replace with your spreadsheet ID
  var dropOffPerDaySheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName('DROP-OFF LIST PER DAY');
  var sheetName = dropOffPerDaySheet.getRange("B2").getValue(); //'Tuesday'; // Replace with the name of your sheet
  
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  var linkSheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName("Links")
  var lastLine = sheet.getLastRow();
  var data = sheet.getRange("A2:N" + lastLine).getValues();
  var gateKeeper = linkSheet.getRange("A4").getValue();

  
  var lists = []; // An array to store lists of kids
  var currentList = null;
  var k = 1;

  for (var j = 0; j < data.length; j++) {
    if (data[j][0] === 1 && data[j][2] != "") {
      if (currentList) {
        lists.push(currentList); // Push the previous list into the array
      }
      currentList = {
        listName: "Vehicle: " + data[j-2][13],
        driver: data[j-2][9], // Capitalize the first letter
        helper: data[j-2][10] + ': ' + data[j-2][11], // Capitalize helper names
        kids: [] // An array to store child names in the current list
      };
    }
    
    if (currentList && data[j][2] !== "" && data[j][2] !== "CHILD") {
      currentList.kids.push([data[j][0], data[j][2], data[j][4]]);
      k++;
    }
  }

  var totalKids = k - 1;
  if (currentList) {
    lists.push(currentList); // Push the last list into the array (if any)
  }

  // Use a template literal to create the main structure
  let htmlOutput = `
  <html>
  <head>
    <style>
      .container {
        flex: 1;
        box-sizing: border-box;
        padding: 10px;
        border: 2px solid #ccc;
        border-radius: 20px;
        margin: 10px;
        font-size: 18px;
        width: 30%;
      }
      .header {
        font-size: 20px;
        font-weight: bold;
        text-align: center;
      }
      .info {
        font-weight: bold;
        text-align: center;
      }
      .kids-list {
        list-style-type: none;
        padding-left: 0;
        font-size: 28px;
      }
      .flex-container {
        display: flex;
        flex-wrap: wrap;
        justify-content: space-around;
      }
      .image-title {
        font-size: 40px;
        text-align: center;
        font-weight: bold;
        color: white;
        background-color: red;
        padding: 20px;
        margin-bottom: 20px;
        border-radius: 20px;
      }
    </style>
  </head>
  <body>
    <div class="image-title">
      DROP-OFF LIST - ${sheetName}
    </div>
    <div class="flex-container">
  `;

  // Loop through the lists and dynamically insert the data
  lists.forEach(list => {
    htmlOutput += `
    <div class="container">
      <div class="header">${list.listName}</div>
      <div class="driver-helper">
        <div style="display: flex;">
          <p style="margin-right: 10px; font-weight: bold;">Driver:</p>
          <p style="margin-right: 10px;">${list.driver}</p>
          <p style="font-weight: bold;"> ${list.helper}</p>
        </div>
      </div>
      <div class="info">List of Kids: </div>
      <ul class="kids-list">
    `;





  list.kids.forEach(childInfo => {
    // Clean up zero-width spaces and extra spaces
    const fullName = childInfo[1].replace(/\u200B/g, '').replace(/\s+/g, ' ').trim(); // Remove zero-width space and clean up extra spaces

    const nameParts = fullName.split(' '); // Split the name by spaces

    // Handle case when there are at least two parts (first and last name)
    const firstName = nameParts[0]; // Get the first name
    const lastName = nameParts.length > 1 ? nameParts[nameParts.length - 1] : ''; // Get the last name if available

    const formattedName = lastName ? `${firstName} ${lastName}` : firstName; // Combine first and last name

    // Build the HTML output
    htmlOutput += `
      <li>${childInfo[0]} - ${formattedName}</li>
    `;
  });





    htmlOutput += `
      </ul>
    </div>
    `;
  });

  htmlOutput += `
    </div> <!-- Close flex-container -->
    <div style="text-align: center;">
      <p style="margin-top: 50px; font-weight: bold; font-size: 30px;">Total of Kids: ${totalKids}</p>
      <p style="margin-top: 50px; font-weight: bold; font-size: 30px;">Drop-off Manager: ${gateKeeper}</p>
    </div>
  </body>
  </html>
  `;

  return HtmlService.createHtmlOutput(htmlOutput);
}


