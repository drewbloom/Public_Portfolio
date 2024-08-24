function onOpen() {
  DocumentApp.getUi().createMenu('Upload Sheet')
      .addItem('Generate Upload Sheet', 'transferToUploadSheet')
      .addToUi();
}

function transferToUploadSheet() {
  const ui = DocumentApp.getUi();
  const response = ui.alert('Are you sure?', 'You should only create an upload sheet if all questions have been reviewed and are ready for upload into Aqueduct.  Are you sure you want to proceed?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    makeModal('Beginning process. Reading Google Doc for tables...');
    const doc = DocumentApp.getActiveDocument();
    const title = doc.getName(); 
    const tablePairs = getTablePairsFromDoc(doc);
    
    if (tablePairs.length === 0) {
      DocumentApp.getUi().alert('No valid table pairs found in the document.');
      return;
    }

    makeModal('Creating new Google Sheet');
    const sheet = createNewSheet('UPLOAD SHEET: ' + title);
    addHeadersToSheet(sheet);
    makeModal('Writing data to Google Sheet...');
    populateSheetWithPairs(sheet, tablePairs);
    setRowHeight(sheet)
    openSheetInBrowser(sheet);
  } else {
    Logger.log('User decided not to create the sheet. Terminating process')
    ui.alert('Process canceled');
  }
}

function getTablePairsFromDoc(doc) {
  const body = doc.getBody();
  const tablePairs = [];
  
  let currentPair = [];
  const numChildren = body.getNumChildren();
  Logger.log('Number of child elements in the document: %s', numChildren);

  for (let i = 0; i < numChildren; i++) {
    const element = body.getChild(i);
    
    if (element.getType() === DocumentApp.ElementType.TABLE) {
      const table = element.asTable();
      logTableContents(table);

      if (table.getRow(0).getCell(0).getText().trim() === "Question ID") {
        if (currentPair.length > 0) {
          tablePairs.push(currentPair);
          currentPair = [];
        }
        currentPair.push(table); 
      } 
      
      else if (table.getRow(0).getCell(0).getText().trim() === "Vignette:") {
        if (currentPair.length === 1) {
          currentPair.push(table); 
        }
      } else {
        Logger.log('Table does not match "Question ID" or "Vignette:" at child index %s', i);
      }
    }
  }

  if (currentPair.length === 2) {
    tablePairs.push(currentPair);
  }
  Logger.log('Number of table pairs found: %s', tablePairs.length);
  return tablePairs;
}

function createNewSheet(sheetTitle) {
  const ss = SpreadsheetApp.create(sheetTitle);
  return ss.getSheets()[0]; 
}

function addHeadersToSheet(sheet) {
  const headers = [
    "Case [B]",
    "Vignette [D]",
    "Stem [E]",
    "Image [F]",
    "Correct Answer [G]",
    "Choice_A [H]",
    "Choice_B [I]",
    "Choice_C [J]",
    "Choice_D [K]",
    "Choice_E [L]",
    "Item Validity Status [M]",
    "Aquifer Learning Objective [N]",
    "Learning Objective UID [O]",
    "Teaching Point [P]",
    "Teaching Point UID [Q]",
    "System [S]",
    "Question Use (Calibrate or PracticeSmart)",
    "Answer explanation (optional)",
    "Clinical Focus (required)",
    "Clinical Discipline - multiselect (required)",
    "Final Diagnosis (optional)",
    "Clinical Excellence Topics (optional)",
    "Patient age (optional)",
    "Clinical Location (optional)"
  ];
  sheet.appendRow(headers);
}

function populateSheetWithPairs(sheet, tablePairs) {
  for (let i = 0; i < tablePairs.length; i++) {
    const firstTable = tablePairs[i][0];
    const secondTable = tablePairs[i][1];
    const { rowData, errors } = extractTablePairData(firstTable, secondTable);

    const rowPos = sheet.getLastRow() + 1;
    if (rowData.length > 0) {
      sheet.appendRow(rowData);
      if (errors.length > 0) {
        sheet.getRange(rowPos, 1, 1, rowData.length).setBackground('yellow');
        errors.forEach(error => {
          const colPos = getColumnIndex(error.col);
          sheet.getRange(rowPos, colPos).setBackground('red');
        });
      }
      // Color any rows where Correct Answers that do not match Choices A-E
      if (rowData[4] !== rowData[5] && 
      rowData[4] !== rowData[6] && 
      rowData[4] !== rowData[7] && 
      rowData[4] !== rowData[8] && 
      rowData[4] !== rowData[9]) {
        sheet.getRange(rowPos, 1, 1, rowData.length).setBackground('yellow');
        sheet.getRange(rowPos, 5).setBackground('red');
      }
      // Color any rows where Learning Objective had multiples pass through (tested with bullet points)
      const bulletRegex = /[\u2022\u2023\u25E6\u2043\u2219\•\▪\‣\※]/;
      if (bulletRegex.test(rowData[11])) {
        sheet.getRange(rowPos, 1, 1, rowData.length).setBackground('yellow');
        sheet.getRange(rowPos, 12).setBackground('red');
      }
    }
  }
}

function getColumnIndex(columnLetter) {
  return columnLetter.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
}

// Function to extract data from a cell, handling nested tables if present
function extractCellText(cell) {
  const text = cell.getText();
  return text;
}


// Function to extract data from a pair of tables
function extractTablePairData(firstTable, secondTable) {
  const tableData = [];
  const errors = [];

  // New lines to check for question status once that change is approved - add brackets over try blocks and an else to return empty tableData and errors
  const questionStatus = findMarkedValue(extractCellText(secondTable.getRow(14).getCell(1)))
  if (questionStatus == 'Finalized') {

    try { //Actual: A
      tableData.push(extractCellText(firstTable.getRow(0).getCell(1)));  // Case to Col A
    } catch (error) {
      tableData.push(""); 
      errors.push({ row: 0, col: 'A', error });
    }

    try {
      tableData.push(extractCellText(secondTable.getRow(0).getCell(1))); // Vignette to Col B
    } catch (error) {
      tableData.push(""); 
      errors.push({ row: 0, col: 'B', error });
    }

    try {
      tableData.push(extractCellText(secondTable.getRow(1).getCell(1))); // Stem to Column C
    } catch (error) {
      tableData.push(""); 
      errors.push({ row: 0, col: 'C', error });
    }

    try {
      tableData.push("");  // Column D left empty
    } catch (error) {
      tableData.push(""); 
      errors.push({ row: 0, col: 'D', error });
    }

    try {
      tableData.push(extractCellText(secondTable.getRow(2).getCell(1)));  // Correct Answer to Column E
    } catch (error) {
      tableData.push(""); 
      errors.push({ row: 0, col: 'E', error });
    }

    try {
      tableData.push(extractCellText(secondTable.getRow(3).getCell(1)));  // Choice A to Column F
    } catch (error) {
      tableData.push(""); 
      errors.push({ row: 0, col: 'F', error });
    }

    try {
      tableData.push(extractCellText(secondTable.getRow(4).getCell(1)));  // Choice B to Column G
    } catch (error) {
      tableData.push(""); 
      errors.push({ row: 0, col: 'G', error });
    }

    try {
      tableData.push(extractCellText(secondTable.getRow(5).getCell(1)));  // Choice C to Column H
    } catch (error) {
      tableData.push(""); 
      errors.push({ row: 0, col: 'H', error });
    }

    try {
      tableData.push(extractCellText(secondTable.getRow(6).getCell(1)));  // Choice D to Column I
    } catch (error) {
      tableData.push(""); 
      errors.push({ row: 0, col: 'I', error });
    }

    try {
      tableData.push(extractCellText(secondTable.getRow(7).getCell(1)));  // Choice E to Column J
    } catch (error) {
      tableData.push(""); 
      errors.push({ row: 0, col: 'J', error });
    }

    try {
      tableData.push('Validation now');  // Validation now to Column K
    } catch (error) {
      tableData.push(""); // Empty value
      errors.push({ row: 0, col: 'K', error });
    }

    try {
      tableData.push(stripBulletPoints(extractCellText(firstTable.getRow(5).getCell(1))));  // Learning Objective to Column L
    } catch (error) {
      tableData.push(""); 
      errors.push({ row: 0, col: 'L', error });
    }

    try {
      tableData.push(""); // Leave Column M blank
    } catch (error) {
      tableData.push(""); // Empty value
      errors.push({ row: 0, col: 'M', error });
    }

    try {
      tableData.push(extractFirstLine(extractCellText(firstTable.getRow(4).getCell(1))));  // Teaching Point to Column N
    } catch (error) {
      tableData.push(""); 
      errors.push({ row: 0, col: 'N', error });
    }

    try {
      tableData.push(""); // Leave Column O blank
    } catch (error) {
      tableData.push(""); // Empty value
      errors.push({ row: 0, col: 'O', error });
    }

    try { // Use findMarkedValue to extract only the system marked by the user with an 'x'
      tableData.push(findMarkedValue(extractCellText(secondTable.getRow(10).getCell(1))));  // System to Column P
    } catch (error) {
      tableData.push(""); 
      errors.push({ row: 0, col: 'P', error });
    }

    try {
      tableData.push('PracticeSmart'); // Question Use in Column Q
    } catch (error) {
      tableData.push(""); // Empty value
      errors.push({ row: 0, col: 'Q', error });
    }

    try {
      tableData.push(extractCellText(secondTable.getRow(13).getCell(1))); // Answer Explanation in Column R
    } catch (error) {
      tableData.push(""); 
      errors.push({ row: 0, col: 'O', error });
    }

    try { // Regex for x markings
      tableData.push(findMarkedValue(extractCellText(secondTable.getRow(11).getCell(1)))); // Clinical Focus in Column S
    } catch (error) {
      tableData.push(""); // Empty value
      errors.push({ row: 0, col: 'S', error });
    }

    try {
      const clinicalDiscipline = (extractCellText(firstTable.getRow(0).getCell(1)).match(/^[^\d]*/) || [""])[0].trim(); // Clinical Discipline in T
      tableData.push(clinicalDiscipline);
    } catch (error) {
      tableData.push(""); 
      errors.push({ row: 0, col: 'T', error });
    }

    // Fill columns (U, V, W) with empty strings
    tableData.push("");
    tableData.push("");
    tableData.push("");

    try { // Regex for x markings
      tableData.push(findMarkedValue(extractCellText(secondTable.getRow(12).getCell(1)))); // Clinical Location in Column X
    } catch (error) {
      tableData.push(""); // Empty value
      errors.push({ row: 0, col: 'S', error });
    }

    Logger.log("Compiled row data: " + JSON.stringify(tableData));
    return { rowData: tableData, errors };
  } else {
    // Return empty arrays so the script doesn't log the row
    return { rowData: tableData, errors};
  }
}

function openSheetInBrowser(sheet) {
  const url = sheet.getParent().getUrl();
  const htmlOutput = HtmlService.createHtmlOutput('<script>window.open("' + url + '", "_blank");google.script.host.close();</script>');
  DocumentApp.getUi().showModalDialog(htmlOutput, 'Opening Spreadsheet');
  Logger.log(`New Sheet URL for your records: ${url}`);
}

// Open a 'working' message indicating the user should hang on as each process completes - automatically closes to pop new one, auto closes when program finishes
function makeModal(message) {
  var modalHTML = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          .modal-container {
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            width: 100%;
            height: 100%;
            text-align: center;
          }
          
          .spinner {
            width: 100px;
            height: 100px;
            border: 5px solid rgba(0,0,0,0.1);
            border-top: 5px solid #3498db;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            margin-bottom: 10px;
          }

          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }

          .message {
            font-size: 20px;
            color: #333;
          }

          body {
            margin: 0;
            display: flex;
            align-items: center;
            justify-content: center;
            height: 200px;
            width: 200px;
          }
        </style>
      </head>
      <body>
        <div class="modal-container">
          <div class="spinner"></div>
          <div class="message">Please wait...</div>
        </div>
        <script>
          setTimeout(() => {
            google.script.host.close();
          }, 30000);
        </script>
      </body>
    </html>
  `);
  DocumentApp.getUi().showModalDialog(modalHTML, message);
}

function setRowHeight(sheet) {
  const rowLength = sheet.getLastRow()
  sheet.setRowHeightsForced(1, rowLength, 21)
}

function extractFirstLine(inputString) {
  // Log the input string for debugging purposes
  Logger.log(`String passed to extractFirstLine: ${inputString}`);

  // Regex to match everything up to the first newline
  const regex = /^[^\r\n]*/;
  // Extract the first line
  const output = inputString.match(regex)[0].trim();
  
  // Log the output string for debugging purposes
  Logger.log(`extractFirstLine output: ${output}`);
  
  return output;
}

function stripBulletPoints(inputString) {
  // Regex to match the bullet point and any subsequent space(s)
  // Added logs for debugging
  Logger.log(`String passed to stripBulletPoints: ${inputString}`);
  const regex = /^[\u2022\s]*/;
  // Replace the bullet point at the start of the string with an empty string
  const output = inputString.replace(regex, '');
  Logger.log(`stripBulletPoints output: ${output}`);
  
  return output;
}

function findMarkedValue(string) {
  // Check that the string passed matches expected format
  Logger.log(`String passed to findMarkedValue for regex matches: ${string}`);
  
  // Define the regex pattern to match different placements of 'x' within underscores followed by a space and text
  let regexPattern = /(?:_?[xX]_?_?_?_\s)(.+)$/gim;

  // Use the matchAll method to find all matches
  let matches = Array.from(string.matchAll(regexPattern));
  Logger.log('Printing regex matches: ' + JSON.stringify(matches));
  
  // Extract the marked values and remove extra spaces
  let markedValues = matches.map(match => match[1].trim());
  Logger.log('Checking markedValues before joined & returned: ' + JSON.stringify(markedValues));

  // Join all marked values with commas and return the result
  return markedValues.join(', ');
}

function logTableContents(table) {
  const numRows = table.getNumRows();
  
  for (let i = 0; i < numRows; i++) {
    const row = table.getRow(i);
    const numCells = row.getNumCells();
    
    for (let j = 0; j < numCells; j++) {
      const cell = row.getCell(j);
      const cellText = cell.getText();
      Logger.log('Row %s, Cell %s: %s', i, j, cellText);
    }
  }
}
