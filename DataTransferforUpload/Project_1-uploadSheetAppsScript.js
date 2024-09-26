function onOpen() {
  DocumentApp.getUi().createMenu('Upload Sheet')
      .addItem('Copyedit Document', 'copyEdit')
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

  try {
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
        tableData.push(stripBulletPoints(findMarkedValue(extractCellText(firstTable.getRow(5).getCell(1)))));  // Learning Objective to Column L
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
  } catch (error) {
    // Return an empty array if the table doesn't have a 15th row with question status
    tableData.push("");
    errors.push({row: 0, col: 'A', error});
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

// Beginning of pre-upload functionality to copyedit document
/** Goals and functionality
 * Work element-by-element to maintain most of the document's formatting
 * Perform all necessary programmatic functions in a loop through the elements
 * Save every 50 elements by closing and reopening doc
 * Log every change made by each of the regex / editing functions within the loop using an index
 */
function copyEdit() {
  let doc = DocumentApp.getActiveDocument();
  const docID = doc.getId();
  
  // Get all elements in the document body
  let body = doc.getBody();
  const totalElements = body.getNumChildren();
  const errorElements = [];
  let element;

  // Loop through each element in the body
  for (let i = 0; i < totalElements; i++) {
    element = body.getChild(i);
    let elementTracker = "";
    
    // Check type of element and process accordingly
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      elementTracker = "Paragraph";
      try {
        processParagraph(element, i);
      } catch (error) {
        Logger.log(`Failed to process element ${i} (${elementTracker}): ${error}`);
        const elementText = element.getText();
        if (elementText) {
          const textSnip = elementText.substring(0,100);
          errorElements.push(`errorElement ${errorElements.length}: Element ${i} (${elementTracker}):\n${textSnip}`);
        }
        continue
      }
    } else if (element.getType() === DocumentApp.ElementType.TABLE) {
      elementTracker = "Table";
      try {
        processTables(element, i);
      } catch (error) {
        Logger.log(`Failed to process element ${i} (${elementTracker}): ${error}`);
        const elementText = element.getText();
        const textSnip = elementText.substring(0,100);
        errorElements.push(`errorElement ${errorElements.length}: Element ${i} (${elementTracker}):\n${textSnip}`);
        continue
      }
      try {
        handleTableSpecificTasks(element);
      } catch (error) {
        const elementText = element.getText();
        const textSnip = elementText.substring(0,100);
        Logger.log(`Failed handleTableSpecificTasks for Table ${i} - Text Snip:\n${textSnip}\nError: ${error}`);
      }
    }
    
    // Perform save and continue conditionally, if processing too many elements
    if (i > 0 && i % 50 === 0) {
      doc.saveAndClose();
      Logger.log(`Saved and closed document at element ${i}`);
      try {
        doc = DocumentApp.openById(docID);
        body = doc.getBody();
      } catch (error) {
        Logger.log(`Error opening Doc ID: ${docID}\n -- ${error}`);
        throw new Error(`Stopped Script: Error reopening Doc`);
      }
    }
    

    
    Logger.log(`Completed element ${i} (${elementTracker})`);
  }

  // Display all elements with errors for examination
  if (errorElements.length > 0) {
    Logger.log(`Printing all elements with errors:\n${errorElements.join('\n')}`);
  }

  setSameFont(docID);
  Logger.log('Copyediting complete.');
}

function processParagraph(paragraph, elementIndex) {
  let text = paragraph.getText();
  
  text = removeDoublePeriods(text, elementIndex);
  text = removeDoubleSpace(text, elementIndex);
  text = capitalizeFirstWord(text, elementIndex);
  text = deniesToDNR(text, elementIndex);
  text = removeTrailingWhitespace(text, elementIndex);
  text = removeXXX(text, elementIndex);
  text = changeEKGtoECG(text, elementIndex);
  text = changeERtoED(text, elementIndex);
  text = correctXray(text, elementIndex);
  // Abbreviation highlights are making the test non-text for setText below
  // text = addAbbreviationHighlights(text, elementIndex);
  
  paragraph.setText(text);
}

function processTables(table, elementIndex) {
  const numRows = table.getNumRows();
  const numCols = table.getRow(0).getNumCells();

  for (let r = 0; r < numRows; r++) {
    for (let c = 0; c < numCols; c++) {
      const cell = table.getCell(r, c);
      let text = cell.getText();

      text = removeDoublePeriods(text, elementIndex);
      text = removeDoubleSpace(text, elementIndex);
      text = capitalizeFirstWord(text, elementIndex);
      text = deniesToDNR(text, elementIndex);
      text = removeTrailingWhitespace(text, elementIndex);
      text = removeXXX(text, elementIndex);
      text = changeEKGtoECG(text, elementIndex);
      text = changeERtoED(text, elementIndex);
      text = correctXray(text, elementIndex);
      // Removing highlights for testing
      // text = addAbbreviationHighlights(text, elementIndex);

      cell.setText(text);
    }
  }
}

// Function to remove highlighting from each paragraph
function removeHighlighting(docID) {
  const doc = DocumentApp.openById(docID);
  const body = doc.getBody();
  const totalElements = body.getNumChildren();

  for (let i = 0; i < totalElements; i++) {
    const element = body.getChild(i);

    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
      processRemoveHighlightingFromParagraph(element, i);
    }
  }

  Logger.log('Removed all text highlighting.');
}

function processRemoveHighlightingFromParagraph(paragraph, elementIndex) {
  const text = paragraph.getText();
  const textElement = paragraph.editAsText();
  const length = text.length;

  for (let j = 0; j < length; j++) {
    textElement.setBackgroundColor(j, j, null);  // Setting background color to null to remove highlighting
  }
  
  Logger.log(`Removed highlighting in paragraph at index ${elementIndex}.`);
}

// Function to set font for all elements to Arial
function setSameFont(docID) {
  const doc = DocumentApp.openById(docID);
  const body = doc.getBody();
  const text = body.getText();
  body.editAsText().setFontFamily('Arial'); // Sets font to Arial
  Logger.log('Set font to Arial.');
  body.editAsText().setFontSize(10);
  Logger.log('Set font size to 10');
}

// Helper function to log specific changes
function logChange(operation, position, textBefore, textAfter) {
  Logger.log(`${operation} at position ${position}: "${textBefore}" -> "${textAfter}"`);
}

// Helper functions with regex tasks
function removeDoublePeriods(text, offset = 0) {
  const regex = /\.{2,}/g;
  let match;
  while ((match = regex.exec(text)) !== null) {
    logChange('Removed double periods', offset + match.index, match[0], ".");
  }
  return text.replace(regex, ".");
}

function removeDoubleSpace(text, offset = 0) {
  const regex = / {2,}/g;
  let match;
  while ((match = regex.exec(text)) !== null) {
    logChange('Removed double spaces', offset + match.index, match[0], " ");
  }
  return text.replace(regex, " ");
}

// Preserve all whitespace characters but return a capitalized first letter of the next word
function capitalizeFirstWord(text, offset = 0) {
  // Regex to match '.', '!', or '?' followed by one or more whitespace characters and a word character
  const regex = /(\.|\!|\?)\s+(\w)/g;
  let match;
  
  // Use the replace function to capitalize the matched character and log the change
  let newText = text.replace(regex, function(match, p1, p2, index) {
    const capitalizedChar = p2.toUpperCase(); 
    const replacement = p1 + match.slice(p1.length, -1) + capitalizedChar; // Reconstruct the match with the capitalized character
    logChange('Capitalized first word after punctuation', offset + index, match, replacement);
    return replacement;
  });
  
  return newText;
}

function deniesToDNR(text, offset = 0) {
  const regex = /\bdenies\b/gi;
  let match;
  while ((match = regex.exec(text)) !== null) {
    logChange('Replaced "denies" with "does not report"', offset + match.index, match[0], "does not report");
  }
  return text.replace(regex, "does not report");
}

function removeTrailingWhitespace(text, offset = 0) {
  const regex = /\s+$/gm;
  // Use matchAll to capture all matches
  const matches = [...text.matchAll(regex)];
  
  // Log each change
  matches.forEach(match => {
    logChange('Removed trailing whitespace', offset + match.index, match[0], "");
  });

  // Remove all trailing whitespaces
  return text.replace(regex, "");
}

function removeXXX(text, offset = 0) {
  const regex = /\[xxx\]/gi;
  let match;
  while ((match = regex.exec(text)) !== null) {
    logChange('Removed [xxx]', offset + match.index, match[0], "");
  }
  return text.replace(regex, "");
}

function changeEKGtoECG(text, offset = 0) {
  const regex = /\bEKG\b/gi;
  let match;
  while ((match = regex.exec(text)) !== null) {
    logChange('Replaced "EKG" with "ECG"', offset + match.index, match[0], "ECG");
  }
  return text.replace(regex, "ECG");
}

function changeERtoED(text, offset = 0) {
  const regexList = [
    /\bemergency room\b/gi,
    /\bE\.R\.\b/g,
    /\bER\b/g,
    /\bE\.D\.\b/g,
    /\bEmergency Department\b/g
  ];
  let newText = text;
  regexList.forEach(regex => {
    let match;
    while ((match = regex.exec(newText)) !== null) {
      logChange('Replaced "ER" variants with "emergency department"', offset + match.index, match[0], "emergency department");
    }
    newText = newText.replace(regex, "emergency department");
  });
  return newText;
}

function correctXray(text, offset = 0) {
  const regexList = [
    /\bX-ray\b/g,
    /\bX-Ray\b/g,
    /\bxray\b/g
  ];
  let newText = text;
  regexList.forEach(regex => {
    let match;
    while ((match = regex.exec(newText)) !== null) {
      logChange('Corrected X-ray notation', offset + match.index, match[0], "x-ray");
    }
    newText = newText.replace(regex, "x-ray");
  });
  return newText;
}

function addAbbreviationHighlights(text, offset = 0) {
  const highlightWords = ['MRI', 'ECG', 'ED', 'CT'];

  highlightWords.forEach(word => {
    let start = 0;
    while((start = text.indexOf(word, start)) !== -1) {
      body.editAsText().setBackgroundColor(start, start + word.length - 1, '#FFFF00');  // Highlight in yellow
      start += word.length;
      logChange('Highlighted abbreviation', offset + start, word, word);
    }
  });

  Logger.log('Highlighted specific abbreviations.');
}

// Cell-specific task handler for tables
function handleTableSpecificTasks(table) {

  if (checkTableIdentifier(table)) {
    // check that it contains Vignette, otherwise skip, then proceed with functions
    Logger.log('Table has Vignette, checking cells for edits: Vignette, Correct Answer, Answer Explanation');
    const vignetteCell = table.getCell(0, 1);
    Logger.log(`VignetteCell Object returned: ${vignetteCell}`);
    let vignetteText = vignetteCell.getText();
    vignetteText = changeSexToPatient(vignetteText);
    vignetteCell.setText(vignetteText);

    // handle Correct Answer
    const correctAnswerCell = findTableCell(table, "Correct Answer:");
    if (correctAnswerCell) {
      Logger.log(`CorrectAnswerCell Object returned: ${correctAnswerCell}`);
      // const targetCell = table.getCell(correctAnswerCell.getRow(), correctAnswerCell.getCell() + 1);
      const targetCell = correctAnswerCell.getNextSibling();
      let answerText = targetCell.getText();
      // Pass the string through both regex to pull out "A. " and "Choice A: "
      answerText = removeChoiceLetters(answerText);
      answerText = removeCorrectAnswerLettering(answerText);
      targetCell.setText(answerText);
    }

    // handle Answer Explanation
    const explanationCell = findTableCell(table, "Answer Explanation:");
    if (explanationCell) {
      Logger.log(`ExplanationCell Object returned: ${explanationCell}`);
      // const targetCell = table.getCell(explanationCell.getRow(), explanationCell.getCell() + 1);
      const targetCell = explanationCell.getNextSibling();
      let explanationText = targetCell.getText();
      explanationText = removeChoiceLetters(explanationText);
      targetCell.setText(explanationText);
    }
  } else {
    Logger.log("Table does not have 'Vignette:', continuing to next");
  }
  
}

function checkTableIdentifier(table) {
  return table.getCell(0, 0).getText() === "Vignette:";
}

function findTableCell(table, targetText) {
  for (let r = 0; r < table.getNumRows(); r++) {
    for (let c = 0; c < table.getRow(r).getNumCells(); c++) {
      if (table.getCell(r, c).getText() === targetText) {
        return table.getCell(r,c);
      }
    }
  }
  return null;
}

function changeSexToPatient(text) {
  const regex = /\b(Female|female|Male|male)\b/g;
  const newText = text.replace(regex, function(match, index) {
    logChange('Replaced sex with patient in first sentence', index, match, "patient");
    return "patient";
  });
  return newText;
}

function removeCorrectAnswerLettering(text) {
  const regex = /^[A-E]\.\s/;
  const newText = text.replace(regex, function(match, index) {
    logChange('Removed correct answer lettering', index, match, "");
    return "";
  });
  return newText;
}

function removeChoiceLetters(text) {
  const regex = /\bChoice [A-E](: )?\s?/gi;
  const newText = text.replace(regex, function(match, index) {
    logChange('Removed choice letters', index, match, "");
    return "";
  });
  return newText;
}

// Space for additional functions:
// Additional revisions: N/A