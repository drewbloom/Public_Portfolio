function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Reporting')
      .addItem('Generate Comparison', 'generateComparison')
      .addItem('Get AAMC Embeddings', 'getAAMCEmbeddings')
      .addToUi();
  }
  
  // Drew's Key, replace if needed
  const OPENAI_API_KEY = 'MY_OPENAI_API_KEY';  
  
  // Function to get embeddings from OpenAI API
  function getEmbedding(text) {
    var url = 'https://api.openai.com/v1/embeddings';
    var payload = {
      model: 'text-embedding-3-small',
      input: text
    };
    
    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + OPENAI_API_KEY
      },
      payload: JSON.stringify(payload)
    };
    
    try {
      var response = UrlFetchApp.fetch(url, options);
      var json = JSON.parse(response.getContentText());
      return json.data[0].embedding;
    } catch (e) {
      Logger.log('Error: ' + e.toString());
      return null;
    }
  }
  
  function getAAMCEmbeddings() {
    makeModal('Starting AAMC Embeddings...');
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Embeddings');
    var lastRow = sheet.getLastRow();
    var cData = sheet.getRange('C2:C' + lastRow).getValues(); // Only get data for populated rows, ignore empty cells
    
    // Map through the data from Column C
    var concatenatedText = cData.map(function(row) {
      return row[0];
    });
    makeModal('Got Sheets data, calling OpenAI for embeddings...');
  
    var embeddings = concatenatedText.map(function(text) {
      var embedding = getEmbedding(text);
      if (embedding === null) {
        // Handle the case where an embedding could not be obtained
        return Array(1536).fill(0); // Placeholder for missing embeddings
      } else {
        return embedding;
      }
    });
  
    // Convert embeddings to strings since Google Sheets cells cannot hold arrays directly
    makeModal('Converting embeddings to write to Sheets...');
    var embeddingsStrings = embeddings.map(function(embedding) {
      return JSON.stringify(embedding);
    });
  
    // Write embeddings to Column D
    makeModal('Writing results to Sheets...');
    var range = sheet.getRange(2, 4, embeddingsStrings.length, 1);
    range.setValues(embeddingsStrings.map(function(e) { return [e]; }));
    closeTheModal();
  }
  
  // Web App function to call by passing the sheetname parameter
  function webGenComparison(newSheet) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(newSheet);
    var lastRow = sheet.getLastRow();  // Get the last populated row
  
    // Clear previous run data if it exists - modify to expand columns to clear as we add more functionality / visuals
    sheet.getRange('E2:P' + lastRow).clearContent();
  
    var aData = sheet.getRange('A2:A' + lastRow).getValues();
  
    // Filter out empty rows
    aData = aData.filter(function(row) {
      return row[0].trim() !== "";
    });
  
    // Extract text entries from Column A
    var programObjectives = aData.map(function(row) {
      return row[0];
    });
  
    // Generate embeddings for Column A entries
    var embeddingsA = programObjectives.map(function(text) {
      return getEmbedding(text);
    });
  
    // Write embeddings for Column A to Column E for checking
    var embeddingsAStrings = embeddingsA.map(function(embedding) {
      return JSON.stringify(embedding);
    });
    sheet.getRange(2, 5, embeddingsAStrings.length, 1).setValues(embeddingsAStrings.map(function(e) { return [e]; }));  // 5 is Column E
  
    // Retrieve and parse static embeddings from Column D
    var dData = sheet.getRange('D2:D' + lastRow).getValues().flat();
    var embeddingsD = dData.map(function(embedString) {
      try {
        return JSON.parse(embedString);
      } catch (e) {
        return Array(1536).fill(0);  // Handle potential parsing issues
      }
    });
  
    // Retrieve static data from Col C
    var cData = sheet.getRange('C2:C' + lastRow).getValues();
  
    // Calculate cosine similarities and store matches
    var matches = [];
    for (var i = 0; i < embeddingsA.length; i++) {
      var rowMatches = [];
      for (var j = 0; j < embeddingsD.length; j++) {
        var similarity = cosineSimilarity(embeddingsA[i], embeddingsD[j]);
        if (similarity >= 0.50) {  // Threshold for matching
          rowMatches.push(`${cData[j]} (${similarity.toFixed(5)})`);  // Append text and include similarity score
        }
      }
      matches.push(rowMatches);
    }
  
    // Convert matches to the correct format for Google Sheets
    // Ensure MatchLength is at least 1 so no error if 0 matches
    var maxMatchLength = Math.max(1, ...matches.map(arr => arr.length));
    var matchStrings = matches.map(function(matchRow) {
      while (matchRow.length < maxMatchLength) {
        matchRow.push("");  // Pad with empty strings for uniform columns count
      }
      return matchRow;
    });
    
    // Write matches to columns starting from Column F (i.e., "Match 1" in Column F, "Match 2" in Column G, etc.)
    var matchRange = sheet.getRange(2, 6, matchStrings.length, maxMatchLength);  // 6 is Column F
    matchRange.setValues(matchStrings);
  
    // Check for matches and populate Column P
    var matchStatus = matches.map(function(matchRow) {
      return [matchRow[0].trim() !== "" ? "Y" : "N"];
    });
    sheet.getRange(2, 16, matchStatus.length, 1).setValues(matchStatus);  // 16 is Column P
  
  }
  
  // Regular gen comparison function for use within google sheets
  function generateComparison() {
    makeModal('Starting Comparison Generation...');
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Embeddings');
    var lastRow = sheet.getLastRow();  // Get the last populated row
  
    // Clear previous run data if it exists - modify to expand columns to clear as we add more functionality / visuals
    sheet.getRange('E2:P' + lastRow).clearContent();
  
    var aData = sheet.getRange('A2:A' + lastRow).getValues();
  
    // Filter out empty rows
    aData = aData.filter(function(row) {
      return row[0].trim() !== "";
    });
  
    // Extract text entries from Column A
    var programObjectives = aData.map(function(row) {
      return row[0];
    });
  
    makeModal('Getting embeddings for MEPO entries...');
  
    // Generate embeddings for Column A entries
    var embeddingsA = programObjectives.map(function(text) {
      return getEmbedding(text);
    });
  
    // Write embeddings for Column A to Column E for checking
    var embeddingsAStrings = embeddingsA.map(function(embedding) {
      return JSON.stringify(embedding);
    });
    sheet.getRange(2, 5, embeddingsAStrings.length, 1).setValues(embeddingsAStrings.map(function(e) { return [e]; }));  // 5 is Column E
  
    makeModal('Getting AAMC embeddings from Sheets...');
  
    // Retrieve and parse static embeddings from Column D
    var dData = sheet.getRange('D2:D' + lastRow).getValues().flat();
    var embeddingsD = dData.map(function(embedString) {
      try {
        return JSON.parse(embedString);
      } catch (e) {
        return Array(1536).fill(0);  // Handle potential parsing issues
      }
    });
  
    // Retrieve static data from Col C
    var cData = sheet.getRange('C2:C' + lastRow).getValues();
  
    makeModal('Calculating similarity scores...');
    // Calculate cosine similarities and store matches
    var matches = [];
    for (var i = 0; i < embeddingsA.length; i++) {
      var rowMatches = [];
      for (var j = 0; j < embeddingsD.length; j++) {
        var similarity = cosineSimilarity(embeddingsA[i], embeddingsD[j]);
        if (similarity >= 0.50) {  // Threshold for matching
          rowMatches.push(`${cData[j]} (${similarity.toFixed(5)})`);  // Append text and include similarity score
        }
      }
      matches.push(rowMatches);
    }
  
    // Convert matches to the correct format for Google Sheets
    // Ensure MatchLength is at least 1 so no error if 0 matches
    var maxMatchLength = Math.max(1, ...matches.map(arr => arr.length));
    var matchStrings = matches.map(function(matchRow) {
      while (matchRow.length < maxMatchLength) {
        matchRow.push("");  // Pad with empty strings for uniform columns count
      }
      return matchRow;
    });
    
    makeModal('Writing matches to Sheets...');
    // Write matches to columns starting from Column F (i.e., "Match 1" in Column F, "Match 2" in Column G, etc.)
    var matchRange = sheet.getRange(2, 6, matchStrings.length, maxMatchLength);  // 6 is Column F
    matchRange.setValues(matchStrings);
  
    // Check for matches and populate Column P
    var matchStatus = matches.map(function(matchRow) {
      return [matchRow[0].trim() !== "" ? "Y" : "N"];
    });
    sheet.getRange(2, 16, matchStatus.length, 1).setValues(matchStatus);  // 16 is Column P
  
    // Close down Modal to finish the execution in the UI
    closeTheModal();
  }
  
  // Function to compute cosine similarity between two vectors
  function cosineSimilarity(vecA, vecB) {
    var dotProduct = 0;
    var normA = 0;
    var normB = 0;
    
    for (var i = 0; i < vecA.length; i++) {
      dotProduct += vecA[i] * vecB[i];
      normA += vecA[i] * vecA[i];
      normB += vecB[i] * vecB[i];
    }
    
    normA = Math.sqrt(normA);
    normB = Math.sqrt(normB);
    return dotProduct / (normA * normB);
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
            function closeModal() {
              google.script.host.close();
            }
          </script>
        </body>
      </html>
    `);
    SpreadsheetApp.getUi().showModalDialog(modalHTML, message);
  }
  
  function closeTheModal() {
    var htmlOutput = HtmlService.createHtmlOutput(`
      <script>
        google.script.host.close();
      </script>
    `);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Finished!');
  }
  
  /* Additional Functions for creating visuals in Sheets
  */
  
  // Return a count of match quality for a pie chart based on custom bounds
  function countMatches(range, lowerBound, upperBound) {
    let count = 0;
    for (let i = 0; i < range.length; i++) {
      for (let j = 0; j < range[i].length; j++) {
        let text = range[i][j];
        let startIdx = text.indexOf('(') + 1;
        let endIdx = text.indexOf(')');
        if (startIdx > 0 && endIdx > startIdx) {
          let score = parseFloat(text.substring(startIdx, endIdx));
          if (!isNaN(score) && score > lowerBound && score <= upperBound) {
            count++;
          }
        }
      }
    }
    return count;
  }
  
  // Return a list of the Program Objectives with their Excellent Matches
  function listExcellentMatches(range, programObjectivesColumn, lowerBound, upperBound) {
    const matches = [];
    
    // Get the values of the Program Objectives
    const programObjectives = programObjectivesColumn.map(function(row) { return row[0]; });
  
    for (let i = 0; i < range.length; i++) {
      for (let j = 0; j < range[i].length; j++) {
        const text = range[i][j];
        const startIdx = text.indexOf('(') + 1;
        const endIdx = text.indexOf(')');
        if (startIdx > 0 && endIdx > startIdx) {
          const score = parseFloat(text.substring(startIdx, endIdx));
          if (!isNaN(score) && score > lowerBound && score <= upperBound) {
            matches.push(programObjectives[i] + ": " + text);
          }
        }
      }
    }
    
    return matches;
  }
  
  // Return a list of Top 5 Program Objectives that have the most matches with # of matches
  function top5MostAligned(range, programObjectivesColumn) {
    const counts = [];
    const output = [];
    
    // Get the values of the Program Objectives
    const programObjectives = programObjectivesColumn.map(row => row[0]);
  
    for (let i = 0; i < range.length; i++) {
      let count = 0; // Make sure to declare count with let within the loop
      for (let j = 0; j < range[i].length; j++) {
        if (range[i][j] && String(range[i][j]).trim() !== '') {
          count++;
        }
      }
      counts.push({ objective: programObjectives[i], count: count });
    }
    
    // Sort the counts in descending order
    counts.sort((a, b) => b.count - a.count);
    
    // Get the top 5 counts
    for (let k = 0; k < 5 && k < counts.length; k++) {
      output.push(`${counts[k].count} Matches: ${counts[k].objective}`);
    }
  
    return output;
  }