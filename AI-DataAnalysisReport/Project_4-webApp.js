// Functions to run the web app UI

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index');
}

// OpenAI API Key already declared outside all functions in Embeddings.gs, can be called by these functions

// System message guidance
const initialSystemMessage = {
  role: 'system',
  content: 'You are an assistant that extracts learning objectives from raw text inputs. Use the few-shot user and assistant messages as examples to guide you. Follow their instructions and formatting rules, and indicate when you are finished with your task. Remember, only respond with listed learning objectives, NULL, TASK_COMPLETE, or CONCATNEXT as needed. Your instructions are contained in the few-shot messages, then you will start your work when you are asked to process a new <user_input>text.'
};

// The few-shot examples and queries
const fewShotExamples = [
  // Few-Shot Examples from real PDFs to show the model how to find these in context and return in desired formats
  { role: 'user', content: 'This user message contains instructions and a sample and is followed by an assistant message containing a correct output. Your main task is to extract learning objectives from raw text inputs.  You can locate them in context by reading section-by-section as we feed you chunks of text. You need to return a list of the learning objectives you located, stripped of any numbers or categorical indicators and separated with a delimiter "|".  These will often be marked with numbers and contain extra categorical indicators for you to work with.  If you dont find any objectives, respond "NULL". Here is a sample text where the assistant returns the correct output.\n<user_input>MEDICAL EDUCATION PROGRAM OBJECTIVES TEXAS A&M UNIVERSITY School of Medicine KNOWLEDGEStudents will successfully: K1. Demonstrate knowledge around the prevention, diagnosis, treatment, management, (cure or palliation) of medical conditions (MK, PC); K2. Describe the anatomy, histology, pathology, and pathophysiology of the human body as it pertains to the organ systems (MK, PC); K3. Identify the processes of patient history taking, physical examination, assessment, and plan development (MK, PC);' 
  },
  { role: 'assistant', content: 'Demonstrate knowledge around the prevention, diagnosis, treatment, management, (cure or palliation) of medical conditions | Describe the anatomy, histology, pathology, and pathophysiology of the human body as it pertains to the organ systems | Identify the processes of patient history taking, physical examination, assessment, and plan development |' 
  },
  { role: 'user', content: `Here is another sample in which the assistant returns a correct output. You have to be ready for strange formatting because much of the text we will parse originates from pdfs. If you cannot find a complete objective and expect it to be finished in the next chunk, mark it CONCATNEXT.  That also means you should be ready to determine from context whether you received a partial objective at the beginning of an input and return the partial as your first learning objective to make the concatenation work.\n<user_input>OBJECTIVES OF THE EDUCATIONAL PROGRAMMEDICAL AND POPULATION HEALTH KNOWLEDGEMedical school graduates must demonstrate knowledge about established and evolving biomedical,clinical, epidemiological, and social-behavioral sciences. The student’s transition from understandinginto action will require the application of this knowledge to patient care and population health. Medicalschool graduates are expected to:1. Demonstrate knowledgeof human development,structure, and function ofthe major organ systems,including their relevance tohealth.2. Demonstrate knowledge of thepathology, pathophysiology,and clinical manifestation ofdisease.3. Demonstrate knowledge ofthe effective use of laboratorytests, radiologic studies,` 
  },
  { role: 'assistant', content: 'Demonstrate knowledge of human development, structure, and function of the major organ systems, including their relevance to health. | Demonstrate knowledge of the pathology, pathophysiology,and clinical manifestation of disease. | Demonstrate knowledge of the effective use of laboratory tests, radiologic studies,CONCATNEXT' 
  },
  { role: 'user', content: `Here is another sample in which the assistant returns a correct output. Midway through your work, you may receive a chunk midstream that needs objectives located and indications for combining with other outputs.  Try to locate the objectives based on the formatting and remember to link partial objectives from earlier chunks by returning the first partial string segment.\n<user_input>efforts for patients and populations.
  Page 2
  Clinical Care
  This competency includes patient- and family-centered principles and components of clinical
  care.
  1 Establish mutually respectful student-patient-family relationships based on trust.
  2 Elicit a medical history appropriate to the patient's concerns and clinical context.
  3 Perform a physical exam appropriate to the patient’s presentation and clinical
  context.
  4 Evaluate the appropriateness of diagnostic tests and studies for a particular
  condition and clinical context.
  5 Identify and interpret the results of frequently ordered laboratory, imaging, and
  other diagnostic studies.
  6 Use clinical reasoning to synthesize relevant key patient findings into a concise
  and accurate assessment, including differential diagnosis.
  7
  Formulate a prioritized problem list, and develop and implement a management
  plan guided by the patient’s social context, evidence-based medicine, and critical
  thinking.
  8 Deliver oral presentations appropriate to the patient’s presentation and clinical
  context.
  9 Record clinical information that is accurate, organized, well-reasoned, and timely.
  10 Demonstrate proficiency in performing select clinical and operative procedures
  under appropriate supervision.
  11 Use information technology effectively and responsibly.
  12 Engage patients in shared decision-making, incorporating values and preferences
  in discussions of management options and their expected benefits and harms.
  13 Identify and address the various goals of patient care, including prevention,
  diagnosis, cure, chronic disease management, palliation, and end-of-life care.
  14 Work effectively in various health care settings and systems.
  15
  Contribute to the coordinated care of the patient, including referral of patients,
  ensuring continuity of care throughout transitions between providers or settings,
  and following up on patient progress and outcomes.
  Page 3
  16 Organize and prioritize responsibilities to provide care that is safe, effective, and
  efficient.
  17 Accept and demonstrate responsibility in a graded fashion commensurate with
  one’s roles, abilities, and qualifications.
  Page 4
  Population Health
  This competency includes an understanding of the multiple factors that contribute to
  population health, including the role of health care delivery systems.
  1 Differentiate between and critique measures used to evaluate health and disease
  at the individual and population level.
  2
  Assess the impact of social, environmental, behavioral, economic, cultural, and
  personal factors on the health of individuals, and the incidence a` 
  },
  { role: 'assistant', content: `efforts for patients and populations. | Establish mutually respectful student-patient-family relationships based on trust. | Elicit a medical history appropriate to the patient's concerns and clinical context. | (Assistant continues to list the objectives until the final objective located) | Assess the impact of social, environmental, behavioral, economic, cultural, and personal factors on the health of individuals, and the incidence aCONCATNEXT` 
  },
  { role: 'user', content: 'Here is a final example in which the assistant returns a correct output. You will notice from the context that these objectives are for faculty and the school instead, so be careful to check that the learning objectives are for students, in which case you would return that list as usual.  If no student learning objectives were found yet, you would return "NULL" and move on to process the next chunk. But if you already had an existing list of student learning objectives and saw a section that no longer contained any learning objectives, you should indicate that all learning objectives are found by returning "TASK_COMPLETE", which will cause us to stop sending chunks of text to you. Sample: STATEMENT OF PURPOSEThe F. Edward Hébert School of Medicine is dedicated to preparing ethical, highly competentphysicians who are committed to “caring for those in harm’s way,” and all other Military Health Systembeneficiaries. Towards that end, our faculty adopted the following objectives:1. We select applicants who demonstrate great potential as physicians and leaders and whodemonstrate a strong commitment to the advancement of military medicine and public health,by fulfilling a promise of duty and expertise.2. Our educational program is organized to maximize our students’ preparedness to meet theGeneral Competencies and Milestones of the American Board of Medical Specialties andAccreditation Council for Graduate Medical Education prior to, during, or by the completionof their Graduate Medical Education, as appropriate. It is designed to prepare our studentsto pass each of the components of the United States Medical Licensing Examination that arerequired for medical licensure. Further, it is organized to assure integration of the principlesof officership and leadership, and the knowledge, skills, abilities, and overall professionalismexpected of uniformed medical officers. Toward these goals, our faculty - with input fromstudent representatives – have developed the ensuing Statement of Objectives for theEducational Program of the F. Edward Hébert School of Medicine.3. We provide a student-centered educational program, which includes:• Opportunities for the exchange of both formal and informal feedback throughout the year,along with an opportunity for students to provide more detailed feedback at the end of eachmodule/clerkship and prior to graduation.' 
  },
  { role: 'assistant', content: 'TASK_COMPLETE' 
  },
];
const newQuery = { role: 'user', content: 'Begin your main task, extracting learning objectives from text contained in the system message.  Use the other user and assistant messages as examples to guide you. Follow their instructions and formatting rules and indicate when you are finished with your task. Remember, only respond with listed learning objectives, NULL, TASK_COMPLETE, or CONCATNEXT as needed' };
const STOP_SIGNAL = "TASK_COMPLETE";

// Extract text from a Google Doc
function extractTextFromGoogleDoc(docUrl) {
  const doc = DocumentApp.openByUrl(docUrl);
  return doc.getBody().getText();
}

// Convert PDF to Google Doc and extract text
/*
function extractTextFromPDF(pdfUrl) {
  const pdfId = getIdFromUrl(pdfUrl);
  const file = DriveApp.getFileById(pdfId);
  const blob = file.getBlob();

  const resource = {
    title: file.getName().replace('.pdf', ''),
    mimeType: MimeType.GOOGLE_DOCS
  };

  // Use Advanced Drive Service to insert the file
  const newFile = Drive.Files.create({
    title: resource.title,
    mimeType: resource.mimeType
  }, blob, { convert: true });
  
  const newDocId = newFile.id;
  const text = extractTextFromGoogleDoc(`https://docs.google.com/document/d/${newDocId}/edit`);

  DriveApp.getFileById(newDocId).setTrashed(true);
  return text;
}
*/

function extractTextFromPDF(pdfUrl) {
  const pdfId = getIdFromUrl(pdfUrl);
  const file = DriveApp.getFileById(pdfId);
  const blob = file.getBlob();
  
  const metadata = {
    name: file.getName().replace('.pdf', ''),
    mimeType: MimeType.GOOGLE_DOCS
  };
  
  const boundary = "-------314159265358979323846";
  const delimiter = "\r\n--" + boundary + "\r\n";
  const close_delim = "\r\n--" + boundary + "--";
  
  const requestBody = 
    delimiter +
    'Content-Type: application/json\r\n\r\n' +
    JSON.stringify(metadata) +
    delimiter +
    'Content-Type: ' + blob.getContentType() + '\r\n' +
    'Content-Transfer-Encoding: base64\r\n' +
    '\r\n' +
    Utilities.base64Encode(blob.getBytes()) +
    close_delim;
  
  const options = {
    method: 'post',
    contentType: 'multipart/related; boundary="' + boundary + '"',
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    },
    payload: requestBody,
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', options);
  const newFile = JSON.parse(response.getContentText());
  
  if (response.getResponseCode() !== 200) {
    throw new Error("Failed to create Google Doc from PDF: " + response.getContentText());
  }
  
  const newDocId = newFile.id;
  const text = extractTextFromGoogleDoc(`https://docs.google.com/document/d/${newDocId}/edit`);
  
  DriveApp.getFileById(newDocId).setTrashed(true);
  return text;
}

// Process form input
function processForm(formData) {
  const fileUrl = formData.fileUrl;
  const institution = formData.institution;

  // Check whether sheet already exists and end run if it does to prevent duplicate runs / wasted tokens
  const sheetCheck = `${institution} Embeddings`;
  if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetCheck)) {
    return `Error: Sheet with name "${sheetCheck}" already exists.`;
  }

  Logger.log(`Raw value of isPdf: ${formData.isPdf}`)
  // read isPdf as a boolean, not a string - this was source of validation issues
  const isPdf = formData.isPdf === true;

  Logger.log(`User generated report for institution: ${institution}`)
  Logger.log(`File submitted - URL: ${fileUrl}`)
  Logger.log(`File is PDF: ${isPdf}`);

  if (!fileUrl || !institution) {
    return 'Error: Both the URL and institution name are required.';
  }

  let extractedText;
  try {
    if (isPdf) {
      Logger.log(`Extracting text from PDF: ${fileUrl}`);
      extractedText = extractTextFromPDF(fileUrl);
    } else {
      Logger.log(`Extracting text from Google Doc: ${fileUrl}`);
      extractedText = extractTextFromGoogleDoc(fileUrl);
    }
  } catch (e) {
    Logger.log(`Error during text extraction: ${e.message}`);
    return `Error: Unable to access the file. Please ensure the file is publicly accessible. (${e.message})`;
  }

  Logger.log(`Extracted text length: ${extractedText.length}`);

  let learningObjectives;
  try {
    Logger.log(`Processing extracted text to get Learning Objectives...`);
    learningObjectives = processChunks(extractedText, fewShotExamples, newQuery, STOP_SIGNAL);
    Logger.log(`Extracted Learning Objectives count: ${learningObjectives.length}`);
  } catch (e) {
    Logger.log(`Error during AI parsing: ${e.message}`);
    return `Error: Unable to process the text. (${e.message})`;
  }

  try {
    Logger.log(`Creating sheet for institution: ${institution}`);
    const sheetName = createSheetForInstitution(institution);
    Logger.log(`Setting 500 row range to stop from exceeding workbook cell limit`);
    setSheetSize(sheetName, 500, 26); // May need to call this again at the end of processing
    Logger.log(`Storing Learning Objectives in sheet: ${sheetName}`);
    storeDataInSheet(sheetName, learningObjectives);
    Logger.log(`Calling webGenComparison with sheet name: ${sheetName}`);
    webGenComparison(sheetName);
    Logger.log(`webGenComparison completed for sheet: ${sheetName}`);
    return sheetName;  // Return the new sheet name
  } catch (e) {
    Logger.log(`Error during sheet operations: ${e.message}`);
    return `Error: Unable to create or store data in the spreadsheet. (${e.message})`;
  }

}

// Create new sheet for institution
function createSheetForInstitution(institution) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = spreadsheet.getSheetByName('Embeddings');

  if (!templateSheet) {
    throw new Error('Template sheet "Embeddings" not found.');
  }

  const newSheetName = `${institution} Embeddings`;
  let newSheet = spreadsheet.getSheetByName(newSheetName);

  if (newSheet) {
    throw new Error(`Sheet with name "${newSheetName}" already exists. Please use a unique institution name.`);
  }

  newSheet = templateSheet.copyTo(spreadsheet).setName(newSheetName);
  return newSheetName;
}

// Store Learning Objectives in the sheet
function storeDataInSheet(sheetName, learningObjectives) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" not found.`);
  }

  const data = learningObjectives.map((objective) => [objective]);
  sheet.getRange(2, 1, data.length, 1).setValues(data); // start from row 2, column 1 (A2)
}

// Extract the ID from a Google Docs or Drive URL
function getIdFromUrl(url) {
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

// Make API request to OpenAI
function makeApiRequest(messages) {
  const payload = {
    model: 'gpt-4o-mini',
    messages: messages,
    temperature: 0.25 // 20% error rate at 0.7, trying 0.5 for subsequent tests
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + OPENAI_API_KEY
    },
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
  const jsonResponse = JSON.parse(response.getContentText());
  return jsonResponse.choices[0].message.content.trim();
}

// Process text in chunks
function processChunks(text, fewShotExamples, newQuery, stopSignal) {
  const chunkSize = 4000;
  const chunks = [];
  const allResults = [];
  let ongoingObjective = "";

  for (let i = 0; i < text.length; i += chunkSize) {
    chunks.push(text.slice(i, i + chunkSize));
  }

  for (const chunk of chunks) {
    const messages = [
      initialSystemMessage,
      ...fewShotExamples,
      {
        role: 'user',
        content: 'Please extract learning objectives from the following text. Follow the same instructions and formatting as in the examples.\n<user_input>' + chunk
      },
      // newQuery
    ];

    try {
      const responseContent = makeApiRequest(messages);

      if (responseContent === "NULL") {
        continue;
      } else if (responseContent === stopSignal) {
        break;
      } else {
        const objectives = responseContent.split('|').map(o => o.trim());

        for (const objective of objectives) {
          if (objective.endsWith("CONCATNEXT")) {
            ongoingObjective += objective.replace("CONCATNEXT", "").trim() + " ";
          } else {
            allResults.push(ongoingObjective + objective);
            ongoingObjective = "";
          }
        }
      }
    } catch (error) {
      console.error('Error fetching completion for chunk:', error);
    }
  }

  return allResults;
}

// Main function to be executed
function main() {
  const formData = {}; // your formData
  const result = processForm(formData);
  console.log('Process Result:', result);
}

function fetchSheetData(sheetName) {
  Logger.log(`Fetching data from sheet: ${sheetName}`);
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    const errorMessage = `Sheet ${sheetName} not found.`;
    Logger.log(errorMessage);
    throw new Error(errorMessage);
  }

  Logger.log(`Sheet ${sheetName} found. Starting data fetching...`);

  // Get Objectives
  const objectives = sheet.getRange('A2:A').getValues().flat().filter(val => val);
  Logger.log(`Objectives fetched: ${objectives.length} items`);
  const objectiveCount = sheet.getRange('R64').getValue();
  Logger.log(`Objective Count: ${objectiveCount}`);
  
  // Matched vs Unmatched Counts
  const matchCounts = [
    sheet.getRange('S3').getValue(),
    sheet.getRange('S4').getValue()
  ];
  Logger.log(`Match Counts: ${JSON.stringify(matchCounts)}`);
  
  // Match Quality Data
  const qualityLabels = sheet.getRange('R26:R28').getValues().flat();
  Logger.log(`Quality Labels: ${JSON.stringify(qualityLabels)}`);
  const qualityValues = sheet.getRange('S26:S28').getValues().flat();
  Logger.log(`Quality Values: ${JSON.stringify(qualityValues)}`);
  
  // Excellent Matches
  const excellentMatches = sheet.getRange('U26:U').getValues().flat().filter(val => val);
  Logger.log(`Excellent Matches: ${excellentMatches.length} items`);
  
  // Top 5 Most-Aligned Program Objectives
  const topAligned = sheet.getRange('S54:S58').getValues().flat().filter(val => val);
  Logger.log(`Top Aligned Objectives: ${topAligned.length} items`);
  
  // MEPO Areas for Improvement
  const improvementCount = sheet.getRange('S61').getValue();
  Logger.log(`Improvement Count: ${improvementCount}`);
  const improvements = sheet.getRange('S62:S').getValues().flat().filter(val => val);
  Logger.log(`Improvements: ${improvements.length} items`);

  // New add: AAMC Bar Chart Matches
  const chartData = getChartData(sheetName);
  Logger.log(`Bar Chart Data fetched: ${Object.keys(chartData).length} charts in data`)

  Logger.log(`Data fetching completed for sheet: ${sheetName}`);

  return {
    objectives,
    objectiveCount,
    matchCounts,
    qualityLabels,
    qualityValues,
    excellentMatches,
    topAligned,
    improvementCount,
    improvements,
    chartData
  };
}

function getChartData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const columnCValues = sheet.getRange('C2:C43').getValues().flat();
  const aVals = sheet.getRange("A2:A").getValues(); // Fetch all values in column A
  const aLen = aVals.filter(String).length; // Determine populated length (data is always contiguous - if that changes, need another method)
  const lastRow = aLen + 1; // Add row 1 to get row number
  Logger.log(`Got last row value by checking Col A length: ${lastRow}`)
  
  const rangeToCheck = sheet.getRange(`F2:O${lastRow}`).getValues();
  const columnAValues = sheet.getRange(`A2:A${lastRow}`).getValues().flat();

  const result = {
    charts: [
      { name: "Patient Care", labels: columnCValues.slice(0, 7), values: [], tooltips: [] },
      { name: "Patient Care (cont.)", labels: columnCValues.slice(7, 13), values: [], tooltips: [] },
      { name: "Medical Knowledge", labels: columnCValues.slice(13, 15), values: [], tooltips: [] },
      { name: "Systems-Based Practice", labels: columnCValues.slice(15, 21), values: [], tooltips: [] },
      { name: "Practice-Based Learning and Improvement", labels: columnCValues.slice(21, 26), values: [], tooltips: [] },
      { name: "Professionalism", labels: columnCValues.slice(26, 32), values: [], tooltips: [] },
      { name: "Professionalism (cont.)", labels: columnCValues.slice(32, 38), values: [], tooltips: [] },
      { name: "Interpersonal and Communication Skills", labels: columnCValues.slice(38, 43), values: [], tooltips: [] }
    ]
  };

  // Create the chart labels for each chart
  const chartLabels = [
    columnCValues.slice(0, 7),    // Patient Care
    columnCValues.slice(7, 13),   // Patient Care (cont.)
    columnCValues.slice(13, 15),  // Medical Knowledge
    columnCValues.slice(15, 21),  // Systems-Based Practice
    columnCValues.slice(21, 26),  // Practice-Based Learning and Improvement
    columnCValues.slice(26, 32),  // Professionalism
    columnCValues.slice(32, 38),  // Professionalism (cont.)
    columnCValues.slice(38, 43)   // Interpersonal and Communication Skills
  ];

  const stripCosineSimilarity = (str) => {
    // This regex matches strings that have text followed by a numeric value in parentheses and splits the string into 2 parts based on whitespace+"(" as a delimiter
    const match = str.match(/^(.+?)\s*\(\d+\.\d+\)$/);
    return match ? match[1].trim() : str.trim();
  }

  // Iterate through each chart and labels to find occurrences and prepare data
  chartLabels.forEach((labels, chartIndex) => {
    labels.forEach(label => {
      let count = 0;
      let tooltips = [];
      rangeToCheck.forEach((row, rowIndex) => {
        row.forEach(cell => {
          const strippedCell = stripCosineSimilarity(cell); // Use internal function to strip the appended cosine similarity score from the rangeToCheck cells
          if (strippedCell === label) {
            count++;
            tooltips.push(columnAValues[rowIndex]);
          }
        });
      });
      result.charts[chartIndex].values.push(count);
      result.charts[chartIndex].tooltips.push(tooltips.join('\n'));
    });
  });

  Logger.log('Generated chart data:', JSON.stringify(result));
  return result;
}

// Added 10.15.24 to ensure that new copies of Embeddings for each generation have range capped at 500 rows (13,000 cells)
// Current Embeddings sheet range: 520,000 cells (causing overload of spreadsheet max cell limit by generation #~15 - this was what was occurring on tests, at least)
function setSheetSize(sheetName, numRows, numCols) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const currentRows = sheet.getMaxRows();

  // Resize by row
  if (numRows < currentRows) {
    sheet.deleteRows(numRows + 1, currentRows - numRows);
  } else if (numRows > currentRows) {
    sheet.insertRowsAfter(currentRows, numRows - currentRows);
  }

  // Resize by column
  const currentColumns = sheet.getMaxColumns();
  if (numCols < currentColumns) {
    sheet.deleteColumns(numCols + 1, currentColumns - numCols);
  } else if (numCols > currentColumns) {
    sheet.insertColumnsAfter(currentColumns, numCols - currentColumns);
  }
}

// 10.15.24 addition - The spreadsheet had too many cells to function properly - proceeding to a manual back-end cleanup of existing sheets
// Even with setSheetSize, another ~38 generations would hit the limit again
// This function reduces the range of existing sheets made prior to the introduction of setSheetSize as part of the ProcessData flow from the WebApp
// It will be run manually once then commented out to ensure the spreadsheet stays beneath the 10million or 5million cell limit for Wellspring
function reduceTotalCellsInSpreadsheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    let sheetName = sheets[i].getName();
    setSheetSize(sheetName, 500, 26);
  }
}