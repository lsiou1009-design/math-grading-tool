/**
 * Math Marking System - Backend Code
 * Features: Chunking Support, Strict 1M/1A Grading, Blank Handling, Traditional Chinese
 */

// ==========================================
// CONFIGURATION
// ==========================================
const APP_NAME = "Math Marking System";
const DRIVE_FOLDER_NAME = "Math_Marking_System_Uploads";
const SHEET_NAME = "Math_Scores";
const DEFAULT_BOT = "GPT-5.1"; // Optimized for GPT-5.1 as requested

// ==========================================
// WEB APP SERVING
// ==========================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// INITIAL SETUP
// ==========================================
function setup() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  let folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
  }
  
  const files = DriveApp.getFilesByName(SHEET_NAME);
  let ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(SHEET_NAME);
  }
  
  let sheet = ss.getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Student ID", "File URL", "Score", "Comments", "Graded By"]);
  }
  
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('FOLDER_ID', folder.getId());
  scriptProperties.setProperty('SHEET_ID', ss.getId());
  
  return "Setup Complete! Folder ID: " + folder.getId();
}

// ==========================================
// DRIVE & FILE HANDLING
// ==========================================
function getOrCreateFolder() {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty('FOLDER_ID');
  
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {}
  }
  
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) {
    const folder = folders.next();
    props.setProperty('FOLDER_ID', folder.getId());
    return folder;
  }
  
  const folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
  props.setProperty('FOLDER_ID', folder.getId());
  return folder;
}

function uploadFile(data) {
  try {
    const folder = getOrCreateFolder();
    const blob = Utilities.newBlob(Utilities.base64Decode(data.data), data.mimeType, data.fileName);
    const file = folder.createFile(blob);
    
    return {
      success: true,
      fileId: file.getId(),
      url: file.getUrl(),
      name: file.getName()
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Gets list of Files grouped by Student Paper
 * Groups Generated Reports and CSVs under their original file
 */
function getDriveFiles() {
  const folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
  if (!folderId) return [];
  
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const fileMap = {};

  // 1. First pass: Collect all files
  const allFiles = [];
  while (files.hasNext()) {
    const file = files.next();
    allFiles.push({
      id: file.getId(),
      name: file.getName(),
      url: file.getUrl(),
      mimeType: file.getMimeType(),
      created: file.getDateCreated() // Keep as object for sorting
    });
  }

  // 2. Identify Parents (Uploaded Student PDFs) and Children (Reports/CSVs)
  allFiles.sort((a, b) => b.created - a.created); // Newest first

  allFiles.forEach(file => {
    // Check if it's a generated file
    const isReport = file.name.includes("_Report_") && file.mimeType === MimeType.PDF;
    const isCsv = file.name.includes("_Grades_") && file.mimeType === MimeType.CSV;

    if (isReport || isCsv) {
      // It's a child. Try to find the parent name prefix.
      const separator = isReport ? "_Report_" : "_Grades_";
      const baseName = file.name.split(separator)[0];
      
      if (!fileMap[baseName]) fileMap[baseName] = { children: [] };
      fileMap[baseName].children.push({
        ...file,
        type: isReport ? 'PDF Report' : 'CSV Grades',
        displayDate: formatDate(file.created)
      });
    } else {
      // It's likely a parent (Student Upload)
      if (!fileMap[file.name]) fileMap[file.name] = { children: [] };
      fileMap[file.name].parent = {
        ...file,
        displayDate: formatDate(file.created)
      };
    }
  });

  // 3. Convert Map to List
  const result = [];
  Object.keys(fileMap).forEach(key => {
    const item = fileMap[key];
    if (item.parent) {
      result.push({
        ...item.parent,
        generatedFiles: item.children
      });
    }
  });

  return result;
}

function formatDate(dateObj) {
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
}

function getFileContent(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return {
      success: true,
      data: Utilities.base64Encode(blob.getBytes()),
      name: file.getName(),
      mimeType: file.getMimeType()
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// GRADING LOGIC (POE API)
// ==========================================
function saveGrade(studentId, score, comment) {
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheets()[0];
  
  sheet.appendRow([new Date(), studentId, "", score, comment, "Teacher"]);
  return "Saved successfully!";
}

function callPoeAPI(studentImages, solutionImages, studentIndex, modelName) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('POE_API_KEY');
  if (!apiKey) {
    return { error: "API Key not found. Please add POE_API_KEY to Script Properties." };
  }
  const apiUrl = "https://api.poe.com/v1/chat/completions";
  
  // --- STRICT MARKING SCHEME PROMPT ---
  const systemPrompt = `
    You are a STRICT Math Teacher.
    
    **TASK:** Compare the Student's Work (chunk) against the Solution Key (chunk).
    
    **RULE 1: IDENTIFY QUESTIONS STRICTLY**
    - You must ONLY grade questions that are present in the provided SOLUTION KEY images.
    - Do NOT hallucinate questions. If a question is not in the Solution Key, do not output it.
    - ID Format: Use "Q1", "Q2", "Q3a", "Q3b" exactly as seen in the Solution.

    **RULE 2: HANDLING BLANKS (Missing Work)**
    - If a question EXISTS in the Solution Key but the student has left the space BLANK or Skipped it:
      - Score: "0/Total"
      - Comment: "未作答 (Blank)"
    - If the student work for a question is PARTIALLY cut off by the image chunk, try to grade what is visible. If completely invisible, assume Blank (0).

    **RULE 3: STRICT 1M / 1A GRADING**
    - **M Mark (Method):** 1 mark if method is correct. 0 if missing/wrong.
    - **A Mark (Answer):** 1 mark if FINAL ANSWER matches EXACTLY.
    - **NEGATIVE LOGIC:** If Answer != Solution, A mark is 0.
    - **0 Score:** If Method is wrong, Answer is 0.

    **OUTPUT FORMAT:**
    - Language: **Traditional Chinese (繁體中文)** ONLY.
    - Format: Valid JSON ONLY.
    
    **JSON STRUCTURE:**
    {
      "student_name": "Name",
      "total_score": "ignored",
      "overall_comment": "Summary.",
      "questions": [
        { "id": "Q1", "score": "2/2", "comment": "M1 A1 (全對)" },
        { "id": "Q2", "score": "0/3", "comment": "未作答 (Blank)" }
      ]
    }
  `;
  
  const userContent = [
    { "type": "text", "text": `Grade this exam chunk (Student ${studentIndex}). Detect blanks strictly based on Solution Key.` }
  ];
  
  // 1. Add Solution Images (Limit to 5 to save context window)
  if (solutionImages && solutionImages.length > 0) {
    solutionImages.slice(0, 5).forEach(img => {
      userContent.push({
        "type": "image_url",
        "image_url": { "url": `data:image/jpeg;base64,${img}` }
      });
    });
  }
  
  // 2. Add Student Images (Client handles chunking)
  if (Array.isArray(studentImages)) {
    studentImages.forEach(img => {
      userContent.push({
        "type": "image_url",
        "image_url": { "url": `data:image/jpeg;base64,${img}` }
      });
    });
  }
  
  const payload = {
    "model": modelName || DEFAULT_BOT,
    "messages": [
      { "role": "system", "content": systemPrompt },
      { "role": "user", "content": userContent }
    ],
    // Low temp for strict adherence to rules
    "temperature": 0.1 
  };
  
  const options = {
    "method": "post",
    "contentType": "application/json",
    "headers": { "Authorization": `Bearer ${apiKey}` },
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };
  
  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode !== 200) {
      return { error: `Poe API Error (${responseCode}): ${responseText}` };
    }
    
    let json;
    try {
      json = JSON.parse(responseText);
    } catch (e) {
      return { error: "Invalid JSON response from Poe." };
    }
    
    if (!json.choices || !json.choices[0].message) {
      return { error: "No content generated." };
    }
    
    const textResponse = json.choices[0].message.content;
    const jsonMatch = textResponse.match(/\{[\s\S]*\}/);
    
    if (!jsonMatch) {
      return { error: "No JSON found in response: " + textResponse };
    }
    
    let gradeData;
    try {
      gradeData = JSON.parse(jsonMatch[0]);
    } catch (e) {
      return { error: "JSON Parse Failed: " + e.toString() };
    }

    // --- MATH CORRECTION LOGIC (Re-calculate sum for this chunk) ---
    if (gradeData.questions && Array.isArray(gradeData.questions)) {
      let calculatedObtained = 0;
      let calculatedTotal = 0;
      let hasDenominator = false;

      gradeData.questions.forEach(q => {
        if (q.score) {
          const parts = q.score.toString().split('/');
          const obtained = parseFloat(parts[0]);
          if (!isNaN(obtained)) {
            calculatedObtained += obtained;
          }
          if (parts.length > 1) {
            const possible = parseFloat(parts[1]);
            if (!isNaN(possible)) {
              calculatedTotal += possible;
              hasDenominator = true;
            }
          }
        }
      });

      if (hasDenominator && calculatedTotal > 0) {
        gradeData.total_score = `${calculatedObtained}/${calculatedTotal}`;
      } else {
        gradeData.total_score = `${calculatedObtained}`;
      }
    }

    return gradeData;
    
  } catch (e) {
    console.error("Critical Error in callPoeAPI:", e);
    return { error: "Connection Failed. Details: " + e.toString() };
  }
}

// ==========================================
// REPORT GENERATION (PDF & CSV)
// ==========================================
function createPdfReport(gradingData, sourceFileName) {
  let html = `
    <html>
      <head>
        <style>
          body { font-family: 'Microsoft JhengHei', sans-serif; padding: 40px; line-height: 1.6; }
          .student-section { margin-bottom: 40px; page-break-inside: avoid; border-bottom: 1px dashed #ccc; padding-bottom: 20px; }
          .header { font-size: 1.2em; font-weight: bold; margin-bottom: 10px; }
          .score { font-weight: bold; color: #d93025; }
          .label { font-weight: bold; }
          .question-item { margin-left: 20px; }
        </style>
      </head>
      <body>
  `;
  
  gradingData.forEach((item, index) => {
    const safeName = escapeHtml(item.student_name || "Student " + (index + 1));
    const safeScore = escapeHtml(item.total_score);
    const safeComment = escapeHtml(item.overall_comment);
    
    let qHtml = "";
    if (item.questions && item.questions.length > 0) {
      item.questions.forEach(q => {
        const safeId = escapeHtml(q.id);
        const safeQScore = escapeHtml(q.score);
        const safeQComment = escapeHtml(q.comment);
        qHtml += `<div class="question-item"><strong>${safeId}:</strong> (分數: ${safeQScore}) (評語: ${safeQComment})</div>`;
      });
    } else {
      qHtml = "<div class='question-item'>No specific questions found.</div>";
    }
    html += `
      <div class="student-section">
        <div class="header">學生 (${index + 1}): ${safeName}</div>
        <div><span class="label">分數:</span> <span class="score">${safeScore}</span></div>
        <div><span class="label">整體評語:</span> ${safeComment}</div>
        <div><span class="label">細項評語:</span></div>
        ${qHtml}
      </div>
    `;
  });
  
  html += `</body></html>`;
  
  const blob = HtmlService.createHtmlOutput(html).getAs(MimeType.PDF);
  
  // naming convention: [OriginalName]_Report_[Timestamp].pdf
  const cleanSourceName = sourceFileName ? sourceFileName.replace(/\.pdf$/i, "") : "Exam";
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmm");
  blob.setName(`${cleanSourceName}_Report_${timestamp}.pdf`);
  
  const folder = getOrCreateFolder();
  const file = folder.createFile(blob);
  
  return file.getUrl();
}

function saveCsvReport(csvContent, sourceFileName) {
  try {
    const folder = getOrCreateFolder();
    const cleanSourceName = sourceFileName ? sourceFileName.replace(/\.pdf$/i, "") : "Exam";
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmm");
    const fileName = `${cleanSourceName}_Grades_${timestamp}.csv`;
    
    const blob = Utilities.newBlob(csvContent, MimeType.CSV, fileName);
    const file = folder.createFile(blob);
    
    return { success: true, url: file.getUrl() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// UTILITIES
// ==========================================
function escapeHtml(text) {
  if (!text) return "";
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
