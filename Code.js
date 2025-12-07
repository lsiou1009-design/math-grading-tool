/**
 * Math Marking System - Backend Code
 * Features: Chunking Support, Strict 1M/1A Grading, Blank Handling, Traditional Chinese
 * Update: Fixed Dashboard Empty State & Error Handling
 */

// ==========================================
// CONFIGURATION
// ==========================================
const APP_NAME = "Math Marking System";
const DRIVE_FOLDER_NAME = "Math_Marking_System_Uploads";
const SHEET_NAME = "Math_Scores";
const DEFAULT_BOT = "GPT-5.1"; 

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
  
  // Verify if folder actually exists (in case it was deleted manually)
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      // Folder invalid/deleted, fall through to create/search
    }
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
 */
function getDriveFiles() {
  // Removed try-catch to allow errors to propagate to the frontend for debugging
  const folder = getOrCreateFolder(); 
  const files = folder.getFiles();
  const fileMap = {};
  const allFiles = [];

  // 1. First pass: Collect all files
  while (files.hasNext()) {
    const file = files.next();
    allFiles.push({
      id: file.getId(),
      name: file.getName(),
      url: file.getUrl(),
      mimeType: file.getMimeType(),
      created: file.getDateCreated()
    });
  }

  // 2. Sort newest first
  allFiles.sort((a, b) => b.created - a.created); 

  // 3. Group Parents and Children
  allFiles.forEach(file => {
    // Check if it's a generated report or CSV
    const isReport = file.name.includes("_Report_") && file.mimeType === MimeType.PDF;
    const isCsv = file.name.includes("_Grades_") && file.mimeType === MimeType.CSV;
    
    let baseName;

    if (isReport || isCsv) {
      // Child: Extract base name (e.g., "Exam_Report_..." -> "Exam")
      const separator = isReport ? "_Report_" : "_Grades_";
      baseName = file.name.split(separator)[0];
      
      if (!fileMap[baseName]) fileMap[baseName] = { children: [] };
      
      fileMap[baseName].children.push({
        ...file,
        type: isReport ? 'PDF Report' : 'CSV Grades',
        displayDate: formatDate(file.created)
      });
    } else {
      // Parent: Remove extension to match Child key (e.g., "Exam.pdf" -> "Exam")
      baseName = file.name.replace(/\.[^/.]+$/, ""); // Strip extension
      
      if (!fileMap[baseName]) fileMap[baseName] = { children: [] };
      
      // Only set parent if not already set
      if (!fileMap[baseName].parent) {
        fileMap[baseName].parent = {
          ...file,
          displayDate: formatDate(file.created)
        };
      }
    }
  });

  // 4. Convert Map to List (Include Orphans)
  const result = [];
  Object.keys(fileMap).forEach(key => {
    const item = fileMap[key];
    
    if (item.parent) {
      // Normal case: Student PDF exists
      result.push({
        ...item.parent,
        generatedFiles: item.children
      });
    } else if (item.children.length > 0) {
      // Orphan case: Student PDF deleted, but reports exist. Show them anyway.
      // Use the newest report's date and ID as a fallback to ensure it renders
      const newestChild = item.children[0];
      result.push({
        id: newestChild.id, // Fallback ID
        name: key + " [Source File Missing]",
        url: "#",
        mimeType: "application/pdf", // Fake mime to ensure render
        displayDate: newestChild.displayDate,
        generatedFiles: item.children,
        isOrphan: true // Flag for UI
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
    
    **TASK:** Compare the Student's Work (chunk) against the Solution Key.
    
    **RULE 0: NON-CHRONOLOGICAL ORDER (CRITICAL)**
    - The Student's Work may be in **ANY ORDER** (e.g., Q5 may appear before Q1).
    - The Student's Page Count may NOT match the Solution Key Page Count.
    - You must **SEARCH the entire provided Solution Key** to find the specific Question ID that corresponds to the student's work.
    - Do NOT assume the first page of Student Work matches the first page of the Solution.

    **RULE 1: SEPARATE ALL SUB-QUESTIONS**
    - You MUST output every sub-question as a separate JSON entry (e.g., "Q1a", "Q1b", "Q5a(i)").
    - **Do NOT group them.**
    
    **RULE 2: IDENTIFY QUESTIONS STRICTLY**
    - Only grade questions that appear in the provided Solution Key.
    - If you cannot find the question in the Solution Key images, ignore it (do not hallucinate).

    **RULE 3: HANDLING BLANKS**
    - If a question EXISTS in the Solution Key but is blank/skipped by student:
      - Score: "0/Total"
      - Comment: "未作答 (Blank)"
    
    **RULE 4: STRICT 1M / 1A GRADING**
    - **M Mark (Method):** 1 mark if method is correct.
    - **A Mark (Answer):** 1 mark if FINAL ANSWER matches EXACTLY.
    - **NEGATIVE LOGIC:** If Answer != Solution, A mark is 0.

    **OUTPUT FORMAT:**
    - Language: **Traditional Chinese (繁體中文)** ONLY.
    - Format: Valid JSON ONLY.
    
    **JSON STRUCTURE:**
    {
      "student_name": "Name",
      "total_score": "ignored",
      "overall_comment": "Summary.",
      "questions": [
        { "id": "Q1a", "score": "2/2", "comment": "1M 1A (全對)" },
        { "id": "Q5", "score": "0/3", "comment": "未作答 (Blank)" }
      ]
    }
  `;
  
  const userContent = [
    { "type": "text", "text": `Grade this exam chunk (Student ${studentIndex}). Search Solution Key for matching questions. Separate sub-questions.` }
  ];
  
  // 1. Add Solution Images 
  if (solutionImages && solutionImages.length > 0) {
    solutionImages.slice(0, 15).forEach(img => {
      userContent.push({
        "type": "image_url",
        "image_url": { "url": `data:image/jpeg;base64,${img}` }
      });
    });
  }
  
  // 2. Add Student Images
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

    // --- MATH CORRECTION LOGIC ---
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

function escapeHtml(text) {
  if (!text) return "";
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function resetSystem() {
  // Delete the saved ID so the script is forced to create a new folder next time
  PropertiesService.getScriptProperties().deleteProperty('FOLDER_ID');
  return "System Reset! The script will create a new folder on the next run.";
}

function checkFolderStatus() {
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty('FOLDER_ID');
  
  if (!folderId) {
    Logger.log("No Folder ID saved. The script is ready to create a new folder on the next run.");
    return;
  }
  
  try {
    const folder = DriveApp.getFolderById(folderId);
    Logger.log("-------------------------------------");
    Logger.log("Folder Name: " + folder.getName());
    Logger.log("Is Trashed:  " + folder.isTrashed()); // <--- LOOK AT THIS
    Logger.log("Folder URL:  " + folder.getUrl());
    Logger.log("-------------------------------------");
  } catch (e) {
    Logger.log("Error: Folder ID exists but the folder is deleted permanently.");
  }
}
